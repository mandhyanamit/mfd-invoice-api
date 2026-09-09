#!/usr/bin/env python3
"""
FFI Invoice Renumbering Tool
============================
Renumbers AMC brokerage invoices (CAMS consolidated PDFs and KFin/individual
PDFs) with a user-supplied invoice number series, assigned in alphabetical
order of AMC name across all uploaded files.

- CAMS consolidated PDFs (multiple invoices, one per page) are rebuilt as a
  single consolidated PDF with pages reordered alphabetically by AMC.
- Individual PDFs (one invoice per file) are saved as renumbered copies,
  filename prefixed with the assigned serial.
- Only the invoice number is changed. Everything else is untouched.

Requires: pip install PyMuPDF openpyxl Pillow numpy
          (tkinter is bundled with standard Python on Windows)

Build .exe:
    pip install pyinstaller
    pyinstaller --onefile --windowed --name FFI_Invoice_Renumber invoice_renumber_gui.py
"""

import json
import os
import re
import sys
import traceback
from dataclasses import dataclass, field

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None  # handled at startup with a visible error message

try:
    from PIL import Image
except ImportError:
    Image = None  # signature-image feature disabled with a clear message

try:
    import numpy as np
except ImportError:
    np = None

# ----------------------------------------------------------------------------
# ENGINE
# ----------------------------------------------------------------------------

LABEL_RE = re.compile(r'^(Invoice\s*No\.?|Inv\s*serial\s*No\.?)\s*:?\s*$', re.I)
AMC_RE = re.compile(r'Mutual Fund\s*$', re.I)
OWN_NAME_EXCLUDE = ("friends financial",)          # never treat as AMC
DESC_EXCLUDE_PREFIXES = ("sale", "of ", "for ")    # description-line fragments

FONT_NAME = "helv"       # built-in Helvetica; matches Arial/Helvetica 9-9.6pt sources
FONT_NAME_BOLD = "hebo"  # built-in Helvetica-Bold

# --- Signature-block removal (always-on) --------------------------------
# CAMS: clear the centre-column content BELOW this anchor line (removes the
#       "For Friends Financial Inc" + "Authorised Signatory" block), while
#       keeping the small vertical version stamp in the far-right margin.
CAMS_SIG_ANCHOR = "If yes, amount of GST payable"
CAMS_MARGIN_KEEP_X = 640      # spans with x0 >= this are the right-margin stamp
# KFin: delete these exact labels wherever they appear.
KFIN_SIG_LABELS = ["Name of the Signatory", "Designation / Status", "Signature"]


@dataclass
class Invoice:
    """One invoice = one labelled page (+ any trailing unlabelled pages)."""
    src_path: str
    pages: list            # page indices in the source doc belonging to this invoice
    amc: str               # AMC name as printed
    old_number: str
    rect: tuple            # bbox of the old number span (x0,y0,x1,y1)
    origin: tuple          # baseline origin (x,y) of the old number span
    size: float            # font size of the old number span
    page_label: int        # 1-based page number of the labelled page (for display)
    new_number: str = ""
    serial: int = 0
    # register fields (extracted for the optional Excel export)
    date: str = ""
    gstin: str = ""
    taxable: float = None
    igst: float = None
    total_value: float = None
    pdf_name: str = ""     # output PDF filename assigned during writing


@dataclass
class SourceFile:
    path: str
    invoices: list = field(default_factory=list)

    @property
    def is_consolidated(self):
        return len(self.invoices) > 1


def _spans(page):
    out = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                if s["text"].strip():
                    out.append(s)
    return out


def _find_invoice_on_page(page):
    """Return (amc, old_number, rect, origin, size) or None if page has no invoice label."""
    ss = _spans(page)

    label = next((s for s in ss if LABEL_RE.match(s["text"].strip())), None)
    if label is None:
        return None

    # Invoice number = nearest span to the right of the label on the same line
    ly = (label["bbox"][1] + label["bbox"][3]) / 2
    cands = [s for s in ss
             if s is not label
             and s["bbox"][2] > label["bbox"][2] - 1
             and abs((s["bbox"][1] + s["bbox"][3]) / 2 - ly) < 3
             and s["text"].strip() not in (":", "")]
    cands.sort(key=lambda s: s["bbox"][0])
    if not cands:
        raise ValueError("Invoice No label found but no number next to it")
    num = cands[0]

    # AMC = topmost span ending in "Mutual Fund" that is not us / not description
    amcs = [s for s in ss
            if AMC_RE.search(s["text"].strip())
            and not any(x in s["text"].lower() for x in OWN_NAME_EXCLUDE)
            and not s["text"].strip().lower().startswith(DESC_EXCLUDE_PREFIXES)]
    amcs.sort(key=lambda s: s["bbox"][1])
    if not amcs:
        raise ValueError("Could not detect AMC name on this page")

    return (amcs[0]["text"].strip(), num["text"].strip(),
            tuple(num["bbox"]), tuple(num["origin"]), num["size"])


GSTIN_RE = re.compile(r'\d{2}[A-Z]{5}\d{4}[A-Z]\d[A-Z][A-Z\d]')
_MONEY = re.compile(r'^\d[\d,]*\.\d+$|^\d+$')


def _extract_register_fields(page):
    """Extract date, AMC GSTIN, taxable value, IGST amount, total invoice
    value from an invoice page. Missing fields come back as '' / None rather
    than raising — the renumbering must not fail if the Excel export can't
    read one number."""
    text = page.get_text()
    words = page.get_text("words")   # (x0,y0,x1,y1,word,block,line,wordno)

    # Date: 'June 08, 2026' (CAMS) or '03/06/2026' (KFin)
    md = (re.search(r'([A-Z][a-z]+ \d{1,2}, \d{4})', text)
          or re.search(r'(\d{2}/\d{2}/\d{4})', text))
    date = md.group(1) if md else ""

    # Recipient (AMC) GSTIN by POSITION, not by hardcoded identity:
    # every invoice prints the SUPPLIER's GSTIN in the top header block and the
    # RECIPIENT's below it. So collect all GSTINs top-to-bottom; the first is
    # the supplier, and the recipient is the first later GSTIN that differs.
    found = []  # (y, gstin)
    for s in _spans(page):
        for m in GSTIN_RE.finditer(s["text"].strip()):
            found.append((s["bbox"][1], m.group()))
    found.sort(key=lambda t: t[0])
    gstins = [g for _, g in found]
    supplier = gstins[0] if gstins else ""
    recip = next((g for g in gstins[1:] if g != supplier), "")
    gstin = recip or (gstins[0] if gstins else "")

    # Taxable + IGST from the numeric summary "Total" row: the money tokens to
    # the right are [taxable, cgst, sgst, igst]; IGST is the last.
    def _num(x):
        try:
            return float(x.replace(",", ""))
        except (ValueError, AttributeError):
            return None
    taxable = igst = None
    for tw in [w for w in words if w[4] == "Total"]:
        y = (tw[1] + tw[3]) / 2
        row = sorted([w for w in words
                      if abs((w[1] + w[3]) / 2 - y) < 4 and w[0] > tw[2] - 1],
                     key=lambda w: w[0])
        money = [w[4] for w in row if _MONEY.match(w[4])]
        if len(money) >= 4:
            taxable, igst = _num(money[0]), _num(money[-1])
            break

    mt = re.search(r'Total [Ii]nvoice [Vv]alue[^\d]*([\d,]+\.\d+)', text)
    total_value = _num(mt.group(1)) if mt else None

    return date, gstin, taxable, igst, total_value


def scan_files(paths):
    """Scan PDFs, return (list[SourceFile], list[error_strings])."""
    sources, errors = [], []
    for path in paths:
        try:
            doc = fitz.open(path)
        except Exception as e:
            errors.append(f"{os.path.basename(path)}: cannot open ({e})")
            continue
        sf = SourceFile(path=path)
        current = None
        for pno in range(len(doc)):
            try:
                found = _find_invoice_on_page(doc[pno])
            except ValueError as e:
                errors.append(f"{os.path.basename(path)} p{pno+1}: {e}")
                found = None
            if found:
                amc, old, rect, origin, size = found
                date, gstin, taxable, igst, total_value = \
                    _extract_register_fields(doc[pno])
                current = Invoice(src_path=path, pages=[pno], amc=amc,
                                  old_number=old, rect=rect, origin=origin,
                                  size=size, page_label=pno + 1,
                                  date=date, gstin=gstin, taxable=taxable,
                                  igst=igst, total_value=total_value)
                sf.invoices.append(current)
            elif current is not None:
                current.pages.append(pno)      # continuation page
            else:
                errors.append(f"{os.path.basename(path)} p{pno+1}: "
                              f"no invoice found on first page")
        doc.close()
        if sf.invoices:
            sources.append(sf)
        elif not any(path in e for e in errors):
            errors.append(f"{os.path.basename(path)}: no invoices detected")
    return sources, errors


def assign_numbers(sources, fmt, start, pad):
    """Sort all invoices alphabetically by AMC (case-insensitive) and assign
    serials. Returns the sorted invoice list."""
    if "{n}" not in fmt:
        raise ValueError('Format must contain {n} — e.g. FFI/2026-27/{n}')
    all_inv = [inv for sf in sources for inv in sf.invoices]
    all_inv.sort(key=lambda i: i.amc.lower())
    n = start
    for inv in all_inv:
        serial = str(n).zfill(pad) if pad > 0 else str(n)
        inv.new_number = fmt.replace("{n}", serial)
        inv.serial = n
        n += 1
    return all_inv


def _sig_regions(page):
    """Return (clear_rects, place_point) for a page's signature block.

    clear_rects : list of fitz.Rect to blank out (the signature block).
    place_point : (x, y) baseline where optional replacement text should go,
                  or None if this page has no recognisable signature block.

    CAMS: everything below the "If yes, amount of GST payable" anchor except
          the far-right vertical version stamp.
    KFin: the three fixed labels.
    """
    ss = _spans(page)
    text = page.get_text()
    clears, place = [], None

    if CAMS_SIG_ANCHOR in text:
        anchor_y = max((s["bbox"][3] for s in ss
                        if CAMS_SIG_ANCHOR in s["text"]), default=None)
        if anchor_y is not None:
            for s in ss:
                if s["bbox"][1] > anchor_y and s["bbox"][0] < CAMS_MARGIN_KEEP_X:
                    clears.append(fitz.Rect(s["bbox"]))
            # place replacement where the block began (first cleared line)
            if clears:
                top = min(clears, key=lambda r: r.y0)
                place = (top.x0, top.y0 + top.height * 0.8)

    kf = [s for s in ss if s["text"].strip() in KFIN_SIG_LABELS]
    if kf:
        for s in kf:
            clears.append(fitz.Rect(s["bbox"]))
        first = min(kf, key=lambda s: s["bbox"][1])
        place = (first["bbox"][0], first["origin"][1])

    return clears, place


def prepare_signature_image(path, remove_white=True, thresh=225):
    """Load a signature image, optionally knock out its white background, and
    trim to the ink bounding box. Returns (png_bytes, (w, h)).
    Raises ValueError with a clear message on any problem."""
    if Image is None or np is None:
        missing = "Pillow" if Image is None else "numpy"
        raise ValueError(f"Signature-image support needs the {missing} "
                         f"library.\nInstall it with:  python -m pip install "
                         f"{missing}")
    import io
    try:
        im = Image.open(path).convert("RGBA")
    except Exception as e:
        raise ValueError(f"Cannot open signature image: {e}")
    arr = np.asarray(im).astype(int)
    if remove_white:
        lum = arr[:, :, :3].mean(axis=2)
        # white -> transparent, ink -> opaque, smooth ramp; keep existing alpha
        ramp = np.clip((thresh - lum) / thresh * 255 * 1.6, 0, 255)
        arr[:, :, 3] = np.minimum(arr[:, :, 3], 255)
        arr[:, :, 3] = np.maximum(np.where(arr[:, :, 3] < 255, ramp,
                                           ramp), 0)
        im = Image.fromarray(arr.astype("uint8"), "RGBA")
    a2 = np.asarray(im)
    mask = a2[:, :, 3] > 10
    if not mask.any():
        raise ValueError("Signature image looks blank after background "
                         "removal — try unchecking 'remove white background'.")
    ys, xs = mask.nonzero()
    im = im.crop((int(xs.min()), int(ys.min()),
                  int(xs.max()) + 1, int(ys.max()) + 1))
    buf = io.BytesIO()
    im.save(buf, "PNG")
    return buf.getvalue(), im.size


def _apply_signature(page, sig_above="", sig_below="", sig_png=None,
                     sig_size=None, sig_width_mm=32):
    """Clear the signature block and compose, in order:
        sig_above (one or more lines)  ->  signature image  ->  sig_below.
    Any part may be empty. Returns True if a signature block was found and
    cleared."""
    clears, place = _sig_regions(page)
    if not clears:
        return False
    for r in clears:
        page.add_redact_annot(r + (-1, -1, 1, 1))
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    if not place or (not sig_above and not sig_below and not sig_png):
        return True

    x, y = place
    lead = 9.5 * 1.6
    cursor = y

    def draw_lines(block):
        nonlocal cursor
        for line in block.split("\n"):
            if line.strip():
                page.insert_text(fitz.Point(x, cursor + 9.5), line,
                                 fontname=FONT_NAME_BOLD, fontsize=9.5,
                                 color=(0, 0, 0))
            cursor += lead

    if sig_above:
        draw_lines(sig_above)
    if sig_png and sig_size:
        img_w = sig_width_mm * 72 / 25.4
        img_h = img_w * (sig_size[1] / sig_size[0])
        cursor += lead * 0.2
        page.insert_image(fitz.Rect(x, cursor, x + img_w, cursor + img_h),
                          stream=sig_png, keep_proportion=True)
        cursor += img_h + lead * 0.5
    if sig_below:
        draw_lines(sig_below)
    return True


def _replace_number(page, inv):
    """White-out the old invoice number and write the new one."""
    page.add_redact_annot(fitz.Rect(*inv.rect) + (-1, -1, 1, 1))
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    page.insert_text(fitz.Point(*inv.origin), inv.new_number,
                     fontname=FONT_NAME, fontsize=inv.size, color=(0, 0, 0))


def safe_name(s, maxlen=40):
    return re.sub(r'[^A-Za-z0-9_-]+', '_', s).strip('_')[:maxlen]


def _page_item_rects(page):
    """All content rectangles on a page (text blocks, drawings, images)."""
    rects = [fitz.Rect(b[:4]) for b in page.get_text("blocks")]
    rects += [d["rect"] for d in page.get_drawings()]
    rects += [fitz.Rect(i["bbox"]) for i in page.get_image_info()]
    return [r for r in rects if not r.is_empty]


def _content_bbox(page):
    """Tight bbox around everything on the page, with a small margin."""
    rects = _page_item_rects(page)
    if not rects:
        return page.rect
    r = fitz.Rect(rects[0])
    for x in rects[1:]:
        r |= x
    r += (-6, -6, 6, 6)
    return r & page.rect


def detect_body_zone(lh_page):
    """Find the clear zone between the letterhead's header graphics and its
    footer. Raises ValueError if the letterhead is too busy to host an
    invoice (e.g. full-page watermark/background)."""
    H, W = lh_page.rect.height, lh_page.rect.width
    top, bottom = 20.0, H - 20.0
    for r in _page_item_rects(lh_page):
        cy = (r.y0 + r.y1) / 2
        if cy < H / 2:
            top = max(top, min(r.y1, H))       # header graphics push body down
        else:
            bottom = min(bottom, max(r.y0, 0)) # footer graphics push body up
    body = fitz.Rect(15, top + 10, W - 15, bottom - 8)
    if body.height < 0.45 * H or body.width < 0.5 * W:
        raise ValueError(
            "This letterhead's graphics leave only "
            f"{body.height / 28.35:.1f} cm of clear vertical space — the "
            "invoice would be unreadably small. Use a letterhead with a "
            "header band and/or footer only (clear middle), or run without "
            "a letterhead.")
    return body


def _emit(pairs, out_path, lh_doc=None, body=None):
    """Write one output PDF from (src_doc, Invoice) pairs, in given order.
    With lh_doc set, every page is rebuilt on the letterhead's page size with
    the invoice content scaled into the clear body zone."""
    new = fitz.open()
    for src, inv in pairs:
        for pidx in inv.pages:
            if lh_doc is None:
                new.insert_pdf(src, from_page=pidx, to_page=pidx)
            else:
                lr = lh_doc[0].rect
                pg = new.new_page(width=lr.width, height=lr.height)
                pg.show_pdf_page(pg.rect, lh_doc, 0)
                pg.show_pdf_page(body, src, pidx,
                                 clip=_content_bbox(src[pidx]))
    new.save(out_path, garbage=3, deflate=True)
    new.close()


def merged_filename(fmt, lo, hi, pad):
    """Build the combined-PDF filename from the user's format string.
    'MKM/2026-27/{n}', 19, 32, pad 3  ->  MKM_Invoices_2026-27_019-032.pdf"""
    parts = [safe_name(p.replace("{n}", "")).strip("_-") for p in fmt.split("/")]
    parts = [p for p in parts if p]
    head = parts[0] if parts else "Invoices"
    rest = "_".join(parts[1:])
    z = (lambda n: str(n).zfill(pad)) if pad > 0 else str
    name = head + "_Invoices" + (f"_{rest}" if rest else "") + f"_{z(lo)}-{z(hi)}"
    return name + ".pdf"


def write_register_xlsx(invoices, path):
    """Write the invoice register .xlsx with the seven requested columns,
    invoices already in serial order. Blank cells for any field that
    couldn't be read (flagged so they're easy to spot)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    headers = ["INVOICE DATE", "INVOICE NUMBER", "NAME OF THE MUTUAL FUND",
               "GSTIN", "TOTAL INVOICE VALUE", "TAXABLE VALUE", "IGST AMOUNT",
               "PDF FILE NAME"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice Register"

    hfill = PatternFill("solid", fgColor="1A7A6E")
    hfont = Font(bold=True, color="FFFFFF", size=11)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    money = '#,##0.00'

    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = hfill; cell.font = hfont; cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)

    sum_total = sum_tax = sum_igst = 0.0
    for r, inv in enumerate(invoices, start=2):
        row = [inv.date, inv.new_number, inv.amc, inv.gstin,
               inv.total_value, inv.taxable, inv.igst, inv.pdf_name]
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = border
            if c in (5, 6, 7):
                cell.number_format = money
                cell.alignment = Alignment(horizontal="right")
                if val is None:
                    cell.value = "MISSING"
                    cell.font = Font(color="C0392B", italic=True)
        sum_total += inv.total_value or 0
        sum_tax += inv.taxable or 0
        sum_igst += inv.igst or 0

    tr = len(invoices) + 2
    tcell = ws.cell(row=tr, column=4, value="TOTAL")
    tcell.font = Font(bold=True); tcell.alignment = Alignment(horizontal="right")
    for c, val in [(5, sum_total), (6, sum_tax), (7, sum_igst)]:
        cell = ws.cell(row=tr, column=c, value=round(val, 2))
        cell.number_format = money; cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="right")
        cell.border = Border(top=Side(style="double"))

    widths = [15, 20, 34, 20, 20, 18, 16, 22]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(invoices)+1}"
    wb.save(path)


# FFI's own GSTIN — the supplier the return is filed for.
OWN_GSTIN = "09ADYPM9750C1ZW"
GSTR1_VERSION = "GST3.2.4"


def _json_date(datestr):
    """Convert an extracted invoice date to GSTR-1 JSON format 'DD-MM-YYYY'.
    Accepts 'June 08, 2026' or '03/06/2026'."""
    from datetime import datetime
    for fmt in ("%B %d, %Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(datestr.strip(), fmt).strftime("%d-%m-%Y")
        except (ValueError, AttributeError):
            pass
    return datestr  # leave as-is if unparseable


def detect_filing_period(invoices):
    """Best-guess filing period 'MMYYYY' from the invoices' dates (the most
    common month). Empty string if none parse."""
    from datetime import datetime
    from collections import Counter
    months = []
    for i in invoices:
        for fmt in ("%B %d, %Y", "%d/%m/%Y"):
            try:
                d = datetime.strptime((i.date or "").strip(), fmt)
                months.append(d.strftime("%m%Y"))
                break
            except (ValueError, AttributeError):
                pass
    return Counter(months).most_common(1)[0][0] if months else ""


def build_gstr1_json(invoices, filing_period, own_gstin=OWN_GSTIN):
    """Build the GSTR-1 offline-tool JSON dict from the batch.

    Invoices are grouped by recipient GSTIN (ctin) exactly as the portal
    expects. Matches the user's filed format, including csamt = iamt.
    """
    inv_sorted = sorted(invoices, key=lambda i: i.serial)

    # group by recipient GSTIN, preserving first-seen order
    groups = {}
    for i in inv_sorted:
        groups.setdefault(i.gstin, []).append(i)

    b2b = []
    for ctin, invs in groups.items():
        inv_list = []
        for i in invs:
            iamt = round(i.igst or 0, 2)
            inv_list.append({
                "inum": i.new_number,
                "idt": _json_date(i.date),
                "val": i.total_value,
                "pos": (i.gstin or "")[:2],
                "rchrg": "N",
                "inv_typ": "R",
                "itms": [{
                    "num": 1801,
                    "itm_det": {
                        "txval": round(i.taxable or 0, 2),
                        "rt": 18,
                        "iamt": iamt,
                        "csamt": iamt,   # matches user's filed format
                    },
                }],
            })
        b2b.append({"ctin": ctin, "inv": inv_list})

    lo, hi = inv_sorted[0].new_number, inv_sorted[-1].new_number
    n = len(inv_sorted)

    # HSN summary — one rolled-up row for the commission service (SAC 9971),
    # matching the portal's format: total taxable value and total IGST.
    hsn_txval = round(sum((i.taxable or 0) for i in inv_sorted), 2)
    hsn_iamt = round(sum((i.igst or 0) for i in inv_sorted), 2)

    return {
        "gstin": own_gstin,
        "fp": filing_period,
        "version": GSTR1_VERSION,
        "hash": "hash",
        "b2b": b2b,
        "hsn": {
            "flag": "N",
            "hsn_b2b": [{
                "num": 1,
                "hsn_sc": "9971",
                "desc": "Financial and related services",
                "uqc": "NA",
                "qty": 0,
                "rt": 18,
                "txval": hsn_txval,
                "iamt": hsn_iamt,
            }],
            "hsn_b2c": [],
        },
        "doc_issue": {
            "doc_det": [{
                "doc_num": 1,
                "doc_typ": "Invoices for outward supply",
                "docs": [{
                    "num": 1, "from": lo, "to": hi,
                    "totnum": n, "cancel": 0, "net_issue": n,
                }],
            }],
        },
    }


def write_gstr1_json(invoices, path, filing_period, own_gstin=OWN_GSTIN):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(build_gstr1_json(invoices, filing_period, own_gstin), f,
                  separators=(",", ":"), ensure_ascii=False)


# --- Register .xlsx -> GSTR-1 JSON converter -----------------------------

class RegisterRow:
    """Lightweight stand-in exposing just the attributes build_gstr1_json
    reads, populated from a register .xlsx row."""
    __slots__ = ("serial", "new_number", "date", "gstin",
                 "total_value", "taxable", "igst", "amc")

    def __init__(self, serial, new_number, date, gstin,
                 total_value, taxable, igst, amc=""):
        self.serial = serial
        self.new_number = new_number
        self.date = date
        self.gstin = gstin
        self.total_value = total_value
        self.taxable = taxable
        self.igst = igst
        self.amc = amc


FULL_GSTIN_RE = re.compile(r'^\d{2}[A-Z]{5}\d{4}[A-Z]\d[A-Z][A-Z\d]$')

# GST state codes (first two digits of a GSTIN), from the GSTR-1 master list.
# Used to validate the Place-of-Supply prefix; the JSON stores the bare code.
GST_STATE_CODES = {
    "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab",
    "04": "Chandigarh", "05": "Uttarakhand", "06": "Haryana", "07": "Delhi",
    "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim",
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur",
    "15": "Mizoram", "16": "Tripura", "17": "Meghalaya", "18": "Assam",
    "19": "West Bengal", "20": "Jharkhand", "21": "Odisha",
    "22": "Chattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
    "25": "Daman & Diu", "26": "Dadra & Nagar Haveli", "27": "Maharashtra",
    "28": "Andhra Pradesh", "29": "Karnataka", "30": "Goa", "31": "Lakshadweep",
    "32": "Kerala", "33": "Tamil Nadu", "34": "Puducherry",
    "35": "Andaman & Nicobar Islands", "36": "Telangana",
    "37": "Andhra Pradesh (New)", "38": "Ladakh", "97": "Other Territory",
}

# Register column order written by write_register_xlsx()
_REG_HEADERS = ["INVOICE DATE", "INVOICE NUMBER", "NAME OF THE MUTUAL FUND",
                "GSTIN", "TOTAL INVOICE VALUE", "TAXABLE VALUE", "IGST AMOUNT"]


def read_register_xlsx(path):
    """Read a register .xlsx (as written by this tool) into RegisterRow
    objects. Returns (rows, warnings). Trusts the sheet values (honouring
    manual edits) but collects non-blocking warnings for odd data.
    Raises ValueError if the file can't be read or headers don't match."""
    from openpyxl import load_workbook
    from datetime import datetime, date as _date
    try:
        wb = load_workbook(path, data_only=True)
    except Exception as e:
        raise ValueError(f"Cannot open the Excel file: {e}")
    ws = wb.active

    headers = [ws.cell(1, c).value for c in range(1, 8)]
    if [str(h).strip().upper() if h else "" for h in headers] != _REG_HEADERS:
        raise ValueError(
            "This doesn't look like a register created by this tool.\n"
            "Expected columns: " + ", ".join(_REG_HEADERS))

    rows, warnings = [], []
    fallback_serial = 0
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, 8)]
        if all(v is None for v in vals):
            continue
        raw_date, inum, amc, gstin, total_v, taxable, igst = vals
        # skip a trailing TOTAL row (col D says TOTAL, no invoice number)
        if (inum is None and isinstance(gstin, str)
                and gstin.strip().upper() == "TOTAL"):
            continue
        if inum is None and gstin is None:
            continue
        fallback_serial += 1

        # serial for ordering: trailing number in the invoice no., else row order
        mnum = re.search(r'(\d+)\s*$', str(inum)) if inum else None
        serial = int(mnum.group(1)) if mnum else fallback_serial

        # date -> string the JSON builder understands
        if isinstance(raw_date, (datetime, _date)):
            date_str = raw_date.strftime("%d/%m/%Y")
        else:
            date_str = str(raw_date).strip() if raw_date else ""

        gstin = (str(gstin).strip() if gstin else "")
        label = inum or f"row {r}"
        if not FULL_GSTIN_RE.match(gstin):
            warnings.append(f"{label}: GSTIN looks invalid ({gstin or 'blank'})")
        elif gstin[:2] not in GST_STATE_CODES:
            warnings.append(
                f"{label}: GSTIN state code '{gstin[:2]}' is not a valid "
                f"Indian state code")
        if _json_date(date_str) == date_str and "-" not in date_str:
            warnings.append(f"{label}: date not understood ({date_str})")

        def _f(x):
            try:
                return float(x)
            except (TypeError, ValueError):
                return None
        tv, tx, ig = _f(total_v), _f(taxable), _f(igst)
        if tx and ig is not None and abs(ig - tx * 0.18) > max(1.0, tx * 0.01):
            warnings.append(
                f"{label}: IGST {ig} is not ~18% of taxable {tx}")

        rows.append(RegisterRow(serial, str(inum).strip() if inum else "",
                                date_str, gstin, tv, tx, ig,
                                str(amc).strip() if amc else ""))
    if not rows:
        raise ValueError("No invoice rows found in the sheet.")
    return rows, warnings


def convert_register_to_gstr1(xlsx_path, out_path, supplier_gstin,
                              filing_period=""):
    """Read a register .xlsx and write a GSTR-1 JSON for the given supplier.
    Returns (out_path, warnings). Raises ValueError on unusable input."""
    sg = (supplier_gstin or "").strip()
    if not FULL_GSTIN_RE.match(sg):
        raise ValueError("Supplier GSTIN is not a valid 15-character GSTIN.")
    if sg[:2] not in GST_STATE_CODES:
        raise ValueError(
            f"Supplier GSTIN state code '{sg[:2]}' is not a valid Indian "
            f"state code.")
    rows, warnings = read_register_xlsx(xlsx_path)
    period = filing_period or detect_filing_period(rows)
    if not re.fullmatch(r"\d{6}", period or ""):
        raise ValueError("Filing period could not be determined; enter it as "
                         "MMYYYY (e.g. 082026).")
    write_gstr1_json(rows, out_path, period, own_gstin=sg)
    return out_path, warnings


def process(sources, out_dir, sig_above="", sig_below="", sig_image="",
            sig_width_mm=32, sig_remove_white=True, fmt="",
            pad=3, letterhead="", write_excel=False, write_gstr1=False,
            gstr1_period="", log=print):
    """Write renumbered PDFs. Returns list of output file paths.

    Each invoice is written as its own PDF, named by the AMC's first word
    (Aditya.pdf, Bandhan.pdf ...; a colliding first word gets _2, _3 ...).

    sig_above / sig_below : text printed above / below the signature image in
                  the (always cleared) signature area. Both empty -> blank.
    sig_image     : optional path to a scanned signature image, placed between
                  the above and below text.
    sig_width_mm  : printed width of the signature image, in millimetres.
    sig_remove_white : knock out the image's white background (for raw scans).
    fmt, pad    : the invoice number format/padding (used for the register and
                  GSTR-1 filenames).
    letterhead  : optional path to a letterhead PDF; every output page is
                  rebuilt on it, invoice scaled into its clear body zone.
                  Raises ValueError before writing anything if unusable.
    write_excel : also write an .xlsx register of all invoices (serial order).
    """
    # Prep the signature image up front — fail before any output is written
    sig_png = sig_size = None
    if sig_image:
        sig_png, sig_size = prepare_signature_image(
            sig_image, remove_white=sig_remove_white)

    # Validate the letterhead up front — fail before any output is written
    lh_doc = body = None
    if letterhead:
        try:
            lh_doc = fitz.open(letterhead)
        except Exception as e:
            raise ValueError(f"Cannot open letterhead PDF: {e}")
        if lh_doc.needs_pass:
            raise ValueError("Letterhead PDF is password-protected.")
        if len(lh_doc) < 1:
            raise ValueError("Letterhead PDF has no pages.")
        body = detect_body_zone(lh_doc[0])   # raises ValueError if too busy

    os.makedirs(out_dir, exist_ok=True)
    written, docs = [], {}

    # Pass 1: renumber + always-on signature removal in the source docs
    for sf in sources:
        doc = fitz.open(sf.path)
        for inv in sf.invoices:
            _replace_number(doc[inv.pages[0]], inv)
            for pidx in inv.pages:
                _apply_signature(doc[pidx], sig_above, sig_below,
                                 sig_png=sig_png, sig_size=sig_size,
                                 sig_width_mm=sig_width_mm)
        docs[sf.path] = doc

    # Pass 2: write outputs — one file per invoice, named by the AMC's first
    # word. The assigned filename is recorded on each invoice so the register
    # can list it.
    all_inv = sorted((inv for sf in sources for inv in sf.invoices),
                     key=lambda i: i.serial)
    used = {}
    for inv in all_inv:
        first_word = safe_name(inv.amc.split()[0]) or "Invoice"
        n = used.get(first_word.lower(), 0) + 1
        used[first_word.lower()] = n
        name = first_word if n == 1 else f"{first_word}_{n}"
        inv.pdf_name = f"{name}.pdf"
        out = os.path.join(out_dir, inv.pdf_name)
        _emit([(docs[inv.src_path], inv)], out, lh_doc, body)
        written.append(out)
        log(f"  written: {inv.pdf_name}")

    if write_excel:
        lo, hi = all_inv[0].serial, all_inv[-1].serial
        xlsx_name = (merged_filename(fmt, lo, hi, pad)[:-4] if fmt
                     else f"Invoices_{lo:03d}-{hi:03d}") + ".xlsx"
        xlsx_path = os.path.join(out_dir, xlsx_name)
        write_register_xlsx(all_inv, xlsx_path)
        written.append(xlsx_path)
