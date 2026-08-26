"""
FFI Invoice Tools — web backend (FastAPI)
=========================================
Wraps the tested invoice engine (engine.py) behind two endpoints:

  POST /api/process   — upload invoice PDFs + options -> zip of outputs
  POST /api/convert   — upload a register .xlsx + supplier GSTIN + period
                        -> GSTR-1 JSON

Design notes
------------
- Every request works in its own temp directory, deleted immediately after
  the response is sent (BackgroundTask). Nothing is retained on the server.
- CORS is restricted to the frontend origin(s) in ALLOWED_ORIGINS.
- A /health endpoint exists so the frontend can wake the free-tier dyno
  and show a "waking up" message during cold start.
"""

import os
import shutil
import tempfile
import zipfile

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from starlette.background import BackgroundTask

import engine

# Comma-separated list of allowed frontend origins, set in the environment on
# Render. Example: "https://yoursite.com,https://www.yoursite.com"
# Falls back to "*" for local testing only.
ALLOWED_ORIGINS = [
    o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "*").split(",")
    if o.strip()
] or ["*"]

MAX_FILES = 60
MAX_TOTAL_BYTES = 60 * 1024 * 1024   # 60 MB per request

app = FastAPI(title="FFI Invoice Tools API", version="1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_methods=["POST", "GET", "OPTIONS"],
    allow_headers=["*"],
)


@app.get("/health")
def health():
    """Lightweight endpoint the frontend pings to wake the dyno."""
    return {"status": "ok"}


def _cleanup(path):
    shutil.rmtree(path, ignore_errors=True)


def _save_uploads(files, dest_dir):
    """Save uploaded files to dest_dir, enforcing limits. Returns paths."""
    if not files:
        raise HTTPException(400, "No files uploaded.")
    if len(files) > MAX_FILES:
        raise HTTPException(400, f"Too many files (max {MAX_FILES}).")
    saved, total = [], 0
    for f in files:
        data = f.file.read()
        total += len(data)
        if total > MAX_TOTAL_BYTES:
            raise HTTPException(400, "Upload too large (60 MB limit).")
        # keep only the basename; never trust client paths
        name = os.path.basename(f.filename or "file.pdf")
        p = os.path.join(dest_dir, name)
        with open(p, "wb") as out:
            out.write(data)
        saved.append(p)
    return saved


@app.post("/api/process")
def process_invoices(
    files: list[UploadFile] = File(...),
    fmt: str = Form(...),
    start: int = Form(...),
    pad: int = Form(3),
    merge: bool = Form(False),
    sig_above: str = Form(""),
    sig_below: str = Form(""),
    sig_width_mm: float = Form(32),
    sig_remove_white: bool = Form(True),
    make_register: bool = Form(False),
    letterhead: UploadFile | None = File(None),
    signature: UploadFile | None = File(None),
):
    """Renumber + (optional) signature/letterhead/merge/register. Returns a
    zip of all outputs."""
    if "{n}" not in fmt:
        raise HTTPException(400, "Format must contain {n} (e.g. FFI/26-27/{n}).")

    work = tempfile.mkdtemp(prefix="ffi_")
    try:
        in_dir = os.path.join(work, "in"); os.makedirs(in_dir)
        out_dir = os.path.join(work, "out"); os.makedirs(out_dir)
        pdf_paths = _save_uploads(files, in_dir)

        lh_path = ""
        if letterhead is not None and letterhead.filename:
            lh_path = os.path.join(work, "letterhead.pdf")
            with open(lh_path, "wb") as f:
                f.write(letterhead.file.read())

        sig_path = ""
        if signature is not None and signature.filename:
            sig_path = os.path.join(work, os.path.basename(signature.filename))
            with open(sig_path, "wb") as f:
                f.write(signature.file.read())

        sources, errors = engine.scan_files(pdf_paths)
        if errors:
            raise HTTPException(422, "Could not read some files:\n"
                                + "\n".join(errors[:20]))
        engine.assign_numbers(sources, fmt, start, pad)

        try:
            engine.process(
                sources, out_dir,
                sig_above=sig_above, sig_below=sig_below,
                sig_image=sig_path, sig_width_mm=sig_width_mm,
                sig_remove_white=sig_remove_white,
                merge=merge, fmt=fmt, pad=pad,
                letterhead=lh_path, write_excel=make_register,
                write_gstr1=False, log=lambda s: None)
        except ValueError as e:
            raise HTTPException(422, str(e))

        # zip everything in out_dir
        zip_path = os.path.join(work, "invoices_output.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for name in sorted(os.listdir(out_dir)):
                z.write(os.path.join(out_dir, name), name)

        return FileResponse(
            zip_path, media_type="application/zip",
            filename="invoices_output.zip",
            background=BackgroundTask(_cleanup, work))
    except HTTPException:
        _cleanup(work)
        raise
    except Exception:
        _cleanup(work)
        raise


@app.post("/api/convert")
def convert_register(
    file: UploadFile = File(...),
    supplier_gstin: str = Form(...),
    period: str = Form(""),
):
    """Register .xlsx -> GSTR-1 JSON."""
    work = tempfile.mkdtemp(prefix="ffi_conv_")
    try:
        xlsx_path = os.path.join(work, os.path.basename(file.filename or "reg.xlsx"))
        with open(xlsx_path, "wb") as f:
            f.write(file.file.read())
        out_path = os.path.join(work, "GSTR1.json")
        try:
            out_path, warnings = engine.convert_register_to_gstr1(
                xlsx_path, out_path, supplier_gstin, period)
        except ValueError as e:
            raise HTTPException(422, str(e))

        # Return JSON content + warnings in a small wrapper so the frontend
        # can show warnings and still offer the file for download.
        with open(out_path, encoding="utf-8") as f:
            content = f.read()
        _cleanup(work)
        return JSONResponse({"json": content, "warnings": warnings})
    except HTTPException:
        _cleanup(work)
        raise
    except Exception:
        _cleanup(work)
        raise
