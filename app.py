import os
import shutil
import traceback
from pathlib import Path
from typing import Optional

import requests
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from pydantic import BaseModel, Field

# Apply PyMuPDF (fitz) monkeypatch to handle unsupported colorspaces (e.g., CMYK)
try:
    import fitz
    _orig_save = fitz.Pixmap.save
    _orig_tobytes = fitz.Pixmap.tobytes

    def _safe_pixmap(pix):
        if pix.colorspace:
            cs_name = pix.colorspace.name
            if cs_name and cs_name not in ("DeviceRGB", "DeviceGray", "RGB", "GRAY"):
                return fitz.Pixmap(fitz.csRGB, pix)
        return pix

    def patched_save(self, *args, **kwargs):
        return _orig_save(_safe_pixmap(self), *args, **kwargs)

    def patched_tobytes(self, *args, **kwargs):
        return _orig_tobytes(_safe_pixmap(self), *args, **kwargs)

    fitz.Pixmap.save = patched_save
    fitz.Pixmap.tobytes = patched_tobytes
    print("[✓] PyMuPDF colorspace monkeypatch applied successfully.")
except Exception as patch_err:
    print(f"[!] Warning: Failed to apply colorspace monkeypatch: {patch_err}")

from pdf2docx import Converter

app = FastAPI(title="PDF to Word Converter API")

# ── CORS ────────────────────────────────────────────────────────────────────
# Browser-side callers (the CLM upload flow in app.esigns.io) post JSON from
# a different origin, so without CORS middleware the preflight + POST get
# rejected with the usual "No 'Access-Control-Allow-Origin' header" error.
#
# `CORS_ALLOW_ORIGINS` lets the deployment lock this down to specific
# domains via an env var (comma-separated). The default "*" keeps the
# converter usable from any origin — fine for a stateless, auth-free
# microservice that only acts on URLs the caller already holds.
_allow_origins_env = os.environ.get("CORS_ALLOW_ORIGINS", "*").strip()
if _allow_origins_env == "*":
    allow_origins = ["*"]
else:
    allow_origins = [o.strip() for o in _allow_origins_env.split(",") if o.strip()]

# allow_credentials must be False when allow_origins is "*" — the browser
# will reject the response otherwise. Set credentials only when the env
# explicitly enumerates origins.
allow_credentials = allow_origins != ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=allow_origins,
    allow_credentials=allow_credentials,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
    max_age=600,
)

# Setup directories
UPLOAD_DIR = Path("temp_uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

# Helper to clean up temporary files
def remove_file(filepath: Path):
    try:
        if filepath.exists():
            filepath.unlink()
    except Exception as e:
        print(f"Error deleting temporary file {filepath}: {e}")

@app.get("/", response_class=HTMLResponse)
def get_index():
    # Load index.html file
    index_path = Path("templates/index.html")
    if not index_path.exists():
        raise HTTPException(status_code=504, detail="Frontend template not found")
    return index_path.read_text(encoding="utf-8")

@app.post("/convert")
async def convert_pdf_endpoint(
    background_tasks: BackgroundTasks, 
    file: UploadFile = File(...)
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are supported.")

    # Generate safe unique names
    temp_id = os.urandom(8).hex()
    input_pdf_path = UPLOAD_DIR / f"{temp_id}_{file.filename}"
    output_docx_path = input_pdf_path.with_suffix(".docx")

    cv = None
    try:
        # Save uploaded PDF file
        with input_pdf_path.open("wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Perform PDF to DOCX conversion
        cv = Converter(str(input_pdf_path))
        cv.convert(str(output_docx_path), start=0, end=None)
        cv.close()
        cv = None # Clear reference after successful close

        # Enqueue background task to clean up files after sending download response
        background_tasks.add_task(remove_file, input_pdf_path)
        background_tasks.add_task(remove_file, output_docx_path)

        # Return file response
        return FileResponse(
            path=output_docx_path,
            filename=output_docx_path.name.replace(f"{temp_id}_", ""),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        import traceback
        print("[✗] Exception occurred during conversion:")
        traceback.print_exc()
        
        # Ensure converter is closed to release file lock before cleanup
        if cv:
            try:
                cv.close()
            except Exception:
                pass
        
        # Clean up files in case of error
        remove_file(input_pdf_path)
        remove_file(output_docx_path)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

# ── S3-URL flow ──────────────────────────────────────────────────────────────
# Companion endpoint to /convert that integrates with the eSigns backend's
# existing PRINT_URL/convert-to-docx pattern: the caller hands us a signed
# S3 GET URL for the source PDF and a signed S3 PUT URL for the destination
# DOCX. We never serve the file back over HTTP — we upload directly to S3
# and return a JSON status. Same shape the FE's convertDocFileToDocxAPI
# already speaks: { download_url, upload_url, file_name }.

# DOCX is the standard MIME for Word documents. S3 signed PUT URLs that
# were signed with this Content-Type require it on the request — pass it
# both as a header and rely on the signing side to use the same value.
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# Cap downloads so a malicious / runaway caller can't drain disk.
MAX_PDF_BYTES = 100 * 1024 * 1024  # 100 MB


class ConvertFromUrlRequest(BaseModel):
    download_url: str = Field(
        ...,
        description="Signed S3 GET URL for the source PDF.",
        min_length=8,
    )
    upload_url: str = Field(
        ...,
        description="Signed S3 PUT URL for the converted DOCX.",
        min_length=8,
    )
    file_name: Optional[str] = Field(
        None,
        description="Original PDF filename (used only for logging / temp filenames).",
    )
    upload_content_type: Optional[str] = Field(
        None,
        description=(
            "Override Content-Type sent on the PUT. Leave empty when the "
            "signed URL doesn't constrain the type (default). Set to "
            "'application/pdf' when reusing a URL signed for the source PDF."
        ),
    )


def _stream_download_to(path: Path, url: str) -> int:
    """Stream-download a URL to disk with a hard size cap.

    Returns the number of bytes written. Raises HTTPException on network
    errors or when the response exceeds MAX_PDF_BYTES.
    """
    total = 0
    try:
        with requests.get(url, stream=True, timeout=60) as r:
            if r.status_code >= 400:
                raise HTTPException(
                    status_code=502,
                    detail=f"download_url returned HTTP {r.status_code}",
                )
            content_length = int(r.headers.get("Content-Length") or 0)
            if content_length and content_length > MAX_PDF_BYTES:
                raise HTTPException(
                    status_code=413,
                    detail=f"Source PDF is {content_length} bytes — limit is {MAX_PDF_BYTES}.",
                )
            with path.open("wb") as fh:
                for chunk in r.iter_content(chunk_size=1024 * 64):
                    if not chunk:
                        continue
                    total += len(chunk)
                    if total > MAX_PDF_BYTES:
                        raise HTTPException(
                            status_code=413,
                            detail=f"Source PDF exceeded {MAX_PDF_BYTES} bytes during download.",
                        )
                    fh.write(chunk)
    except requests.RequestException as exc:
        raise HTTPException(
            status_code=502,
            detail=f"Failed to download source PDF: {exc}",
        )
    return total


def _put_docx(path: Path, url: str, content_type: Optional[str] = None) -> None:
    """PUT the converted DOCX to a signed S3 upload URL.

    Critical: read the file fully into memory and pass bytes (not a file
    handle). When `requests.put(url, data=<file handle>)` is used, the
    library streams the body using `Transfer-Encoding: chunked`, which
    S3's v4 presigned PUT URLs reject unless they were explicitly signed
    for chunked payloads (they aren't by default). Browser `fetch()`
    PUTs work because they send the whole body with a `Content-Length`
    and no Transfer-Encoding — matching what S3 expects. Passing bytes
    here makes `requests` behave the same way.

    Content-Type:
      - When `content_type` is provided, send it (use this when the
        signed URL has a strict MIME requirement).
      - When omitted, send nothing — mirrors the browser's bare PUT
        which is what the existing DOC→DOCX print service uses
        successfully against the same kind of S3 presigned URLs.
    """
    try:
        body_bytes = path.read_bytes()
    except OSError as exc:
        raise HTTPException(
            status_code=500,
            detail=f"Failed to read converted DOCX from disk: {exc}",
        )

    headers: dict = {
        # Tell requests the exact body size — prevents chunked encoding.
        "Content-Length": str(len(body_bytes)),
    }
    if content_type:
        headers["Content-Type"] = content_type

    try:
        put = requests.put(
            url,
            data=body_bytes,
            headers=headers,
            timeout=120,
        )
    except requests.RequestException as exc:
        raise HTTPException(
            status_code=502,
            detail=f"Failed to upload converted DOCX: {exc}",
        )

    print(
        f"[convert-from-url] PUT response: status={put.status_code}, "
        f"bytes_sent={len(body_bytes)}, "
        f"server_etag={put.headers.get('ETag', 'n/a')}",
    )
    if put.status_code >= 400:
        raise HTTPException(
            status_code=502,
            detail=(
                f"upload_url returned HTTP {put.status_code}: "
                f"{put.text[:300] if put.text else 'no body'}"
            ),
        )


@app.post("/convert-from-url")
def convert_pdf_from_url(payload: ConvertFromUrlRequest):
    """Download a PDF from S3, convert to DOCX, upload back to S3.

    Mirrors the eSigns backend's existing DOC→DOCX flow so the FE doesn't
    need a special-case branch: the caller hands us signed URLs and we
    return a JSON status when the upload settles.
    """
    safe_name = (payload.file_name or "document.pdf").replace("/", "_").replace("\\", "_")
    if not safe_name.lower().endswith(".pdf"):
        safe_name = safe_name + ".pdf"

    temp_id = os.urandom(8).hex()
    input_pdf_path = UPLOAD_DIR / f"{temp_id}_{safe_name}"
    output_docx_path = input_pdf_path.with_suffix(".docx")

    cv = None
    try:
        # 1. Download the source PDF.
        bytes_in = _stream_download_to(input_pdf_path, payload.download_url)
        print(f"[convert-from-url] downloaded {bytes_in} bytes → {input_pdf_path.name}")

        # 2. Convert PDF → DOCX with pdf2docx.
        from pdf2docx import Converter
        cv = Converter(str(input_pdf_path))
        cv.convert(str(output_docx_path), start=0, end=None)
        cv.close()
        cv = None
        if not output_docx_path.exists() or output_docx_path.stat().st_size == 0:
            raise HTTPException(
                status_code=500,
                detail="Converter produced an empty DOCX.",
            )
        print(f"[convert-from-url] converted → {output_docx_path.name} ({output_docx_path.stat().st_size} bytes)")

        # 3. Upload the converted DOCX to the provided signed URL.
        _put_docx(
            output_docx_path,
            payload.upload_url,
            content_type=payload.upload_content_type,
        )
        print(f"[convert-from-url] uploaded {output_docx_path.name}")

        return JSONResponse(
            status_code=200,
            content={
                "success": True,
                "message": "PDF converted and uploaded successfully.",
                "bytes_in": bytes_in,
                "bytes_out": output_docx_path.stat().st_size,
            },
        )

    except HTTPException as http_exc:
        # Re-raise so FastAPI can propagate the status code intact, but log
        # for the operator first so failed conversions are visible.
        print(f"[convert-from-url] HTTPException: {http_exc.status_code} — {http_exc.detail}")
        raise
    except Exception as exc:
        print("[convert-from-url] unexpected exception:")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}")
    finally:
        if cv:
            try:
                cv.close()
            except Exception:
                pass
        # Always purge temp files — there's no FileResponse holding a
        # reference here, so we don't need BackgroundTasks.
        remove_file(input_pdf_path)
        remove_file(output_docx_path)


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
