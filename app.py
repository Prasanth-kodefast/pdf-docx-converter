import io
import os
import re
import shutil
import traceback
from pathlib import Path
from typing import Any, Optional

import requests
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import (
    FileResponse,
    HTMLResponse,
    JSONResponse,
    StreamingResponse as responses_StreamingResponse,
)
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

    Two non-obvious requirements that broke this in production:

    1. **No chunked transfer.** When `requests.put(url, data=<file handle>)`
       is used, the library streams the body using
       `Transfer-Encoding: chunked`, which S3's v4 presigned PUT URLs
       reject unless they were explicitly signed for chunked payloads
       (they aren't by default). Browser `fetch()` PUTs work because
       they send the whole body with `Content-Length` and no
       Transfer-Encoding. Read the file fully into memory and pass
       bytes so `requests` behaves the same way.

    2. **Match the Content-Type the URL was signed with.** Our backend
       (`awsS3V3DataServiceProvider.getPreSignedUrl`) defaults PUT
       URLs without an explicit contentType to
       `application/x-www-form-urlencoded` and bakes that into the
       signature. The client PUT *must* send that exact Content-Type
       header or S3 fails with SignatureDoesNotMatch (403). Caller
       can override via `content_type` when the URL was signed for a
       different MIME.
    """
    try:
        body_bytes = path.read_bytes()
    except OSError as exc:
        raise HTTPException(
            status_code=500,
            detail=f"Failed to read converted DOCX from disk: {exc}",
        )

    # The backend signs editable-file PUT URLs with this Content-Type
    # (see api.esigns.io awsS3V3DataServiceProvider.getPreSignedUrl
    # around line 116). Sending a different value — or no header at all —
    # produces a SignatureDoesNotMatch 403, which is what made our
    # earlier "200 from FastAPI but the DOCX isn't there on S3" pattern.
    effective_content_type = content_type or "application/x-www-form-urlencoded"

    headers: dict = {
        # Tell requests the exact body size — prevents chunked encoding.
        "Content-Length": str(len(body_bytes)),
        "Content-Type": effective_content_type,
    }

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
        f"content_type_sent={effective_content_type}, "
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


# ─── EditorJS ↔ Track-Changes DOCX ─────────────────────────────────────────
# Endpoints that let the eSigns backend turn a CLM contract's per-page
# EditorJS blocks (with `<ins>` / `<mark>` redlines) into a real .docx
# with native Word Track Changes, and parse a returned .docx from a
# counterparty back into EditorJS blocks that preserve the same marks.
# See editorjs_to_docx.py + docx_to_editorjs.py for the OOXML details.


class ExportEditorToDocxRequest(BaseModel):
    """Payload for `/export-editor-to-docx`.

    The eSigns BE loads per-page blocks + headers/footers from Mongo and
    posts them here. We stay stateless — no auth, no DB, just the
    render.
    """

    pages: list = Field(
        ...,
        description=(
            "Per-page EditorJS blocks. Each entry is "
            "`{ pageNo: int, blocks: [...editorjs blocks...] }`."
        ),
    )
    title: str = Field(
        ...,
        description="Contract title. Appears in Word's title bar + as an H1.",
    )
    author: str = Field(
        ...,
        description=(
            "Human-readable author for `w:author` on every tracked "
            "change (typically 'FirstName LastName' or the email)."
        ),
    )
    message: Optional[str] = Field(
        default=None,
        description="Optional cover message rendered above the body.",
    )
    date: Optional[str] = Field(
        default=None,
        description=(
            "ISO-8601 date to stamp on every `w:ins` / `w:del`. Defaults "
            "to now-in-UTC when omitted."
        ),
    )
    headers: Optional[Any] = Field(
        default=None,
        description=(
            "Per-page header blocks, keyed by page index (`{0: {blocks: "
            "[...]}}`) or as a list of `{pageNo, blocks}`. Optional."
        ),
    )
    footers: Optional[Any] = Field(
        default=None,
        description="Per-page footer blocks, same shape as `headers`.",
    )
    enable_track_changes: Optional[bool] = Field(
        default=True,
        description=(
            "Whether Word should open the file with Track Changes ON so "
            "the counterparty's further edits are also recorded. Defaults "
            "true — turn off only if the sender wants a snapshot the "
            "counterparty can freely rewrite without new revisions."
        ),
    )


class ExportEditorToDocxToS3Request(ExportEditorToDocxRequest):
    """Same fields as the inline export, plus a signed S3 PUT URL to
    upload the rendered `.docx` to. When the eSigns BE calls us it
    knows the S3 key up front (it presigned it) so we upload directly
    instead of streaming the file back only for the BE to re-upload.
    """

    upload_url: str = Field(
        ...,
        description="Signed S3 PUT URL for the rendered DOCX.",
        min_length=8,
    )
    upload_content_type: Optional[str] = Field(
        default=None,
        description=(
            "Override Content-Type on the PUT. Defaults to whatever "
            "eSigns' presigner defaults to for editable files."
        ),
    )
    original_docx_url: Optional[str] = Field(
        default=None,
        description=(
            "Signed S3 GET URL of the counterparty's ORIGINAL uploaded "
            ".docx. When provided we open it, splice in the sender's "
            "<w:ins>/<w:del> at the matching paragraphs (leaving "
            "un-edited paragraphs pixel-identical), and upload the "
            "patched file — enterprise round-trip. When omitted we fall "
            "back to rendering a brand-new DOCX from the blocks."
        ),
    )


@app.post("/export-editor-to-docx")
def export_editor_to_docx(payload: ExportEditorToDocxRequest):
    """Render the given EditorJS pages into a Track-Changes DOCX and
    return the file bytes as the HTTP response. Convenient for
    debugging + previewing; production traffic goes through the
    S3-uploading sibling so the BE doesn't have to re-store the bytes.
    """
    from editorjs_to_docx import render_contract_docx

    try:
        docx_bytes = render_contract_docx(
            pages=payload.pages,
            meta={
                "title": payload.title,
                "author": payload.author,
                "message": payload.message,
                "date": payload.date,
                "headers": payload.headers,
                "footers": payload.footers,
                "enable_track_changes": bool(payload.enable_track_changes),
            },
        )
    except Exception as exc:
        print("[export-editor-to-docx] render failed:")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"DOCX render failed: {exc}")

    # Sanitize the filename in the Content-Disposition — Word rejects
    # some non-ASCII characters in the download prompt on Windows.
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", payload.title)[:80] or "contract"
    filename = f"{safe}.docx"

    return responses_StreamingResponse(
        io.BytesIO(docx_bytes),
        media_type=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Length": str(len(docx_bytes)),
        },
    )


@app.post("/export-editor-to-docx-to-s3")
def export_editor_to_docx_to_s3(payload: ExportEditorToDocxToS3Request):
    """Two paths depending on whether the caller supplied an
    `original_docx_url`:

      • YES → **patch-in-place**. Fetch the original DOCX, splice in
        the sender's <w:ins>/<w:del> at matching paragraphs, leave
        every un-edited paragraph byte-identical, upload the patched
        file. This is the enterprise round-trip Ironclad + DocuSign
        Negotiate use — the counterparty gets back a file that looks
        exactly like the one they sent, plus the sender's redlines.

      • NO  → **render from scratch**. Used for sender-initiated
        contracts where there IS no original .docx (the sender
        authored blocks directly in-app). Same as the old behaviour.

    Response is a small JSON with the upload status so the BE can log
    it — bytes uploaded, and which path was taken.
    """
    from editorjs_to_docx import render_contract_docx, patch_contract_docx

    used_patch = False
    try:
        if payload.original_docx_url:
            # Patch-in-place path. Fetch original bytes first.
            try:
                with requests.get(payload.original_docx_url, timeout=120) as r:
                    if r.status_code >= 400:
                        raise HTTPException(
                            status_code=502,
                            detail=(
                                f"original_docx_url returned HTTP {r.status_code}: "
                                f"{r.text[:400]}"
                            ),
                        )
                    original_bytes = r.content
            except requests.RequestException as exc:
                raise HTTPException(
                    status_code=502,
                    detail=f"Failed to fetch original DOCX: {exc}",
                )
            docx_bytes = patch_contract_docx(
                original_bytes=original_bytes,
                pages=payload.pages,
                meta={
                    "title": payload.title,
                    "author": payload.author,
                    "message": payload.message,
                    "date": payload.date,
                    "enable_track_changes": bool(payload.enable_track_changes),
                },
            )
            used_patch = True
        else:
            docx_bytes = render_contract_docx(
                pages=payload.pages,
                meta={
                    "title": payload.title,
                    "author": payload.author,
                    "message": payload.message,
                    "date": payload.date,
                    "headers": payload.headers,
                    "footers": payload.footers,
                    "enable_track_changes": bool(payload.enable_track_changes),
                },
            )
    except HTTPException:
        raise
    except Exception as exc:
        print("[export-editor-to-docx-to-s3] render/patch failed:")
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail=f"DOCX render failed: {exc}",
        )

    # Reuse the same _put_docx bytes-first-no-chunking recipe that the
    # PDF→DOCX flow already proved out on signed URLs. It writes to a
    # temp file so we can share the exact same helper; keeps the URL /
    # header pitfalls in one place.
    tmp_path = UPLOAD_DIR / f"editor_export_{os.getpid()}_{re.sub(r'[^A-Za-z0-9]+', '', payload.title)[:30] or 'contract'}.docx"
    try:
        tmp_path.write_bytes(docx_bytes)
        _put_docx(
            tmp_path,
            payload.upload_url,
            content_type=payload.upload_content_type,
        )
    finally:
        remove_file(tmp_path)

    return JSONResponse(
        {
            "success": True,
            "bytes": len(docx_bytes),
            "used_patch": used_patch,
        }
    )


class ParseDocxTrackedChangesRequest(BaseModel):
    """Payload for `/parse-docx-tracked-changes`. Either a signed
    download URL for the counterparty's returned .docx, or an inline
    upload via a separate multipart endpoint (below). Both routes exist
    because the eSigns BE prefers the URL path (S3-to-service, no
    proxying) but a manual QA flow benefits from posting a file
    directly."""

    download_url: str = Field(
        ...,
        description="Signed S3 GET URL for the counterparty's returned DOCX.",
        min_length=8,
    )


@app.post("/parse-docx-tracked-changes")
def parse_docx_tracked_changes(payload: ParseDocxTrackedChangesRequest):
    """Fetch a `.docx` from the given URL and return the parsed EditorJS
    blocks + counterparty authors + change count. See docx_to_editorjs.py
    for the OOXML → EditorJS mapping."""
    from docx_to_editorjs import parse_counterparty_docx

    try:
        with requests.get(payload.download_url, timeout=120) as r:
            if r.status_code >= 400:
                raise HTTPException(
                    status_code=502,
                    detail=f"download_url returned HTTP {r.status_code}",
                )
            docx_bytes = r.content
    except requests.RequestException as exc:
        raise HTTPException(
            status_code=502,
            detail=f"Failed to fetch counterparty DOCX: {exc}",
        )

    try:
        result = parse_counterparty_docx(docx_bytes)
    except Exception as exc:
        print("[parse-docx-tracked-changes] parse failed:")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Parse failed: {exc}")

    return JSONResponse(result)


@app.post("/parse-docx-tracked-changes/upload")
async def parse_docx_tracked_changes_upload(file: UploadFile = File(...)):
    """Multipart-upload variant of `/parse-docx-tracked-changes` for QA
    + manual runs. The eSigns BE uses the URL variant in production."""
    from docx_to_editorjs import parse_counterparty_docx

    try:
        docx_bytes = await file.read()
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Failed to read upload: {exc}")

    try:
        result = parse_counterparty_docx(docx_bytes)
    except Exception as exc:
        print("[parse-docx-tracked-changes/upload] parse failed:")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Parse failed: {exc}")

    return JSONResponse(result)


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
