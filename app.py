import os
import shutil
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse, HTMLResponse

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

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
