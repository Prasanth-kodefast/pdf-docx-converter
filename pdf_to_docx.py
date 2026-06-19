import sys
import argparse
from pathlib import Path

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
except Exception:
    pass

def print_status(message, success=True):
    prefix = "[✓]" if success else "[✗]"
    print(f"{prefix} {message}")

def convert_pdf_to_docx(pdf_path, docx_path, start_page=0, end_page=None):
    try:
        from pdf2docx import Converter
    except ImportError:
        print_status("pdf2docx is required. Install it using: pip install pdf2docx", False)
        return False

    cv = None
    try:
        print_status(f"Opening PDF file: {pdf_path}")
        cv = Converter(str(pdf_path))
        
        print_status("Parsing pages, layout structure, tables, and images...")
        # Convert all pages by default (start_page=0, end_page=None converts all)
        cv.convert(str(docx_path), start=start_page, end=end_page)
        
        cv.close()
        cv = None
        print_status(f"Conversion complete! Word file saved at: {docx_path}")
        return True
    except Exception as e:
        if cv:
            try:
                cv.close()
            except Exception:
                pass
        print_status(f"Failed to convert PDF to DOCX: {e}", False)
        import traceback
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF documents to editable Microsoft Word (.docx) files, preserving text, layout, tables, and images."
    )
    parser.add_argument("input_pdf", help="Path to the input PDF file")
    parser.add_argument("-o", "--output_docx", help="Path to the output DOCX file (optional)", default=None)
    parser.add_argument("-s", "--start", type=int, default=0, help="Start page index (0-indexed, default is 0)")
    parser.add_argument("-e", "--end", type=int, default=None, help="End page index (exclusive, default is all pages)")

    args = parser.parse_args()

    pdf_file = Path(args.input_pdf)
    if not pdf_file.exists():
        print_status(f"Input PDF file does not exist: {pdf_file}", False)
        sys.exit(1)

    if pdf_file.suffix.lower() != ".pdf":
        print_status(f"Input file must be a .pdf file. Got: {pdf_file.suffix}", False)
        sys.exit(1)

    if args.output_docx:
        docx_file = Path(args.output_docx)
    else:
        docx_file = pdf_file.with_suffix(".docx")

    # Perform the conversion
    success = convert_pdf_to_docx(pdf_file, docx_file, args.start, args.end)
    
    if success:
        print_status("Finished successfully!")
    else:
        print_status("Process failed.", False)
        sys.exit(1)

if __name__ == "__main__":
    main()
