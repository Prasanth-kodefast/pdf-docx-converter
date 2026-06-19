import os
import sys
import argparse
from pathlib import Path

# Helper to print status
def print_status(message, success=True):
    prefix = "[✓]" if success else "[✗]"
    print(f"{prefix} {message}")

# 1. Text to PDF Conversion
def convert_txt_to_pdf(input_path, output_path):
    try:
        from fpdf import FPDF
    except ImportError:
        print_status("fpdf2 is required for text conversion. Install it using: pip install fpdf2", False)
        return False

    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Helvetica", size=11)
        
        with open(input_path, 'r', encoding='utf-8', errors='replace') as f:
            for line in f:
                pdf.cell(0, 6, txt=line.rstrip('\n'), ln=True)
                
        pdf.output(output_path)
        print_status(f"Converted text file to {output_path}")
        return True
    except Exception as e:
        print_status(f"Failed to convert text to PDF: {e}", False)
        return False

# 2. Image to PDF Conversion
def convert_img_to_pdf(input_path, output_path):
    try:
        from PIL import Image
    except ImportError:
        print_status("Pillow is required for image conversion. Install it using: pip install Pillow", False)
        return False

    try:
        image = Image.open(input_path)
        if image.mode in ("RGBA", "P"):
            image = image.convert("RGB")
        image.save(output_path, "PDF", resolution=100.0)
        print_status(f"Converted image to {output_path}")
        return True
    except Exception as e:
        print_status(f"Failed to convert image to PDF: {e}", False)
        return False

# 3. HTML to PDF Conversion
def convert_html_to_pdf(input_path, output_path):
    try:
        from xhtml2pdf import pisa
    except ImportError:
        print_status("xhtml2pdf is required for HTML conversion. Install it using: pip install xhtml2pdf", False)
        return False

    try:
        with open(input_path, "r", encoding="utf-8") as html_file:
            source_html = html_file.read()
            
        with open(output_path, "w+b") as result_file:
            pisa_status = pisa.CreatePDF(source_html, dest=result_file)
            
        if pisa_status.err:
            print_status(f"Error during HTML to PDF conversion.", False)
            return False
            
        print_status(f"Converted HTML to {output_path}")
        return True
    except Exception as e:
        print_status(f"Failed to convert HTML to PDF: {e}", False)
        return False

# 4. DOCX (Word) to PDF Conversion
def convert_docx_to_pdf(input_path, output_path):
    if sys.platform == "win32":
        try:
            import comtypes.client
            print_status("Using Windows COM interface to convert DOCX...")
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            
            abs_input = os.path.abspath(input_path)
            abs_output = os.path.abspath(output_path)
            
            doc = word.Documents.Open(abs_input)
            doc.SaveAs(abs_output, FileFormat=17)
            doc.Close()
            word.Quit()
            print_status(f"Converted DOCX to {output_path}")
            return True
        except Exception as e:
            print_status(f"Failed using MS Word COM interface: {e}", False)
            print_status("Make sure Microsoft Word is installed and activated.", False)
            
    try:
        from docx2pdf import convert
        print_status("Attempting conversion via docx2pdf library...")
        convert(input_path, output_path)
        print_status(f"Converted DOCX to {output_path}")
        return True
    except ImportError:
        print_status("docx2pdf is not installed. Install it using: pip install docx2pdf", False)
    except Exception as e:
        print_status(f"docx2pdf conversion failed: {e}", False)

    print_status("Please ensure LibreOffice is installed if you are on Linux/macOS, or Microsoft Word if on Windows.", False)
    return False

def main():
    parser = argparse.ArgumentParser(description="Convert various document formats (TXT, HTML, PNG/JPG, DOCX) to PDF.")
    parser.add_argument("input", help="Path to the input file to convert")
    parser.add_argument("-o", "--output", help="Path to the output PDF file (optional)", default=None)

    args = parser.parse_args()

    input_file = Path(args.input)
    if not input_file.exists():
        print_status(f"Input file does not exist: {input_file}", False)
        sys.exit(1)

    if args.output:
        output_file = Path(args.output)
    else:
        output_file = input_file.with_suffix(".pdf")

    suffix = input_file.suffix.lower()
    success = False

    if suffix in (".txt",):
        success = convert_txt_to_pdf(input_file, output_file)
    elif suffix in (".png", ".jpg", ".jpeg", ".bmp", ".gif"):
        success = convert_img_to_pdf(input_file, output_file)
    elif suffix in (".html", ".htm"):
        success = convert_html_to_pdf(input_file, output_file)
    elif suffix in (".docx", ".doc"):
        success = convert_docx_to_pdf(input_file, output_file)
    else:
        print_status(f"Unsupported file type '{suffix}'. Supported extensions: .txt, .docx, .html, .png, .jpg, .jpeg, .bmp, .gif", False)
        sys.exit(1)

    if success:
        print_status("Conversion completed successfully!")
    else:
        print_status("Conversion failed.", False)
        sys.exit(1)

if __name__ == "__main__":
    main()
