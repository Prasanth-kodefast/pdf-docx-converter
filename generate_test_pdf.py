from fpdf import FPDF
from pathlib import Path

def generate_sample_pdf(output_path):
    print(f"Generating sample PDF with text and tables at: {output_path}")
    
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "PDF to Word DOCX Conversion Test", ln=True, align="C")
    pdf.ln(10)
    
    pdf.set_font("Helvetica", "", 12)
    pdf.multi_cell(0, 7, "This document serves as a test file to verify the layout, structure, and table reconstruction capability of the pdf2docx converter tool in Python.")
    pdf.ln(10)
    
    # Let's create a table
    # Set headers
    pdf.set_font("Helvetica", "B", 11)
    col_width = 50
    row_height = 8
    
    pdf.cell(col_width, row_height, "Department", border=1)
    pdf.cell(col_width, row_height, "Staff Count", border=1)
    pdf.cell(col_width, row_height, "Budget Status", border=1)
    pdf.ln()
    
    # Rows
    pdf.set_font("Helvetica", "", 11)
    data = [
        ("Engineering", "45", "Approved"),
        ("Product Design", "12", "Pending Approval"),
        ("Customer Success", "28", "Approved"),
        ("Marketing & Sales", "18", "Under Review")
    ]
    
    for row in data:
        pdf.cell(col_width, row_height, row[0], border=1)
        pdf.cell(col_width, row_height, row[1], border=1)
        pdf.cell(col_width, row_height, row[2], border=1)
        pdf.ln()
        
    pdf.ln(10)
    pdf.multi_cell(0, 7, "If the converter works successfully, this table will be reconstructed as a native Microsoft Word Table object in the output .docx file, with fully editable cells rather than simple static text boxes.")
    
    pdf.output(output_path)
    print("Sample PDF generated successfully!")

if __name__ == "__main__":
    out_dir = Path("samples")
    out_dir.mkdir(exist_ok=True)
    generate_sample_pdf(out_dir / "sample.pdf")
