# PDF to Word DOCX Converter (CLI & Web API)

This project provides both a command-line interface (CLI) and a FastAPI Web Endpoint with a premium glassmorphic UI to convert PDF files into editable Microsoft Word (`.docx`) documents.

## Features

- **Structure Preservation**: Detects and groups paragraphs together.
- **Table Reconstruction**: Detects boundaries to build clean, editable Microsoft Word tables.
- **Image Extraction**: Extracts PDF images and embeds them back inline.
- **Web Interface**: Modern, drag-and-drop web UI for uploading PDFs and downloading converted DOCX files.
- **REST Endpoint**: Single API endpoint (`/convert`) for programmatically uploading and converting files.

---

## Quick Start

### 1. Setup Virtual Environment

Create and activate a virtual environment to manage dependencies locally:

```bash
# Create the virtual environment
python -m venv venv

# Activate on Windows (PowerShell)
.\venv\Scripts\Activate.ps1

# Activate on Windows (Command Prompt)
venv\Scripts\activate.bat

# Activate on macOS/Linux
source venv/bin/activate
```

### 2. Install Requirements

```bash
pip install -r requirements.txt
```

---

## Option A: Running the Web App & API Endpoint (Recommended)

Start the FastAPI backend server:

```bash
python app.py
```

1. Open your browser and go to: **`http://127.0.0.1:8000`**
2. Drag and drop any PDF file onto the interface, or click to upload.
3. Click **Convert to DOCX**. Once the conversion completes, the editable Word document will automatically download to your computer.

### Using the API Endpoint Programmatically

You can perform conversions programmatically by sending a POST request to `/convert`:

```bash
curl -X POST "http://127.0.0.1:8000/convert" -F "file=@your_document.pdf" --output converted_document.docx
```

---

## Option B: Using the CLI Script

If you prefer to run it directly from your terminal, you can convert a file using the CLI script:

### 1. Generate a Sample PDF (Optional)
If you don't have a PDF file on hand to test, run this script to generate a sample PDF containing text and tables:
```bash
python generate_test_pdf.py
```
This creates a sample file at `samples/sample.pdf`.

### 2. Convert to DOCX
```bash
# Convert the sample PDF to a DOCX file
python pdf_to_docx.py samples/sample.pdf

# Or convert a specific PDF to a custom output path
python pdf_to_docx.py path/to/your.pdf -o path/to/output.docx
```
