# Excel Merger

A Flask web app for merging Excel files and converting other file formats (CSV, TSV, JSON, PDF, images) to Excel.

## Features

### Merge Files
- **Append rows** - stack two sheets vertically
- **Join on a column** - inner, left, right, or outer join
- **Merge sheets** - combine all sheets from both files into one workbook

### Convert to Excel
Upload a single file and convert it to `.xlsx` with preview and download:
- **CSV / TSV** - auto-detects delimiters
- **JSON** - handles nested objects and arrays via `json_normalize`
- **PDF (credit card statements)** - extracts transaction rows (Trans. Date, Post Date, Description, Amount) via text parsing; skips payments/credits (negative amounts)
- **Images (JPG/PNG)** - OCR-based table extraction via Tesseract with image preprocessing (grayscale, 2x upscale, sharpen, contrast boost) for improved accuracy; uses header row positions to define column boundaries

Multi-sheet PDFs are supported with a sheet selector in the preview UI.

Non-Excel files uploaded in merge mode are auto-converted before merging.

## Setup

```bash
pip install -r requirements.txt
```

For image (OCR) support, install [Tesseract OCR](https://github.com/tesseract-ocr/tesseract):
- **Windows**: `winget install UB-Mannheim.TesseractOCR`
- **macOS**: `brew install tesseract`
- **Linux**: `sudo apt install tesseract-ocr`

All other conversions (CSV, TSV, JSON, PDF) work without Tesseract.

## Usage

```bash
python app.py
```

Open http://127.0.0.1:5000 in your browser.

## Project Structure

```
excel-merger/
  app.py             # Flask routes and merge logic
  converters.py      # File-to-Excel conversion handlers
  requirements.txt   # Python dependencies
  templates/
    index.html       # Frontend template
  static/
    script.js        # Client-side logic
    style.css        # Styles
```

## Dependencies

- Flask
- pandas
- openpyxl
- pdfplumber
- Pillow
- pytesseract
