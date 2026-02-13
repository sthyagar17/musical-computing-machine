# Excel Merger

A Flask web app for merging Excel files and converting other file formats (CSV, TSV, JSON, PDF, images) to Excel.

## Features

### Merge Files
Upload two files and merge them using one of three strategies:
- **Append rows** - stack two sheets vertically (union of columns)
- **Join on a column** - inner, left, right, or outer join on a shared column
- **Merge sheets** - combine all sheets from both files into one workbook

Sheet and column selectors update dynamically based on uploaded file contents.

### Convert to Excel
Upload a single file and convert it to `.xlsx` with a table preview and download:

| Format | How it works |
|--------|-------------|
| **CSV** | Auto-detects delimiter via `csv.Sniffer`, falls back to pandas default |
| **TSV** | Tab-separated, read with `pandas.read_csv(sep='\t')` |
| **JSON** | Handles arrays, objects, and nested data via `pd.json_normalize` |
| **PDF** | Parses credit card statement transactions (Trans. Date, Post Date, Description, Amount); skips payments and credits (negative amounts) |
| **JPG / PNG** | OCR via Tesseract with image preprocessing (grayscale, 2x upscale, sharpen, contrast boost); uses header row positions to define column boundaries |

### Auto-conversion in Merge Mode
Non-Excel files (CSV, JSON, PDF, etc.) uploaded in merge mode are automatically converted to Excel before merging. A notice shows which files were converted.

## Setup

### Python dependencies

```bash
pip install -r requirements.txt
```

### Tesseract OCR (optional, for image conversion only)

Install [Tesseract OCR](https://github.com/tesseract-ocr/tesseract):

| Platform | Command |
|----------|---------|
| Windows | `winget install UB-Mannheim.TesseractOCR` |
| macOS | `brew install tesseract` |
| Linux | `sudo apt install tesseract-ocr` |

On Windows, the app automatically finds Tesseract at `C:\Program Files\Tesseract-OCR\tesseract.exe` if it's not on your PATH.

All other conversions (CSV, TSV, JSON, PDF) work without Tesseract.

## Usage

```bash
python app.py
```

Open http://127.0.0.1:5000 in your browser.

### Merge workflow
1. Select two files and click **Upload**
2. Choose merge type, sheets, and options
3. Click **Merge** to see a preview
4. Click **Download merged.xlsx**

### Convert workflow
1. Switch to the **Convert to Excel** tab
2. Select a file and click **Convert to Excel**
3. Preview the table (use the sheet selector for multi-sheet PDFs)
4. Click **Download converted.xlsx**

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/` | Main page |
| POST | `/upload` | Upload two files for merging, returns sheet/column info |
| POST | `/merge` | Merge uploaded files with specified options |
| GET | `/download` | Download the merged Excel file |
| POST | `/convert` | Upload and convert a single file to Excel |
| GET | `/convert-sheet?sheet=name` | Preview a specific sheet from the converted file |
| GET | `/download-converted` | Download the converted Excel file |

## Project Structure

```
excel-merger/
  app.py             # Flask routes, merge logic, upload handling
  converters.py      # File-to-Excel conversion handlers (CSV, TSV, JSON, PDF, image OCR)
  requirements.txt   # Python dependencies
  templates/
    index.html       # Frontend template with merge and convert tabs
  static/
    script.js        # Client-side logic (uploads, previews, tab switching)
    style.css        # Responsive styles
```

## Dependencies

| Package | Purpose |
|---------|---------|
| Flask | Web framework |
| pandas | Data manipulation and Excel I/O |
| openpyxl | Excel file engine |
| pdfplumber | PDF text extraction |
| Pillow | Image loading for OCR |
| pytesseract | OCR engine interface |
