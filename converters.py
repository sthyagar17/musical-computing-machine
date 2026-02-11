"""
Converters: transform CSV, TSV, JSON, Image (OCR), and PDF files into Excel.
"""

import csv
import io
import os

import pandas as pd


class ConversionError(Exception):
    """User-friendly conversion error."""


SUPPORTED_EXTENSIONS = {
    ".xlsx", ".xls",
    ".csv", ".tsv",
    ".json",
    ".jpg", ".jpeg", ".png",
    ".pdf",
}


def convert_to_excel(input_path: str, output_path: str) -> None:
    """Detect file type by extension and convert to .xlsx."""
    ext = os.path.splitext(input_path)[1].lower()

    handlers = {
        ".csv": _convert_csv,
        ".tsv": _convert_tsv,
        ".json": _convert_json,
        ".jpg": _convert_image,
        ".jpeg": _convert_image,
        ".png": _convert_image,
        ".pdf": _convert_pdf,
    }

    handler = handlers.get(ext)
    if handler is None:
        raise ConversionError(f"Unsupported file type: {ext}")

    handler(input_path, output_path)


def _convert_csv(input_path: str, output_path: str) -> None:
    """Convert CSV to Excel, auto-detecting delimiter."""
    try:
        with open(input_path, "r", newline="", encoding="utf-8-sig") as f:
            sample = f.read(8192)
        dialect = csv.Sniffer().sniff(sample)
        df = pd.read_csv(input_path, sep=dialect.delimiter, encoding="utf-8-sig")
    except Exception:
        # Fallback: let pandas guess
        df = pd.read_csv(input_path, encoding="utf-8-sig")

    if df.empty:
        raise ConversionError("CSV file is empty or could not be parsed.")
    df.to_excel(output_path, index=False, engine="openpyxl")


def _convert_tsv(input_path: str, output_path: str) -> None:
    """Convert TSV to Excel."""
    df = pd.read_csv(input_path, sep="\t", encoding="utf-8-sig")
    if df.empty:
        raise ConversionError("TSV file is empty or could not be parsed.")
    df.to_excel(output_path, index=False, engine="openpyxl")


def _convert_json(input_path: str, output_path: str) -> None:
    """Convert JSON to Excel using json_normalize for nested data."""
    import json

    with open(input_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, dict):
        # If top-level dict, try to find an array value to normalize
        for key, val in data.items():
            if isinstance(val, list):
                data = val
                break
        else:
            # Single object - wrap in list
            data = [data]

    if not isinstance(data, list):
        raise ConversionError("JSON must be an array or object with an array field.")

    df = pd.json_normalize(data)
    if df.empty:
        raise ConversionError("JSON produced no tabular data.")
    df.to_excel(output_path, index=False, engine="openpyxl")


def _convert_image(input_path: str, output_path: str) -> None:
    """Convert image to Excel via OCR (pytesseract)."""
    try:
        from PIL import Image
        import pytesseract
    except ImportError:
        raise ConversionError(
            "Image conversion requires Pillow and pytesseract. "
            "Install them with: pip install Pillow pytesseract"
        )

    # Set Tesseract path on Windows if not on PATH
    import shutil
    if not shutil.which("tesseract"):
        win_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.isfile(win_path):
            pytesseract.pytesseract.tesseract_cmd = win_path

    try:
        img = Image.open(input_path)
    except Exception as exc:
        raise ConversionError(f"Cannot open image: {exc}")

    try:
        ocr_data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
    except Exception as exc:
        raise ConversionError(
            f"OCR failed: {exc}. "
            "Make sure Tesseract OCR is installed on your system."
        )

    # Reconstruct rows from OCR output: group text by block/line numbers
    rows = {}
    for i, text in enumerate(ocr_data["text"]):
        text = text.strip()
        if not text:
            continue
        block = ocr_data["block_num"][i]
        line = ocr_data["line_num"][i]
        key = (block, line)
        if key not in rows:
            rows[key] = []
        rows[key].append(text)

    if not rows:
        raise ConversionError("OCR could not extract any text from the image.")

    sorted_rows = [rows[k] for k in sorted(rows.keys())]

    # Pad rows to equal length
    max_cols = max(len(r) for r in sorted_rows)
    for row in sorted_rows:
        row.extend([""] * (max_cols - len(row)))

    # Use first row as header if it looks like a header (all strings, no digits)
    first_row = sorted_rows[0]
    if all(not cell.replace(".", "").replace(",", "").isdigit() for cell in first_row):
        df = pd.DataFrame(sorted_rows[1:], columns=first_row)
    else:
        df = pd.DataFrame(sorted_rows)

    df.to_excel(output_path, index=False, engine="openpyxl")


def _convert_pdf(input_path: str, output_path: str) -> None:
    """Convert PDF to Excel: extract tables, fall back to text."""
    try:
        import pdfplumber
    except ImportError:
        raise ConversionError(
            "PDF conversion requires pdfplumber. "
            "Install it with: pip install pdfplumber"
        )

    all_rows = []
    has_tables = False

    try:
        with pdfplumber.open(input_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    has_tables = True
                    for table in tables:
                        all_rows.extend(table)
                else:
                    # Fallback: extract text lines
                    text = page.extract_text()
                    if text:
                        for line in text.split("\n"):
                            line = line.strip()
                            if line:
                                # Split on multiple spaces to approximate columns
                                cells = [c.strip() for c in line.split("  ") if c.strip()]
                                all_rows.append(cells)
    except Exception as exc:
        raise ConversionError(f"Failed to read PDF: {exc}")

    if not all_rows:
        raise ConversionError("PDF contains no extractable text or tables.")

    # Pad rows to equal length
    max_cols = max(len(r) for r in all_rows)
    for row in all_rows:
        while len(row) < max_cols:
            row.append("")
        # Replace None values
        for i, cell in enumerate(row):
            if cell is None:
                row[i] = ""

    # If tables were found, treat first row as header
    if has_tables and len(all_rows) > 1:
        df = pd.DataFrame(all_rows[1:], columns=all_rows[0])
    else:
        df = pd.DataFrame(all_rows)

    df.to_excel(output_path, index=False, engine="openpyxl")
