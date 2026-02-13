"""
Converters: transform CSV, TSV, JSON, Image (OCR), and PDF files into Excel.
"""

import csv
import io
import json
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
        from PIL import ImageEnhance, ImageFilter
    except ImportError:
        pass

    try:
        img = Image.open(input_path)
    except Exception as exc:
        raise ConversionError(f"Cannot open image: {exc}")

    # Preprocess: grayscale, upscale 2x, sharpen, boost contrast for better OCR
    img = img.convert("L")
    img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)
    img = img.filter(ImageFilter.SHARPEN)
    img = ImageEnhance.Contrast(img).enhance(2.0)

    try:
        ocr_data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
    except Exception as exc:
        raise ConversionError(
            f"OCR failed: {exc}. "
            "Make sure Tesseract OCR is installed on your system."
        )

    # Collect words with their positions
    words = []
    for i, text in enumerate(ocr_data["text"]):
        text = text.strip()
        if not text:
            continue
        words.append({
            "text": text,
            "left": ocr_data["left"][i],
            "top": ocr_data["top"][i],
            "width": ocr_data["width"][i],
            "right": ocr_data["left"][i] + ocr_data["width"][i],
        })

    if not words:
        raise ConversionError("OCR could not extract any text from the image.")

    # Group words into rows by y-coordinate (top), with tolerance for
    # slight vertical misalignment between words on the same visual row.
    words.sort(key=lambda w: (w["top"], w["left"]))
    rows_by_y = []  # list of (avg_y, [words])
    for w in words:
        placed = False
        for row in rows_by_y:
            if abs(w["top"] - row[0]) <= 15:
                row[1].append(w)
                # Update average y
                row[0] = sum(rw["top"] for rw in row[1]) // len(row[1])
                placed = True
                break
        if not placed:
            rows_by_y.append([w["top"], [w]])

    # Sort rows top to bottom, words within each row left to right
    rows_by_y.sort(key=lambda r: r[0])
    for row in rows_by_y:
        row[1].sort(key=lambda w: w["left"])

    # Use the HEADER ROW (first row by y-position) to define column boundaries.
    # Header words are well-separated (e.g. "Date", "Description", "Location", "Amount").
    header_words = rows_by_y[0][1]
    num_cols = len(header_words)

    if num_cols <= 1:
        # Single column - just concatenate each row
        table_rows = [[" ".join(w["text"] for w in r[1])] for r in rows_by_y]
    else:
        # Define column boundaries using midpoints between header words.
        col_bounds = []  # list of (col_left, col_right) for each column
        for ci, hw in enumerate(header_words):
            if ci == 0:
                left_bound = 0
            else:
                prev_right = header_words[ci - 1]["right"]
                left_bound = (prev_right + hw["left"]) // 2
            if ci == num_cols - 1:
                right_bound = 999999
            else:
                next_left = header_words[ci + 1]["left"]
                right_bound = (hw["right"] + next_left) // 2
            col_bounds.append((left_bound, right_bound))

        def _get_col_index(x):
            for ci, (cl, cr) in enumerate(col_bounds):
                if cl <= x < cr:
                    return ci
            return num_cols - 1

        # Build table rows by assigning each word to a column
        table_rows = []
        for _, row_words in rows_by_y:
            cells = [""] * num_cols
            for w in row_words:
                ci = _get_col_index(w["left"])
                if cells[ci]:
                    cells[ci] += " " + w["text"]
                else:
                    cells[ci] = w["text"]
            table_rows.append(cells)

    # Use first row as header if it looks like a header (all non-numeric)
    first_row = table_rows[0]
    if all(not cell.replace(".", "").replace(",", "").replace("/", "").isdigit()
           for cell in first_row if cell):
        df = pd.DataFrame(table_rows[1:], columns=first_row)
    else:
        df = pd.DataFrame(table_rows)

    df.to_excel(output_path, index=False, engine="openpyxl")


def _load_rules():
    """Load categorization rules from rules.json next to this module."""
    rules_path = os.path.join(os.path.dirname(__file__), "rules.json")
    try:
        with open(rules_path, "r", encoding="utf-8") as f:
            return json.load(f).get("rules", [])
    except (FileNotFoundError, json.JSONDecodeError):
        return []


def _apply_rules(description, rules):
    """Return (type, hl_exp_category, exp_category) for the first matching rule."""
    desc_upper = description.upper()
    for rule in rules:
        for keyword in rule["keywords"]:
            if keyword.upper() in desc_upper:
                return rule["type"], rule["hl_exp_category"], rule["exp_category"]
    return "", "", ""


def _convert_pdf(input_path: str, output_path: str) -> None:
    """Convert PDF to Excel: extract transaction rows from credit card statements."""
    try:
        import pdfplumber
    except ImportError:
        raise ConversionError(
            "PDF conversion requires pdfplumber. "
            "Install it with: pip install pdfplumber"
        )
    import re
    from datetime import datetime

    # Pattern: MM/DD at start, optional second MM/DD, description, then $amount
    txn_re = re.compile(
        r'^(\d{2}/\d{2})\s+'           # Trans. date
        r'(?:(\d{2}/\d{2})\s+)?'       # Post date (optional)
        r'(.+?)\s+'                     # Description
        r'(-?\$[\d,]+\.\d{2})\b'        # Amount
    )
    # Pattern to extract billing period for year context: MM/DD/YY-MM/DD/YY
    billing_re = re.compile(r'Billing Period:\s*(\d{2}/\d{2}/(\d{2}))-(\d{2}/\d{2}/(\d{2}))')

    all_text = []
    start_year = None
    end_year = None

    try:
        with pdfplumber.open(input_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                all_text.append(text)

                # Extract billing period years
                if start_year is None:
                    bm = billing_re.search(text)
                    if bm:
                        start_year = 2000 + int(bm.group(2))
                        end_year = 2000 + int(bm.group(4))
    except Exception as exc:
        raise ConversionError(f"Failed to read PDF: {exc}")

    if not all_text:
        raise ConversionError("PDF contains no extractable text.")

    # Fallback: try to get year from filename (e.g., "Jan 2024.pdf")
    if start_year is None:
        year_match = re.search(r'(20\d{2})', os.path.basename(input_path))
        if year_match:
            end_year = int(year_match.group(1))
            start_year = end_year
        else:
            start_year = end_year = datetime.now().year

    # Load categorization rules
    rules = _load_rules()

    # Parse transactions
    MONTH_NAMES = [
        "", "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    ]

    transactions = []
    for text in all_text:
        for line in text.split('\n'):
            m = txn_re.match(line.strip())
            if not m:
                continue
            amount_str = m.group(4)
            if amount_str.startswith("-"):
                continue

            trans_mm, trans_dd = m.group(1).split("/")
            month = int(trans_mm)
            day = int(trans_dd)

            # Determine year: if billing spans Dec-Jan, Dec dates use start_year
            if start_year != end_year and month >= 10:
                year = start_year
            else:
                year = end_year

            date_str = f"{month:02d}/{day:02d}/{year}"
            month_name = MONTH_NAMES[month]
            post_date = m.group(2) or ""
            description = m.group(3).strip()
            amount = float(amount_str.replace("$", "").replace(",", ""))

            txn_type, hl_cat, exp_cat = _apply_rules(description, rules)

            transactions.append([
                post_date,      # Details (Post Date)
                date_str,       # Date
                month_name,     # Month
                day,            # Day
                year,           # Year
                description,    # Description
                amount,         # Amount
                txn_type,       # Type
                hl_cat,         # HL_Exp_Category
                exp_cat,        # Exp_Category
            ])

    if not transactions:
        raise ConversionError(
            "No transaction rows found with the expected format "
            "(Trans. Date, Post Date, Description, Amount)."
        )

    df = pd.DataFrame(transactions, columns=[
        "Details", "Date", "Month", "Day", "Year",
        "Description", "Amount", "Type", "HL_Exp_Category", "Exp_Category",
    ])
    df.to_excel(output_path, index=False, engine="openpyxl")
