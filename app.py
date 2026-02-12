import os
import tempfile
import uuid

from flask import Flask, render_template, request, jsonify, send_file, session

import pandas as pd

from converters import convert_to_excel, ConversionError, SUPPORTED_EXTENSIONS

app = Flask(__name__)
app.secret_key = os.urandom(24)

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "excel_merger")
os.makedirs(UPLOAD_DIR, exist_ok=True)

EXCEL_EXTENSIONS = {".xlsx", ".xls"}


def _session_dir():
    """Return a per-session temp directory, creating it if needed."""
    sid = session.get("sid")
    if not sid:
        sid = uuid.uuid4().hex
        session["sid"] = sid
    path = os.path.join(UPLOAD_DIR, sid)
    os.makedirs(path, exist_ok=True)
    return path


def _is_excel(filename):
    return os.path.splitext(filename)[1].lower() in EXCEL_EXTENSIONS


def _save_and_convert(file_storage, dest_xlsx_path):
    """Save uploaded file and convert to xlsx if needed.

    Returns a tuple (was_converted: bool, original_ext: str).
    """
    original_name = file_storage.filename
    ext = os.path.splitext(original_name)[1].lower()

    if ext in EXCEL_EXTENSIONS:
        file_storage.save(dest_xlsx_path)
        return False, ext

    # Save with original extension, then convert
    sdir = os.path.dirname(dest_xlsx_path)
    temp_input = os.path.join(sdir, f"_temp_input{ext}")
    file_storage.save(temp_input)
    try:
        convert_to_excel(temp_input, dest_xlsx_path)
    finally:
        if os.path.exists(temp_input):
            os.remove(temp_input)
    return True, ext


@app.route("/")
def index():
    accept_str = ",".join(sorted(SUPPORTED_EXTENSIONS))
    return render_template("index.html", supported_extensions=accept_str)


@app.route("/upload", methods=["POST"])
def upload():
    file1 = request.files.get("file1")
    file2 = request.files.get("file2")

    if not file1 or not file2:
        return jsonify({"error": "Please upload two files."}), 400

    sdir = _session_dir()
    path1 = os.path.join(sdir, "file1.xlsx")
    path2 = os.path.join(sdir, "file2.xlsx")

    conversions = []
    try:
        conv1, ext1 = _save_and_convert(file1, path1)
        if conv1:
            conversions.append(f"{file1.filename} ({ext1} → .xlsx)")
        conv2, ext2 = _save_and_convert(file2, path2)
        if conv2:
            conversions.append(f"{file2.filename} ({ext2} → .xlsx)")
    except ConversionError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"Failed to process files: {exc}"}), 400

    try:
        xl1 = pd.ExcelFile(path1, engine="openpyxl")
        xl2 = pd.ExcelFile(path2, engine="openpyxl")
    except Exception as exc:
        return jsonify({"error": f"Failed to read Excel files: {exc}"}), 400

    # Read columns for each sheet
    def sheet_info(xl):
        info = {}
        for name in xl.sheet_names:
            df = xl.parse(name, nrows=0)
            info[name] = list(df.columns.astype(str))
        return info

    result = {
        "file1_sheets": sheet_info(xl1),
        "file2_sheets": sheet_info(xl2),
    }
    if conversions:
        result["conversions"] = conversions

    return jsonify(result)


@app.route("/merge", methods=["POST"])
def merge():
    data = request.get_json()
    merge_type = data.get("merge_type")
    sheet1 = data.get("sheet1")
    sheet2 = data.get("sheet2")
    join_column = data.get("join_column")
    join_how = data.get("join_how", "inner")

    sdir = _session_dir()
    path1 = os.path.join(sdir, "file1.xlsx")
    path2 = os.path.join(sdir, "file2.xlsx")

    if not os.path.exists(path1) or not os.path.exists(path2):
        return jsonify({"error": "Files not found. Please upload again."}), 400

    try:
        if merge_type == "append":
            df1 = pd.read_excel(path1, sheet_name=sheet1, engine="openpyxl")
            df2 = pd.read_excel(path2, sheet_name=sheet2, engine="openpyxl")
            result = pd.concat([df1, df2], ignore_index=True)
            out = os.path.join(sdir, "merged.xlsx")
            result.to_excel(out, index=False, engine="openpyxl")

            preview = result.head(20)
            return jsonify({
                "columns": list(preview.columns.astype(str)),
                "rows": preview.fillna("").astype(str).values.tolist(),
                "total_rows": len(result),
            })

        elif merge_type == "join":
            if not join_column:
                return jsonify({"error": "Please select a column to join on."}), 400
            df1 = pd.read_excel(path1, sheet_name=sheet1, engine="openpyxl")
            df2 = pd.read_excel(path2, sheet_name=sheet2, engine="openpyxl")
            result = pd.merge(df1, df2, on=join_column, how=join_how)
            out = os.path.join(sdir, "merged.xlsx")
            result.to_excel(out, index=False, engine="openpyxl")

            preview = result.head(20)
            return jsonify({
                "columns": list(preview.columns.astype(str)),
                "rows": preview.fillna("").astype(str).values.tolist(),
                "total_rows": len(result),
            })

        elif merge_type == "sheets":
            out = os.path.join(sdir, "merged.xlsx")
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                xl1 = pd.ExcelFile(path1, engine="openpyxl")
                for name in xl1.sheet_names:
                    df = xl1.parse(name)
                    safe = name if name not in writer.sheets else f"{name}_file1"
                    df.to_excel(writer, sheet_name=safe, index=False)

                xl2 = pd.ExcelFile(path2, engine="openpyxl")
                for name in xl2.sheet_names:
                    df = xl2.parse(name)
                    safe = name if name not in writer.sheets else f"{name}_file2"
                    df.to_excel(writer, sheet_name=safe, index=False)

            # Preview: show first sheet's first 20 rows
            preview_df = pd.read_excel(out, sheet_name=0, engine="openpyxl")
            preview = preview_df.head(20)

            xl_out = pd.ExcelFile(out, engine="openpyxl")
            return jsonify({
                "columns": list(preview.columns.astype(str)),
                "rows": preview.fillna("").astype(str).values.tolist(),
                "total_rows": len(preview_df),
                "sheet_names": xl_out.sheet_names,
            })

        else:
            return jsonify({"error": "Unknown merge type."}), 400

    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.route("/download")
def download():
    sdir = _session_dir()
    out = os.path.join(sdir, "merged.xlsx")
    if not os.path.exists(out):
        return jsonify({"error": "No merged file found."}), 404
    return send_file(out, as_attachment=True, download_name="merged.xlsx")


@app.route("/convert", methods=["POST"])
def convert():
    """Standalone convert: single file upload -> Excel preview."""
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "Please upload a file."}), 400

    ext = os.path.splitext(file.filename)[1].lower()
    if ext in EXCEL_EXTENSIONS:
        return jsonify({"error": "File is already an Excel file."}), 400
    if ext not in SUPPORTED_EXTENSIONS:
        return jsonify({"error": f"Unsupported file type: {ext}"}), 400

    sdir = _session_dir()
    temp_input = os.path.join(sdir, f"convert_input{ext}")
    output_path = os.path.join(sdir, "converted.xlsx")

    file.save(temp_input)
    try:
        convert_to_excel(temp_input, output_path)
    except ConversionError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"Conversion failed: {exc}"}), 400
    finally:
        if os.path.exists(temp_input):
            os.remove(temp_input)

    # Read and return preview (handle multi-sheet files from PDF conversion)
    try:
        xl = pd.ExcelFile(output_path, engine="openpyxl")
        sheet_names = xl.sheet_names

        # Preview the first sheet
        df = xl.parse(sheet_names[0])
        preview = df.head(20)
        result = {
            "columns": list(preview.columns.astype(str)),
            "rows": preview.fillna("").astype(str).values.tolist(),
            "total_rows": len(df),
            "original_name": file.filename,
        }
        if len(sheet_names) > 1:
            result["sheet_names"] = sheet_names
        return jsonify(result)
    except Exception as exc:
        return jsonify({"error": f"Failed to read converted file: {exc}"}), 400


@app.route("/convert-sheet")
def convert_sheet():
    """Return preview of a specific sheet from the converted file."""
    sheet = request.args.get("sheet")
    if not sheet:
        return jsonify({"error": "No sheet specified."}), 400

    sdir = _session_dir()
    output_path = os.path.join(sdir, "converted.xlsx")
    if not os.path.exists(output_path):
        return jsonify({"error": "No converted file found."}), 404

    try:
        df = pd.read_excel(output_path, sheet_name=sheet, engine="openpyxl")
        preview = df.head(20)
        return jsonify({
            "columns": list(preview.columns.astype(str)),
            "rows": preview.fillna("").astype(str).values.tolist(),
            "total_rows": len(df),
        })
    except Exception as exc:
        return jsonify({"error": f"Failed to read sheet: {exc}"}), 400


@app.route("/download-converted")
def download_converted():
    sdir = _session_dir()
    out = os.path.join(sdir, "converted.xlsx")
    if not os.path.exists(out):
        return jsonify({"error": "No converted file found."}), 404
    return send_file(out, as_attachment=True, download_name="converted.xlsx")


if __name__ == "__main__":
    app.run(debug=True)
