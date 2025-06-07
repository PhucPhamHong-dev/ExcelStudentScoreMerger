import os
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

app = Flask(__name__)


def is_processed_excel(file_stream):
    """
    Check for our hidden marker sheet '_processed_marker'.
    """
    file_stream.seek(0)
    try:
        wb = load_workbook(file_stream, read_only=True)
        return "_processed_marker" in wb.sheetnames
    except:
        return False
    finally:
        file_stream.seek(0)


def merge_scores(raw_df, template_df):
    # Validate raw file
    if "Login" not in raw_df.columns or "Mark(10)" not in raw_df.columns:
        raise KeyError("Raw file phải có cột 'Login' và 'Mark(10)'")
    raw_marks = raw_df[["Login", "Mark(10)"]].copy()
    raw_marks.columns = ["RollNumber", "Mark"]
    lookup = raw_marks.set_index("RollNumber")["Mark"].to_dict()

    # Validate template
    required = ["Class", "RollNumber", "Mark", "Note"]
    missing = [c for c in required if c not in template_df.columns]
    if missing:
        raise KeyError(f"Template thiếu cột: {missing}")

    # Merge
    template_df["Mark"] = template_df["RollNumber"].map(lookup)
    return template_df


@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    if request.method == "POST":
        raw_file = request.files.get("raw_exam")
        template_file = request.files.get("template")
        if not raw_file or not template_file:
            error = "Vui lòng chọn cả hai file."
            return render_template("index.html", error=error)

        # 1) Check filename suffix
        def has_processed_suffix(fname):
            base, _ = os.path.splitext(fname.lower())
            return base.endswith(("_filled","_merged"))
        if has_processed_suffix(raw_file.filename):
            error = "File Raw có vẻ đã được xử lý rồi. Vui lòng chọn file gốc."
            return render_template("index.html", error=error)
        if has_processed_suffix(template_file.filename):
            error = "File Template có vẻ đã được xử lý rồi. Vui lòng chọn file gốc."
            return render_template("index.html", error=error)

        # 2) Check hidden marker sheet
        if is_processed_excel(raw_file.stream):
            error = "File Raw có dấu hiệu đã xử lý rồi. Vui lòng chọn file gốc."
            return render_template("index.html", error=error)
        if is_processed_excel(template_file.stream):
            error = "File Template có dấu hiệu đã xử lý rồi. Vui lòng chọn file gốc."
            return render_template("index.html", error=error)

        # Read Excel files
        try:
            raw_df = pd.read_excel(raw_file, sheet_name="Result", engine="openpyxl")
        except:
            raw_df = pd.read_excel(raw_file, engine="openpyxl")
        template_df = pd.read_excel(template_file, engine="openpyxl")

        # Merge
        try:
            result_df = merge_scores(raw_df, template_df)
        except Exception as e:
            error = str(e)
            return render_template("index.html", error=error)

        # Write output with hidden marker sheet
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result_df.to_excel(writer, index=False)
            wb = writer.book
            marker = wb.create_sheet("_processed_marker")
            marker.sheet_state = "veryHidden"
        output.seek(0)

        # Construct a user‑friendly filename based on template
        template_name = secure_filename(template_file.filename)
        base, ext = os.path.splitext(template_name)
        output_name = f"{base}_filled{ext}"

        return send_file(
            output,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    return render_template("index.html", error=error)


if __name__ == "__main__":
    app.run(debug=True)