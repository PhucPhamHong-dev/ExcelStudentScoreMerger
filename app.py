import os
import zipfile
from tempfile import TemporaryDirectory
from io import BytesIO
from flask import Flask, request, send_file, render_template, abort
import pandas as pd
from collections import defaultdict

# Initialize Flask application: serve static files from 'static/', templates from 'templates/'
app = Flask(__name__, static_folder='static', template_folder='templates')
app.config['SECRET_KEY'] = 'change_this_to_secure'  # Replace with a strong secret key in production

@app.route('/', methods=['GET'])
def form():
    """
    Display the upload form (index.html should be inside the 'templates/' directory)
    """
    return render_template('index.html')

@app.route('/', methods=['POST'])
def index():
    """
    Handle file uploads for raw exam data and template files,
    generate filled Excel files named by exam code with suffixes for duplicates,
    package them into a ZIP in memory, and send to client.
    All on-disk temporary files are cleaned up automatically.
    """
    raw_file = request.files.get('raw_exam')
    tpl_files = request.files.getlist('templates')

    # Validate raw exam file
    if not raw_file or not raw_file.filename.lower().endswith(('.xls', '.xlsx')):
        return render_template('index.html', error='Vui lòng tải lên file Điểm Gốc hợp lệ (.xls/.xlsx)')
    # Validate at least one template is provided
    if not tpl_files:
        return render_template('index.html', error='Vui lòng chọn ít nhất một file template')

    # Use TemporaryDirectory for intermediate files
    with TemporaryDirectory() as tmp_out:
        counts = defaultdict(int)
        stored_codes = []

        # Process each template: extract exam code and filled DataFrame
        for tpl in tpl_files:
            if not tpl.filename.lower().endswith('.xlsx'):
                continue
            code_only, df_out = process_template(raw_file, tpl)
            counts[code_only] += 1
            suffix = f"({counts[code_only]})" if counts[code_only] > 1 else ''
            filename = f"{code_only}_filled{suffix}.xlsx"
            stored_codes.append(code_only)

            # Write filled Excel named by exam code with suffix
            filled_path = os.path.join(tmp_out, filename)
            df_out.to_excel(filled_path, index=False)

        # Create ZIP archive on disk
        disk_zip_path = os.path.join(tmp_out, 'archive.zip')
        with zipfile.ZipFile(disk_zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for fn in os.listdir(tmp_out):
                if fn.lower().endswith('_filled.xlsx'):
                    zf.write(os.path.join(tmp_out, fn), arcname=fn)

        # Load ZIP into memory
        mem_zip = BytesIO()
        with open(disk_zip_path, 'rb') as f:
            mem_zip.write(f.read())
        mem_zip.seek(0)

        # Determine download filename: if single distinct code, use that code
        distinct = set(stored_codes)
        if len(distinct) == 1:
            # use code with _filled.zip
            download_name = f"{list(distinct)[0]}_filled.zip"
        else:
            download_name = 'FE_Merge.zip'

        return send_file(
            mem_zip,
            mimetype='application/zip',
            as_attachment=True,
            download_name=download_name
        )


def process_template(raw_file_storage, tpl_file_storage):
    """
    Read raw exam and template into DataFrames,
    extract exam code, merge data, and return
    tuple (exam_code, result DataFrame).
    """
    # Load raw and template data
    df_raw = pd.read_excel(raw_file_storage, dtype=str)
    df_tpl = pd.read_excel(tpl_file_storage, dtype=str)

    # Standardize raw data
    df_raw.rename(columns={'GroupName': 'Class'}, inplace=True)
    df_raw['RollNumber']  = df_raw['RollNumber'].str.upper().str.strip()
    df_raw['SubjectCode'] = df_raw['SubjectCode'].str.upper().str.strip()
    df_raw['SlotType']    = df_raw.get('SlotType', '').fillna('')

    # Trim template headers
    df_tpl.rename(columns=lambda c: c.strip(), inplace=True)
    required = {'Exam Code', 'Login', 'Mark(10)'}
    if not required.issubset(df_tpl.columns):
        abort(400, description='Template thiếu cột bắt buộc: ' + ', '.join(required))

    # Extract exam code
    exam_codes = (
        df_tpl['Exam Code']
        .dropna()
        .astype(str)
        .str.strip()
        .str.upper()
    )
    if exam_codes.empty:
        abort(400, description="Không tìm thấy giá trị trong cột 'Exam Code'")
    code_full = exam_codes.iloc[0]
    code_only = code_full.split('_')[0]

    # Merge on RollNumber vs Login
    df_tpl['Login'] = df_tpl['Login'].str.upper().str.strip()
    merged = pd.merge(
        df_raw[df_raw['SubjectCode'] == code_only],
        df_tpl,
        left_on='RollNumber', right_on='Login', how='right'
    )

    # Build output DataFrame
    df_out = pd.DataFrame({
        'Class':      merged['Class'],
        'RollNumber': merged['RollNumber'],
        'FullName':   merged.get('FullName', merged.get('Login', '')),
        'Mark':       merged['Mark(10)'],
        'Note':       merged['SlotType'].replace('', pd.NA)
    })
    return code_only, df_out

# Run Flask app
if __name__ == '__main__':
    app.run(debug=True)
