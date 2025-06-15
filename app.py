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
    Handle file uploads for raw exam data and template files.
    Use a single TemporaryDirectory for all temporary storage (raw and filled files).
    Generate filled Excel files named by exam code with suffixes for duplicates,
    package them into an in-memory ZIP, and send to the client.
    All temporary files are cleaned up automatically when done.
    """
    raw_file = request.files.get('raw_exam')
    tpl_files = request.files.getlist('templates')

    # Validate raw exam file
    if not raw_file or not raw_file.filename.lower().endswith(('.xls', '.xlsx')):
        return render_template('index.html', error='Vui lòng tải lên file Điểm Gốc hợp lệ (.xls/.xlsx)')
    # Validate at least one template
    if not tpl_files:
        return render_template('index.html', error='Vui lòng chọn ít nhất một file template')

    # Use one TemporaryDirectory for both reading raw and storing outputs
    with TemporaryDirectory() as workdir:
        counts = defaultdict(int)
        stored_codes = []

        # Save raw file into temporary directory for consistent pandas reading
        raw_path = os.path.join(workdir, 'raw_exam.xlsx')
        raw_file.save(raw_path)

        # Process each template
        for tpl in tpl_files:
            if not tpl.filename.lower().endswith('.xlsx'):
                continue

            tpl_path = os.path.join(workdir, tpl.filename)
            tpl.save(tpl_path)

            # process_template now takes file paths as input
            code_only, df_out = process_template(raw_path, tpl_path)
            counts[code_only] += 1
            suffix = f"({counts[code_only]})" if counts[code_only] > 1 else ''
            filename = f"{code_only}_filled{suffix}.xlsx"
            stored_codes.append(code_only)

            # write filled file
            df_out.to_excel(os.path.join(workdir, filename), index=False)

        # create ZIP archive on disk within temporary directory
        disk_zip = os.path.join(workdir, 'result.zip')
        with zipfile.ZipFile(disk_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
            for fn in os.listdir(workdir):
                if fn.lower().endswith('_filled.xlsx'):
                    zf.write(os.path.join(workdir, fn), arcname=fn)

        # load ZIP into memory buffer
        mem_zip = BytesIO()
        with open(disk_zip, 'rb') as f:
            mem_zip.write(f.read())
        mem_zip.seek(0)

        # choose download name
        distinct = set(stored_codes)
        download_name = f"{list(distinct)[0]}_filled.zip" if len(distinct)==1 else 'FE_Merge.zip'

        return send_file(
            mem_zip,
            mimetype='application/zip',
            as_attachment=True,
            download_name=download_name
        )


def process_template(raw_path, tpl_path):
    """
    Read raw exam and template from file paths,
    extract exam code, merge data, and return (code, DataFrame).
    """
    df_raw = pd.read_excel(raw_path, dtype=str)
    df_tpl = pd.read_excel(tpl_path, dtype=str)

    # standardize raw
    df_raw.rename(columns={'GroupName':'Class'}, inplace=True)
    df_raw['RollNumber'] = df_raw['RollNumber'].str.upper().str.strip()
    df_raw['SubjectCode'] = df_raw['SubjectCode'].str.upper().str.strip()
    df_raw['SlotType'] = df_raw.get('SlotType','').fillna('')

    # trim template headers
    df_tpl.rename(columns=lambda c: c.strip(), inplace=True)
    required = {'Exam Code','Login','Mark(10)'}
    if not required.issubset(df_tpl.columns):
        abort(400, description='Template thiếu cột bắt buộc: ' + ', '.join(required))

    # extract code
    codes = df_tpl['Exam Code'].dropna().astype(str).str.strip().str.upper()
    if codes.empty:
        abort(400, description="Không tìm thấy giá trị trong cột 'Exam Code'")
    code_full = codes.iloc[0]
    code_only = code_full.split('_')[0]

    # merge
    df_tpl['Login'] = df_tpl['Login'].str.upper().str.strip()
    merged = pd.merge(
        df_raw[df_raw['SubjectCode']==code_only],
        df_tpl,
        left_on='RollNumber', right_on='Login', how='right'
    )

    df_out = pd.DataFrame({
        'Class':merged['Class'],
        'RollNumber':merged['RollNumber'],
        'FullName':merged.get('FullName', merged.get('Login','')),
        'Mark':merged['Mark(10)'],
        'Note':merged['SlotType'].replace('', pd.NA)
    })
    return code_only, df_out

# Run Flask app
if __name__ == '__main__':
    app.run(debug=True)
