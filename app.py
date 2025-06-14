from flask import Flask, request, send_file, flash, render_template
import os, tempfile, shutil, zipfile
import pandas as pd

app = Flask(__name__)
app.config['SECRET_KEY'] = 'change_this_to_secure'
UPLOAD_FOLDER = 'temp'
OUTPUT_FOLDER = 'output'

# Kiểm tra file Raw Exam
ALLOWED_RAW_EXT = ('.xls', '.xlsx')
def allowed_raw(name):
    return name.lower().endswith(ALLOWED_RAW_EXT)

@app.route('/', methods=['GET', 'POST'])
def index():
    error = None
    if request.method == 'POST':
        raw = request.files.get('raw_exam')
        tpl_files = request.files.getlist('templates')

        if not raw or not allowed_raw(raw.filename):
            error = 'Vui lòng tải lên file Điểm Gốc (.xlsx)'
            return render_template('index.html', error=error)
        if not tpl_files:
            error = 'Vui lòng chọn ít nhất một file Templates'
            return render_template('index.html', error=error)

        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        # Lưu file tổng
        raw_path = os.path.join(UPLOAD_FOLDER, 'raw_exam.xlsx')
        raw.save(raw_path)

        # Đọc & chuẩn hóa file tổng
        df = pd.read_excel(raw_path, dtype=str)
        df.rename(columns={'GroupName': 'Class'}, inplace=True)
        df['RollNumber'] = df['RollNumber'].str.upper()
        df['SubjectCode'] = df['SubjectCode'].str.upper()
        df['SlotType'] = df['SlotType'].fillna('')

        # Lưu tất cả template vào thư mục tạm
        tempdir = tempfile.mkdtemp()
        codes_available = set()
        saved_paths = []
        for f in tpl_files:
            if not f.filename.lower().endswith('.xlsx'):
                continue
            name = os.path.basename(f.filename)
            dest = os.path.join(tempdir, name)
            f.save(dest)
            saved_paths.append(dest)
            codes_available.add(name.split('_')[0].upper())

        # Báo thiếu template nhưng vẫn tiếp tục
        needed = set(df['SubjectCode'].unique())
        missing = needed - codes_available
        if missing:
            flash('Thiếu template cho môn: ' + ', '.join(sorted(missing)), 'warning')

        # Xử lý từng template có sẵn
        for path_tpl in saved_paths:
            code = os.path.basename(path_tpl).split('_')[0].upper()
            sub = df[df['SubjectCode'] == code]
            if sub.empty:
                flash(f'Không có dữ liệu môn {code} trong file tổng', 'warning')
                continue

            # Mở và đóng ExcelFile ngay khi parse xong
            sheet_df = None
            with pd.ExcelFile(path_tpl) as xls:
                for sh in xls.sheet_names:
                    tmp = xls.parse(sh, dtype=str)
                    cols = [c.strip() for c in tmp.columns]
                    if 'Login' in cols and 'Mark(10)' in cols:
                        sheet_df = tmp.copy()
                        break

            if sheet_df is None:
                flash(f'Không tìm thấy sheet Login & Mark(10) trong {os.path.basename(path_tpl)}', 'warning')
                continue

            sheet_df['Login'] = sheet_df['Login'].str.upper()
            merged = pd.merge(
                sub,
                sheet_df,
                left_on='RollNumber',
                right_on='Login',
                how='right'
            )
            out = pd.DataFrame({
                'Class': merged['Class'],
                'RollNumber': merged['RollNumber'],
                'FullName': merged['FullName'],
                'Mark': merged['Mark(10)'],
                'Note': merged['SlotType'].replace('', pd.NA)
            })

            out_name = f"{code}_filled.xlsx"
            out.to_excel(os.path.join(OUTPUT_FOLDER, out_name), index=False)
            flash(f'Đã tạo {out_name}', 'success')

        # Dọn dẹp thư mục tạm
        shutil.rmtree(tempdir, ignore_errors=True)

        # Nén kết quả
        zip_out = os.path.join(OUTPUT_FOLDER, 'FE_Merge.zip')
        with zipfile.ZipFile(zip_out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for fn in os.listdir(OUTPUT_FOLDER):
                if fn.endswith('_filled.xlsx'):
                    zout.write(os.path.join(OUTPUT_FOLDER, fn), arcname=fn)

        return send_file(zip_out, as_attachment=True)

    return render_template('index.html', error=error)

if __name__ == '__main__':
    app.run(debug=True)
