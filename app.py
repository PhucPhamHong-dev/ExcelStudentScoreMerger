from flask import Flask, request, send_file, flash, render_template
import os, tempfile, shutil, zipfile
import pandas as pd

app = Flask(__name__)
app.config['SECRET_KEY'] = 'change_this_to_secure'
UPLOAD_FOLDER = 'temp'
OUTPUT_FOLDER = 'output'
ALLOWED_RAW_EXT = ('.xls', '.xlsx')

def allowed_raw(name):
    return name.lower().endswith(ALLOWED_RAW_EXT)

@app.route('/', methods=['GET', 'POST'])
def index():
    error = None
    if request.method == 'POST':
        raw       = request.files.get('raw_exam')
        tpl_files = request.files.getlist('templates')

        # 1) Validate uploads
        if not raw or not allowed_raw(raw.filename):
            error = 'Vui lòng tải lên file Điểm Gốc (.xlsx)'
            return render_template('index.html', error=error)
        if not tpl_files:
            error = 'Vui lòng chọn ít nhất một file Templates'
            return render_template('index.html', error=error)

        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        # 2) Read & normalize raw
        raw_path = os.path.join(UPLOAD_FOLDER, 'raw_exam.xlsx')
        raw.save(raw_path)
        df_raw = pd.read_excel(raw_path, dtype=str)
        df_raw.rename(columns={'GroupName':'Class'}, inplace=True)
        df_raw['RollNumber']  = df_raw['RollNumber'].str.upper()
        df_raw['SubjectCode'] = df_raw['SubjectCode'].str.upper()
        df_raw['SlotType']    = df_raw['SlotType'].fillna('')

        # 3) Extract Exam Code
        tempdir = tempfile.mkdtemp()
        templates_info=[]
        for fs in tpl_files:
            if not fs.filename.lower().endswith('.xlsx'):
                flash(f'⚠️ Bỏ qua file không phải .xlsx: "{fs.filename}"','warning')
                continue
            dest=os.path.join(tempdir,fs.filename)
            fs.save(dest)
            exam_code=None; sheet_df=None
            with pd.ExcelFile(dest) as xls:
                for sh in xls.sheet_names:
                    tmp=xls.parse(sh,dtype=str)
                    cols=[c.strip() for c in tmp.columns]
                    if {'Exam Code','Login','Mark(10)'}.issubset(cols):
                        vals=(tmp['Exam Code']
                                 .dropna()
                                 .astype(str)
                                 .str.strip()
                                 .str.upper())
                        if not vals.empty:
                            rawc=vals.iloc[0]; exam_code=rawc.split('_')[0]; sheet_df=tmp
                        break
            if exam_code is None:
                flash(f'⚠️ Template "{fs.filename}" thiếu cột "Exam Code"/"Login"/"Mark(10)"','warning')
            else:
                templates_info.append({'code':exam_code,'df':sheet_df})

        # 4) missing/extra warnings
        raw_codes=set(df_raw['SubjectCode'])
        tpl_codes=set(t['code'] for t in templates_info)
        missing=sorted(raw_codes - tpl_codes)
        extra=sorted(tpl_codes - raw_codes)
        if missing: flash('⚠️ Thiếu template cho mã môn: '+', '.join(missing),'warning')
        if extra:   flash('⚠️ Có template thừa cho mã môn: '+', '.join(extra),'warning')

        # 5) Merge available only
        for info in templates_info:
            code=info['code']; tpl_df=info['df']
            sub=df_raw[df_raw['SubjectCode']==code]
            if sub.empty:
                flash(f'⚠️ Không có dữ liệu cho mã môn {code}','warning')
                continue
            tpl_df['Login']=tpl_df['Login'].str.upper()
            merged=pd.merge(sub, tpl_df, left_on='RollNumber', right_on='Login', how='right')
            out=pd.DataFrame({
                'Class':merged['Class'],
                'RollNumber':merged['RollNumber'],
                'FullName':merged['FullName'],
                'Mark':merged['Mark(10)'],
                'Note':merged['SlotType'].replace('',pd.NA)
            })
            name=f"{code}_filled.xlsx"
            out.to_excel(os.path.join(OUTPUT_FOLDER,name),index=False)
            flash(f'✅ Đã tạo file: {name}','success')

        # 6) cleanup + zip
        shutil.rmtree(tempdir,ignore_errors=True)
        zip_out=os.path.join(OUTPUT_FOLDER,'FE_Merge.zip')
        with zipfile.ZipFile(zip_out,'w',zipfile.ZIP_DEFLATED) as zf:
            for fn in os.listdir(OUTPUT_FOLDER):
                if fn.endswith('_filled.xlsx'):
                    zf.write(os.path.join(OUTPUT_FOLDER,fn),arcname=fn)

        return send_file(zip_out,as_attachment=True)

    return render_template('index.html', error=error)

if __name__=='__main__':
    app.run(debug=True)
