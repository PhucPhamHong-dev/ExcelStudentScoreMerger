# Merge Exam Scores

A simple Flask web app that merges raw exam scores from a system‑exported Excel file into your predefined Excel template and produces a filled‑in result file.

---

## 🚀 Features

- **Web upload form**  
  Upload both your raw exam file and your template file via a clean TailwindCSS interface.

- **Double “already processed” guard**  
  1. Rejects files whose name ends with `_filled` or `_merged`.  
  2. Rejects any Excel containing a hidden sheet `_processed_marker`.

- **Automatic merge logic**  
  - Reads raw data from the sheet named **Result** (or the first sheet if “Result” is absent).  
  - Extracts `Login` and `Mark(10)` columns, renames them to `RollNumber` & `Mark`.  
  - Reads your template (must contain `Class`, `RollNumber`, `Mark`, `Note`) and maps each `RollNumber` → `Mark`.

- **Hidden marker sheet**  
  Outputs a sheet called `_processed_marker` (veryHidden) to prevent re‑processing the same file.

- **User‑friendly output**  
  Downloads a file named `<template_basename>_filled.xlsx` ready for final reporting.

---

## 🔧 Installation

### Prerequisites

- Python 3.8 or higher  
- pip

### Steps

1. **Clone this repository**  
   ```bash
   git clone https://github.com/your‑username/excel-score-merger.git
   cd excel-score-merger
   ```

2. **Create & activate a virtual environment**  
   ```bash
   python -m venv venv
   # Windows
   venv\Scripts\activate
   # macOS / Linux
   source venv/bin/activate
   ```

3. **Install dependencies**  
   ```bash
   pip install Flask pandas openpyxl
   ```

---

## 🗂️ Project Structure

```
excel-score-merger/
├── app.py
├── templates/
│   └── index.html      # Upload form (TailwindCSS)
├── requirements.txt    # (optional) pin your dependencies here
└── README.md           # You’re reading it!
```

- **app.py** – Main Flask application  
- **templates/index.html** – HTML form for file uploads  
- **requirements.txt** – (optional) list of Python packages

---

## ▶️ Running the App

1. Activate your virtual environment (if not already).  
2. Start the Flask server:
   ```bash
   python app.py
   ```
3. Open your browser at `http://localhost:5000`.  
4. Upload **Raw Exam** and **Template** files, then click **Tạo file kết quả** to download the merged Excel.

---

## ⚙️ How It Works

1. **Upload & Validation**  
   - Ensures both files are provided.  
   - Checks filenames do not end with `_filled`/`_merged`.  
   - Scans for a hidden `_processed_marker` sheet.

2. **Read Excel Data**  
   - **Raw file**: tries sheet `"Result"`, falling back to the first sheet.  
   - **Template file**: always reads the first sheet.

3. **Merge Logic**  
   ```python
   raw_marks = raw_df[["Login", "Mark(10)"]].copy()
   raw_marks.columns = ["RollNumber", "Mark"]
   lookup = raw_marks.set_index("RollNumber")["Mark"].to_dict()
   template_df["Mark"] = template_df["RollNumber"].map(lookup)
   ```

4. **Write Output**  
   - Writes the merged DataFrame to a new Excel.  
   - Creates a veryHidden sheet `_processed_marker`.  
   - Returns the file named `<template_basename>_filled.xlsx`.

---

## 📦 Dependencies

- **Flask** – web framework  
- **pandas** – data manipulation & Excel read/write  
- **openpyxl** – Excel engine for pandas  

Install with:
```bash
pip install Flask pandas openpyxl
```

---

## 🤝 Contributing

Contributions, issues and feature requests are welcome!  
1. Fork the repo  
2. Create a feature branch (`git checkout -b feature/YourFeature`)  
3. Commit your changes (`git commit -m 'Add some feature'`)  
4. Push to the branch (`git push origin feature/YourFeature`)  
5. Open a Pull Request

---

## 📄 License

This project is provided “as‑is” under no specific license. Feel free to copy, adapt, and use for personal or educational purposes.
