# 📊 Merge Exam Mark Web Application


---

## 🚀 Overview

A **Flask**-based web application that merges raw exam scores into standardized Excel templates. Featuring a responsive front-end built with **TailwindCSS** and **AlpineJS**, plus client-side validation via **SheetJS**, it delivers filled result files in a clean ZIP archive.

---

## 📑 Table of Contents

* [Features](#-features)
* [Installation](#-installation)
* [Project Structure](#-project-structure)
* [Usage](#-usage)
* [Changelog](#-changelog)
* [License](#-license)

---

## ✨ Features

1. **User-Friendly Interface**

   * Responsive styling with TailwindCSS.
   * Dynamic file pickers, status banners, and error messages powered by AlpineJS.

2. **Flexible File Upload**

   * Single **Raw Exam** `.xlsx` upload.
   * Multiple **Template** `.xlsx` files or folder import.
   * Add/remove templates before processing.

3. **Client-Side Validation**

   * Extract `SubjectCode` values in-browser with SheetJS.
   * Real-time warnings for **missing** or **extra** templates.
   * “Ready” banner when all required templates are detected.

4. **Server-Side Processing**

   * Flask handles uploads and serves the web form.
   * **Pandas** + **OpenPyXL** merge logic:

     1. Normalize raw data (`GroupName` → `Class`, uppercase/trimming fields).
     2. Match `RollNumber` to `Login` in each template.
     3. Output columns: `Class`, `RollNumber`, `FullName`, `Mark`, `Note`.
   * Uses a single **TemporaryDirectory** for all file IO—no leftover folders.
   * Auto-deduplicates filenames by appending `(2)`, `(3)`, etc. for repeated `exam_code`.
   * Zips **only** the generated `_filled.xlsx` results (inputs excluded).
   * Dynamic ZIP naming:

     * Single exam: `<exam_code>_filled.zip`
     * Multiple exams: `FE_Merge.zip`

---

## 🛠️ Installation

### Prerequisites

* Python **3.8** or higher
* **pip** package manager

### Setup Steps

```bash
# Clone repository
git clone https://github.com/your-username/exam-merge-app.git
cd exam-merge-app

# Create & activate virtual environment
python -m venv venv
# Windows: venv\Scripts\activate
# macOS/Linux: source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

---

## 📂 Project Structure

```text
exam-merge-app/
├── app.py                 # Flask backend & merge logic
├── requirements.txt       # Python dependencies
├── templates/             # Jinja2 HTML templates
│   └── index.html         # Main upload form
├── static/                # CSS/JS assets (Tailwind, Alpine, SheetJS)
└── CHANGELOG.md           # Version history & updates
```

---

## 🎬 Usage

1. **Start the server**:

   ```bash
   python app.py
   ```
2. Open your browser at `http://localhost:5000`.
3. Upload your **Raw Exam** file (.xlsx).
4. Select one or more **Template** files (or folder).
5. Confirm the “Ready” banner, then click **Tạo file kết quả**.
6. Download the generated ZIP archive.

---

## 📖 Changelog

See [CHANGELOG.md](CHANGELOG.md) for details on recent updates.

---

## 📜 License

This project is provided **as-is** for educational and personal use. No warranty is provided.
