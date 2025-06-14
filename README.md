# Merge Exam Mark Web Application

## Overview

This is a Flask-based web application for merging raw exam scores into standardized Excel templates. The project features a modern, user-friendly front-end built with TailwindCSS and AlpineJS, and client-side validation using SheetJS. Raw score data and multiple subject templates are uploaded through a clean interface; the app then produces filled result files packaged as a single ZIP.

## Features

- **User-Friendly Interface**  
  - TailwindCSS styling for a responsive and professional look.  
  - AlpineJS for dynamic UI interactions: file pickers, real-time status banners, and error handling.

- **Flexible File Upload**  
  - Upload a single **Raw Exam** `.xlsx` file.  
  - Upload multiple **Template** `.xlsx` files or an entire folder.  
  - Ability to add or remove templates before processing.

- **Client-Side Validation**  
  - **SheetJS** parses the Raw Exam file in the browser to extract all `SubjectCode` values.  
  - Real-time detection of **missing** or **extra** template files, with warnings displayed before submission.  
  - “Ready” banner appears when all required templates are present.

- **Server-Side Processing**  
  - Flask handles file uploads and serves the HTML form.  
  - `pandas` + `openpyxl` perform the merge logic:
    1. Normalize raw data: rename `GroupName` → `Class`, uppercase `RollNumber`.  
    2. For each `SubjectCode`, match `RollNumber` to the `Login` column in the template.  
    3. Output only five columns: `Class`, `RollNumber`, `FullName`, `Mark`, `Note`.  
  - All filled templates are saved and zipped into `FE_Merge.zip` for download.

## Installation

### Prerequisites

- Python 3.8+  
- `pip`

### Setup

1. **Clone the repository**  
   ```bash
   git clone https://github.com/your-username/exam-merge-app.git
   cd exam-merge-app
   ```

2. **Create and activate a virtual environment**  
   ```bash
   python -m venv venv
   # Windows
   venv\Scripts\activate
   # macOS/Linux
   source venv/bin/activate
   ```

3. **Install dependencies**  
   ```bash
   pip install Flask pandas openpyxl
   ```

## Project Structure

```
exam-merge-app/
├── app.py
├── templates/
│   └── index.html
├── static/
│   └── (optional CSS/JS assets)
├── README.md
└── requirements.txt
```

- **app.py**: Flask backend and merge logic  
- **templates/index.html**: Front-end form with TailwindCSS, AlpineJS, and SheetJS  

## Running the Application

```bash
python app.py
```

- The server runs at `http://localhost:5000`.  
- Open the URL in your browser, upload the **Raw Exam** file and select your **Template** files or folder.  
- The UI will validate template coverage before enabling the **Create Result** button.  
- Click the button to download `FE_Merge.zip`, containing all merged Excel files.

## License

This project is provided “as-is” for educational and personal use.
