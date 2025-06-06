# Excel Score Merger

## Overview
Excel Score Merger is a simple tool for automatically importing exam scores from system-generated Excel files into a master template. It supports any subject: for each course, take the raw output (containing student IDs and scores) and fill them into your predefined template (listing class, student ID, name, score, and notes).

## Features
- Read raw exam output (`BDI302c_FE.xlsx`) with sheet “Result” (columns: No, Login, …, Mark(10)).
- Read a template file (`BDI302c_FINAL-UP_FUMM.xlsx`) that already lists Class, RollNumber, FullName, Mark, Note.
- Match each student’s Login (ID) to RollNumber and copy Mark(10) into the template’s Mark column.
- Leave any unmatched students’ Mark as blank (or zero if you choose).
- Produce a new filled-in Excel file for final reporting.
- Works for any course: just replace the two input files.

## Installation

### Prerequisites
- Python 3.12

### Setup
1. Clone or download this repository to your local machine.  
2. Install required Python packages:
   ```bash
   pip install pandas openpyxl
   ```

## Usage

1. **Prepare your files**  
   - Place your raw exam output (e.g. `BDI302c_FE.xlsx`) in a folder.  
   - Place your template file (e.g. `BDI302c_FINAL-UP_FUMM.xlsx`) in the same folder.  
   - Create an empty folder named `Result` inside that folder to hold the output.

2. **Run the merger**  
   - If you have hard-coded paths in the script (see below), simply run:
     ```bash
     python merge_scores_to_template.py
     ```
   - If you prefer to pass file paths from the command line, modify the `__main__` section to accept arguments, then run:
     ```bash
     python merge_scores_to_template.py        /path/to/BDI302c_FE.xlsx        /path/to/BDI302c_FINAL-UP_FUMM.xlsx        /path/to/Result/BDI302c_FINAL-UP_FUMM_filled.xlsx        Result
     ```
   - The script will read “Result” sheet from `BDI302c_FE.xlsx`, extract Login → Mark(10), and overwrite the Mark column in `BDI302c_FINAL-UP_FUMM.xlsx`. The filled file is saved to `Result/BDI302c_FINAL-UP_FUMM_filled.xlsx`.

## Project Structure

```
excel-score-merger/
├── BDI302c_FE.xlsx
├── BDI302c_FINAL-UP_FUMM.xlsx
├── merge_scores_to_template.py
└── Result/                     # Output folder (empty before running)
    └── BDI302c_FINAL-UP_FUMM_filled.xlsx   # Generated after running
```

- **BDI302c_FE.xlsx**: Raw output from exam system (sheet “Result”).  
- **BDI302c_FINAL-UP_FUMM.xlsx**: Your template with columns [Class, RollNumber, FullName, Mark, Note].  
- **merge_scores_to_template.py**: Python script that performs the merge.  
- **Result/**: Folder where the merged Excel file is saved.

## How It Works

1. **Read raw exam file**  
   - The script opens `BDI302c_FE.xlsx`, looks for a sheet named “Result” (or defaults to the first sheet).  
   - It checks for columns “Login” (student ID) and “Mark(10)” (score).

2. **Prepare lookup data**  
   - It extracts “Login” and “Mark(10)”, renames them to “RollNumber” and “Mark”—forming a small DataFrame.

3. **Read template file**  
   - It opens `BDI302c_FINAL-UP_FUMM.xlsx` (first sheet).  
   - It verifies that the sheet contains columns: Class, RollNumber, Mark, Note.

4. **Map scores into template**  
   - It builds a dictionary mapping each RollNumber to its score.  
   - It then replaces the template’s Mark column by looking up each RollNumber in that dictionary.  
   - Any RollNumber not found in the raw data remains blank (or zero if you enable fillna).

5. **Save the result**  
   - The final DataFrame (with updated Mark) is written to `Result/BDI302c_FINAL-UP_FUMM_filled.xlsx` without altering other columns.

## Dependencies
- **pandas**: for reading/writing and manipulating Excel files  
- **openpyxl**: Excel engine used by pandas  

Install them via:
```bash
pip install pandas openpyxl
python main.py
```

## License
This project is provided as-is. Feel free to reuse or adapt it for your own courses. No explicit license specified.

## Contributing
Feel free to submit pull requests or open issues if you encounter any bugs or have suggestions.
