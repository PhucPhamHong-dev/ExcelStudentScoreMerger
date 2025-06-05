# Excel Data Merge: Students & Scores

## Description
This repository contains an example Python script that merges two Excel sources into a single output file with the following columns:

- **Class**  
- **RollNumber** (Student ID)  
- **FullName** (Student’s full name)  
- **Mark** (Score on a 10-point scale)  
- **Note** (Empty column for any additional remarks)

The two input sources can be either:
1. Two separate Excel files:  
   - `scores.xlsx`: contains **2 columns** [RollNumber, Mark].  
   - `students.xlsx`: contains **3 columns** [Class, RollNumber, FullName].  

2. One Excel file (`test_data.xlsx`) with **two sheets**:  
   - Sheet **Scores** (2 columns: RollNumber, Mark)  
   - Sheet **Students** (3 columns: Class, RollNumber, FullName)

Running the script will produce `BDI302c_FINAL-UP_FUMM.xlsx` (or any output name you choose) with exactly five columns:  

---

## Requirements
- **Python 3.x**  
- Python packages:
  ```bash
  pip install pandas openpyxl
# ExcelStudentScoreMerger
