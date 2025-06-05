import pandas as pd
import sys

def main(scores_file, students_file, output_file):
    """
    Merge two sources (scores and students) by RollNumber to create a final Excel sheet.

    Arguments:
    - scores_file:   Path to an Excel file containing exactly 2 columns: [RollNumber, Mark].
    - students_file: Path to an Excel file containing at least 3 columns: [Class, RollNumber, FullName].
    - output_file:   Desired path/name of the resulting Excel file with columns [Class, RollNumber, FullName, Mark, Note].
    """

 
    scores_df   = pd.read_excel(scores_file)
    students_df = pd.read_excel(students_file)

    scores_df = scores_df.iloc[:, 0:2].copy()
    scores_df.columns = ["RollNumber", "Mark"]

   
    students_df = students_df.iloc[:, 0:3].copy()
    students_df.columns = ["Class", "RollNumber", "FullName"]

    
    merged = pd.merge(
        students_df,
        scores_df,
        on="RollNumber",
        how="left"  # Keeps all rows from students_df; if no matching score, Mark will be NaN
    )

  
    merged["Note"] = ""
    merged = merged[["Class", "RollNumber", "FullName", "Mark", "Note"]]

   
    merged.to_excel("/mnt/data/BDI302c_FINAL-UP_FUMM.xlsx", index=False, sheet_name="BDI302c_FINAL-UP_FUMM")
    print(f"➡️ File merged has been saved to: ")


if __name__ == "__main__":
    
    if len(sys.argv) != 4:
        print("❌ Usage: python merge_excel.py <scores_file.xlsx> <students_file.xlsx> <output_file.xlsx>")
        sys.exit(1)

    scores_path   = sys.argv[1]
    students_path = sys.argv[2]
    output_path   = sys.argv[3]
    main(scores_path, students_path, output_path)
