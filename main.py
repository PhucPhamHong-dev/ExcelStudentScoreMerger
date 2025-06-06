import pandas as pd

def merge_scores(
    raw_exam_path: str,
    template_path: str,
    output_path: str,
    raw_sheet_name: str = "Result",
    template_sheet_name: str = None
):
    try:
        raw_df = pd.read_excel(raw_exam_path, sheet_name=raw_sheet_name)
    except ValueError:
        raw_df = pd.read_excel(raw_exam_path, sheet_name=0)

    if "Login" not in raw_df.columns or "Mark(10)" not in raw_df.columns:
        raise KeyError(
            f"❌ File raw exam phải có cột 'Login' và 'Mark(10)'.\n"
            f"  Các cột hiện có: {list(raw_df.columns)}"
        )

    raw_marks = raw_df[["Login", "Mark(10)"]].copy()
    raw_marks.columns = ["RollNumber", "Mark"]

    if template_sheet_name is None:
        template_df = pd.read_excel(template_path)
    else:
        try:
            template_df = pd.read_excel(template_path, sheet_name=template_sheet_name)
        except ValueError:
            template_df = pd.read_excel(template_path, sheet_name=0)

    required_cols = ["Class", "RollNumber", "Mark", "Note"]
    missing_cols = [col for col in required_cols if col not in template_df.columns]
    if missing_cols:
        raise KeyError(
            f"❌ File template phải có các cột: {required_cols}\n"
            f"  Thiếu: {missing_cols}\n"
            f"  Các cột hiện tại: {list(template_df.columns)}"
        )

    mark_lookup = raw_marks.set_index("RollNumber")["Mark"].to_dict()
    template_df["Mark"] = template_df["RollNumber"].map(mark_lookup)
    template_df.to_excel(output_path, index=False)
    print(f"✅ Merged file has been saved to: {output_path}")

if __name__ == "__main__":
    merge_scores(
        raw_exam_path=r"C:\Users\patmy\Downloads\Excel_For_Mark\BDI302c_FE.xlsx",
        template_path=r"C:\Users\patmy\Downloads\Excel_For_Mark\BDI302c_FINAL-UP_FUMM.xlsx",
        output_path=r"C:\Users\patmy\Downloads\Excel_For_Mark\Result\BDI302c_FINAL-UP_FUMM_filled.xlsx",
        raw_sheet_name="Result"
    )
