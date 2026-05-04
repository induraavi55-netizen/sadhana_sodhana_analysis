import pandas as pd
from pathlib import Path

input_file = Path("data/blanks_checking.xlsx")

xls = pd.ExcelFile(input_file)

allowed_sheets = {"english", "mathematics", "science"}

with pd.ExcelWriter(
    input_file,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace",
) as writer:

    for sheet in xls.sheet_names:

        if sheet.strip().lower() not in allowed_sheets:
            continue

        df = pd.read_excel(xls, sheet_name=sheet)

        # Keep Student LoginId column
        student_cols = [
            c for c in df.columns
            if c.strip().lower() == "student loginid"
        ]

        # Keep response_marked columns
        response_cols = [
            c for c in df.columns if c.endswith("_response_marked")
        ]

        if not response_cols:
            print(f"No response columns in {sheet}")
            continue

        cols_to_keep = student_cols + response_cols

        df_filtered = df[cols_to_keep]

        df_filtered.to_excel(writer, sheet_name=sheet, index=False)

        print(f"Cleaned: {sheet}")

print("Done.")