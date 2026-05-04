import pandas as pd
from pathlib import Path

DATA_DIR = Path("data")   # change folder if needed


excel_files = [
    f for f in DATA_DIR.glob("*.xlsx")
    if f.name.lower().startswith("grade_")
]

for input_file in excel_files:

    print(f"Updating performance metrics in {input_file.name}...")

    xls = pd.ExcelFile(input_file)

    pivot_rows = []   # store summary for this file

    with pd.ExcelWriter(
        input_file,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:

        for sheet in xls.sheet_names:

            # Only operate on _formatted sheets
            if not sheet.endswith("_formatted"):
                continue

            df = pd.read_excel(input_file, sheet_name=sheet)

            credit_cols = [c for c in df.columns if c.endswith("_credit")]

            if not credit_cols:
                continue

            # ---- calculations ----
            df["Total Credit"] = df[credit_cols].sum(axis=1)

            df["Performance (%)"] = (
                df["Total Credit"] / len(credit_cols) * 100
            ).round(0).astype(int)

            # Write back updated formatted sheet
            df.to_excel(writer, sheet_name=sheet, index=False)

            # ---- collect pivot info ----
            base_name = sheet.replace("_formatted", "")
            avg_perf = round(df["Performance (%)"].mean(), 0)

            pivot_rows.append(
                {
                    "Sheet": base_name,
                    "Average Performance (%)": int(avg_perf),
                }
            )

        # ---------- WRITE PIVOT SHEET ----------
        if pivot_rows:
            pivot_df = pd.DataFrame(pivot_rows)

            pivot_df.to_excel(writer, sheet_name="Sub_wise_avg_perf", index=False)

    print(f"Finished {input_file.name}")

print("All _formatted sheets updated and pivot sheet created.")
