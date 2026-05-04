import pandas as pd
from pathlib import Path

# Folder containing Excel files
DATA_DIR = Path("data")   # change this

# Get only grade files
excel_files = [
    f for f in DATA_DIR.glob("*.xlsx")
    if f.name.lower().startswith("grade_")
]

for file in excel_files:

    print(f"Processing {file.name}...")

    xls = pd.ExcelFile(file)

    with pd.ExcelWriter(
        file,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:

        # Process only _formatted sheets
        formatted_sheets = [
            s for s in xls.sheet_names
            if s.endswith("_formatted")
        ]

        for sheet in formatted_sheets:

            print(f"  Creating pivots for {sheet}...")

            df = pd.read_excel(file, sheet_name=sheet)

            # Ensure numeric
            df["Performance (%)"] = pd.to_numeric(
                df["Performance (%)"],
                errors="coerce"
            )

            df["Total Credit"] = pd.to_numeric(
                df["Total Credit"],
                errors="coerce"
            )

            # -------------------------------------------------
            # SCHOOL-WISE SUBJECT-WISE PIVOT
            # -------------------------------------------------

            school_pivot = pd.pivot_table(
                df,
                index=["District", "SchoolName"],
                columns="Subject",
                values="Performance (%)",
                aggfunc=["mean", "count"],
            )

            school_pivot.columns = [
                f"{agg}_{subject}"
                for agg, subject in school_pivot.columns
            ]

            school_pivot = school_pivot.reset_index()

            school_sheet_name = sheet.replace(
                "_formatted",
                "_school_subject_pivot"
            )

            school_pivot.to_excel(
                writer,
                sheet_name=school_sheet_name,
                index=False
            )

            # -------------------------------------------------
            # DISTRICT-WISE SUBJECT-WISE PIVOT
            # -------------------------------------------------

            district_pivot = pd.pivot_table(
                df,
                index="District",
                columns="Subject",
                values="Performance (%)",
                aggfunc=["mean", "count"],
            )

            district_pivot.columns = [
                f"{agg}_{subject}"
                for agg, subject in district_pivot.columns
            ]

            district_pivot = district_pivot.reset_index()

            district_sheet_name = sheet.replace(
                "_formatted",
                "_district_subject_pivot"
            )

            district_pivot.to_excel(
                writer,
                sheet_name=district_sheet_name,
                index=False
            )

    print(f"Finished {file.name}")

print("All pivots created successfully.")
