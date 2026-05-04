import pandas as pd
from pathlib import Path

DATA_DIR = Path("data")  # change to your folder
OUTPUT_FILE = DATA_DIR / "Top_projection_output.xlsx"

projection_records = []

for file_path in DATA_DIR.glob("*.xlsx"):

    print(f"Processing {file_path.name}")

    xls = pd.ExcelFile(file_path)

    # read only _formatted sheets
    formatted_sheets = [
        s for s in xls.sheet_names if s.endswith("_formatted")
    ]

    if not formatted_sheets:
        continue

    df_all = pd.concat(
        [pd.read_excel(file_path, sheet_name=s) for s in formatted_sheets],
        ignore_index=True
    )

    # ensure numeric
    df_all["Performance (%)"] = pd.to_numeric(
        df_all["Performance (%)"],
        errors="coerce"
    )

    # find students who scored 100% in Science
    science_100_ids = df_all[
        (df_all["Subject"].str.lower() == "science") &
        (df_all["Performance (%)"] >= 80)
    ]["Student LoginId"].unique()

    if len(science_100_ids) == 0:
        continue

    df_proj = df_all[
        df_all["Student LoginId"].isin(science_100_ids)
    ]

    # pivot to wide format
    projection = df_proj.pivot_table(
        index=[
            "District",
            "SchoolName",
            "Student LoginId",
            "school id"
        ],
        columns="Subject",
        values="Performance (%)"
    ).reset_index()

    # rename columns nicely
    projection.columns.name = None
    projection.rename(
        columns=lambda c: f"{c}_Performance (%)"
        if c not in ["District", "SchoolName", "Student LoginId", "school id"]
        else c,
        inplace=True
    )

    projection["Source File"] = file_path.name

    projection_records.append(projection)


# combine all grades
if projection_records:
    final_projection = pd.concat(projection_records, ignore_index=True)
else:
    final_projection = pd.DataFrame()


# write to new sheet called projection
with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    final_projection.to_excel(
        writer,
        sheet_name="projection",
        index=False
    )

print("Projection sheet created.")
