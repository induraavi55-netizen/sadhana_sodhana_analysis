import pandas as pd
from pathlib import Path

DATA_DIR = Path("data")   # change if needed

excel_files = [
    f for f in DATA_DIR.glob("*.xlsx")
    if f.name.lower().startswith("grade") and not f.name.startswith("~$")
]

for input_file in excel_files:

    print(f"Creating LO-wise pivots in {input_file.name}...")

    xls = pd.ExcelFile(input_file)

    with pd.ExcelWriter(
        input_file,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:

        for sheet in xls.sheet_names:

            # Only operate on long sheets
            if not sheet.endswith("_formatted_long"):
                continue

            df = pd.read_excel(input_file, sheet_name=sheet)

            if df.empty:
                continue

            # Base name: english_formatted_long -> english
            base_name = sheet.replace("_formatted_long", "")

            # ---- LO-wise aggregation ----
            lo_df = (
                df.groupby("LO")
                  .agg(
                      Questions=("Question", lambda x: ",".join(sorted(x.unique()))),
                      Avg_Perf=("Credit", "mean"),
                  )
                  .reset_index()
            )

            lo_df["Avg Performance (%)"] = (
                lo_df["Avg_Perf"] * 100
            ).round(0).astype(int)

            lo_df.drop(columns=["Avg_Perf"], inplace=True)

            # ---- Sort descending by performance ----
            lo_df = lo_df.sort_values(
    by=["Avg Performance (%)","LO"],
    ascending=[False,True]
)

            out_sheet = f"{base_name}_lo"

            lo_df.to_excel(writer, sheet_name=out_sheet, index=False)

            print(f"  → wrote {out_sheet}")

    print(f"Finished {input_file.name}")

print("All LO-wise sheets created.")
