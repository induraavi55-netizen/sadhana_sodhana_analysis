import pandas as pd
from pathlib import Path

DATA_DIR = Path("data")

# ✅ Skip temp/lock files (~$)
excel_files = [
    f for f in DATA_DIR.glob("*.xlsx")
    if f.name.lower().startswith("grade") and not f.name.startswith("~$")
]

for input_file in excel_files:
    try:
        print(f"Creating LO-question pivots in {input_file.name}...")

        # ✅ Explicit engine so pandas doesn’t have identity crisis
        xls = pd.ExcelFile(input_file, engine="openpyxl")

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

                # ---- Base name ----
                base_name = sheet.replace("_formatted_long", "")

                # ---- LO + Question aggregation ----
                lo_q_df = (
                    df.groupby(["LO", "Question"])
                      .agg(Avg_Perf=("Credit", "mean"))
                      .reset_index()
                )

                lo_q_df["Performance (%)"] = (
                    lo_q_df["Avg_Perf"] * 100
                ).round(0).astype(int)

                lo_q_df.drop(columns=["Avg_Perf"], inplace=True)

                # ---- Sort ----
                lo_q_df = lo_q_df.sort_values(
                    by=["LO", "Performance (%)"],
                    ascending=[True, False]
                )

                out_sheet = f"{base_name}_lo_question"

                # ⚠️ Excel sheet name max = 31 chars
                out_sheet = out_sheet[:31]

                lo_q_df.to_excel(writer, sheet_name=out_sheet, index=False)

                print(f"  → wrote {out_sheet}")

        print(f"Finished {input_file.name}")

    except Exception as e:
        print(f"❌ Error in {input_file.name}: {e}")

print("All LO-question sheets created.")