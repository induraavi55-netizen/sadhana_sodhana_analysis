import pandas as pd
from pathlib import Path

DATA_DIR = Path("data")

# --------------------------------------------------
# STEP 1: COLLECT *_formatted SHEETS
# --------------------------------------------------

subject_buckets = {}

for file_path in DATA_DIR.glob("Grade*.xlsx"):

    # skip temp/lock files
    if file_path.name.startswith("~$"):
        continue

    file_stem = file_path.stem

    try:
        xls = pd.ExcelFile(file_path, engine="openpyxl")
    except Exception as e:
        print(f"Skipping {file_path.name}: {e}")
        continue

    for sheet in xls.sheet_names:

        sheet_clean = sheet.strip().lower()

        if sheet_clean.endswith("_formatted"):

            base_subject = sheet_clean.replace("_formatted", "")

            try:
                df = pd.read_excel(file_path, sheet_name=sheet)
            except Exception as e:
                print(f"Error reading {sheet} in {file_path.name}: {e}")
                continue

            subject_buckets.setdefault(base_subject, []).append(
                (file_stem, df)
            )

            print(f"Collected {sheet} from {file_path.name}")

# --------------------------------------------------
# STEP 2: WRITE SUBJECT FILES + ALL_GRADES
# --------------------------------------------------

OUTPUT_DIR = DATA_DIR / "clustered"
OUTPUT_DIR.mkdir(exist_ok=True)

for subject, sheets in subject_buckets.items():

    if not sheets:
        continue

    out_file = OUTPUT_DIR / f"{subject}.xlsx"

    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:

        combined_frames = []

        for sheet_name, df in sheets:

            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

            temp = df.copy()
            temp["Source Grade File"] = sheet_name
            combined_frames.append(temp)

        combined_df = pd.concat(combined_frames, ignore_index=True)

        combined_df.to_excel(writer, sheet_name="ALL_GRADES", index=False)

    print(f"Created {out_file.name}")

# --------------------------------------------------
# STEP 3: BUILD all_subjects.xlsx FROM ALL_GRADES
# --------------------------------------------------

MASTER_FILE = OUTPUT_DIR / "all_subjects.xlsx"

with pd.ExcelWriter(MASTER_FILE, engine="openpyxl") as writer:

    for subject_file in OUTPUT_DIR.glob("*.xlsx"):

        if subject_file.name.lower() == "all_subjects.xlsx":
            continue

        try:
            xls = pd.ExcelFile(subject_file, engine="openpyxl")
        except Exception as e:
            print(f"Skipping {subject_file.name}: {e}")
            continue

        if "ALL_GRADES" not in xls.sheet_names:
            print(f"Skipping {subject_file.name} (no ALL_GRADES)")
            continue

        df = pd.read_excel(subject_file, sheet_name="ALL_GRADES")

        subject_name = subject_file.stem.lower()
        new_sheet = f"all_grades_{subject_name}"

        df.to_excel(writer, sheet_name=new_sheet[:31], index=False)

        print(f"Added {new_sheet}")

print("\n✅ Built all_subjects.xlsx")

# --------------------------------------------------
# STEP 4: BUILD PERFORMANCE SUMMARY
# --------------------------------------------------

summary_rows = []

try:
    xls = pd.ExcelFile(MASTER_FILE, engine="openpyxl")
except Exception as e:
    print(f"Error opening master file: {e}")
    exit()

for sheet in xls.sheet_names:

    if not sheet.lower().startswith("all_grades_"):
        continue

    subject = sheet.replace("all_grades_", "")

    df = pd.read_excel(MASTER_FILE, sheet_name=sheet)

    if "Performance (%)" not in df.columns:
        print(f"Skipping {sheet} (no Performance column)")
        continue

    perf = pd.to_numeric(df["Performance (%)"], errors="coerce")

    avg_perf = perf.mean()

    summary_rows.append({
        "Subject": subject,
        "Avg Performance": round(avg_perf, 0)
    })

summary_df = pd.DataFrame(summary_rows)

with pd.ExcelWriter(
    MASTER_FILE,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:

    summary_df.to_excel(writer, sheet_name="Perf_summary", index=False)

print("\n📊 Added Perf_summary")