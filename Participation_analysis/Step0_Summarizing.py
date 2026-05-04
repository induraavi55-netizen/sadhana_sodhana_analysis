import pandas as pd
import numpy as np
from pathlib import Path

# -----------------------------
# CONFIG
# -----------------------------

EXAM_GRADES = [4]

PARTICIPATING_SCHOOLS = [
    "APS GOLCONDA",
    "APS R K PURAM Secunderabad"

]

DATA_DIR = Path("data")
FILE = DATA_DIR / "REG VS PART.xlsx"
SHEET = "Assessment Participation"

# -----------------------------
# LOAD RAW (TWO HEADER ROWS)
# -----------------------------

raw = pd.read_excel(FILE, sheet_name=SHEET, header=None)

print("\n===== DEBUG: RAW SHAPE =====")
print(raw.shape)

# real data starts from row index 2
df = raw.iloc[2:].copy()

# second row contains grade numbers
header_row = raw.iloc[1]

cols = list(header_row)

# fix first columns
cols[0] = "S.No"
cols[1] = "School Name"
cols[2] = "District"

# totals
cols[15] = "Total Registered"
cols[28] = "Total Participated"

# last two
cols[-2] = "Contact Name"
cols[-1] = "Contact Phone"

df.columns = cols

print("\n===== DEBUG: ASSIGNED COLUMNS =====")
for i, c in enumerate(df.columns):
    print(i, c)

# drop total row
df = df[df["School Name"].astype(str).str.lower() != "total"]

print("\n===== DEBUG: DF SHAPE AFTER DROP TOTAL =====")
print(df.shape)

print("\n===== DEBUG: SCHOOL NAMES =====")
print(df["School Name"].unique())


# -----------------------------
# IDENTIFY GRADE COLUMN INDEXES
# -----------------------------

# registered block positions
reg_idx = list(range(3, 15))

# participated block positions
part_idx = list(range(16, 28))

# map index -> grade number safely
reg_grade_map = {}
for i in reg_idx:
    col = str(df.columns[i]).strip()
    try:
        reg_grade_map[i] = int(float(col))
    except ValueError:
        continue

part_grade_map = {}
for i in part_idx:
    col = str(df.columns[i]).strip()
    try:
        part_grade_map[i] = int(float(col))
    except ValueError:
        continue

# keep only exam grades
reg_idx = [i for i in reg_idx if reg_grade_map.get(i) in EXAM_GRADES]
part_idx = [i for i in part_idx if part_grade_map.get(i) in EXAM_GRADES]

print("\n===== DEBUG: REG IDX =====", reg_idx)
print("===== DEBUG: PART IDX =====", part_idx)

# -----------------------------
# FILTER SCHOOLS FIRST
# -----------------------------

df["School Name"] = df["School Name"].astype(str).str.strip()

school_filter = [s.lower() for s in PARTICIPATING_SCHOOLS]

df_filt = df[
    df["School Name"]
        .str.lower()
        .isin(school_filter)
]

# -----------------------------
# SCHOOL TOTALS (SAFE)
# -----------------------------

school_totals = df_filt[["School Name"]].copy()

school_totals["Registered"] = df_filt.iloc[:, reg_idx].sum(axis=1)
school_totals["Participated"] = df_filt.iloc[:, part_idx].sum(axis=1)

school_totals["Not Participated"] = (
    school_totals["Registered"] - school_totals["Participated"]
)

school_totals = school_totals[
    ["School Name", "Participated", "Not Participated", "Registered"]
]

print("\n===== DEBUG: SCHOOL TOTALS =====")
print(school_totals)

# -----------------------------
# OVERALL GRADE TOTALS
# -----------------------------

overall = pd.DataFrame({
    "Grade": [reg_grade_map[i] for i in reg_idx],
    "Registered": df_filt.iloc[:, reg_idx].sum().values,
    "Participated": df_filt.iloc[:, part_idx].sum().values,
})

overall["Registered"] = pd.to_numeric(overall["Registered"], errors="coerce").fillna(0)
overall["Participated"] = pd.to_numeric(overall["Participated"], errors="coerce").fillna(0)

# ADDED: Not Participated column
overall["Not Participated"] = overall["Registered"] - overall["Participated"]

overall["Participation %"] = np.where(
    overall["Registered"] > 0,
    (overall["Participated"] / overall["Registered"] * 100),
    0
)

overall["Participation %"] = (
    overall["Participation %"]
        .replace([np.inf, -np.inf], 0)
        .fillna(0)
        .round(0)
        .astype(int)
)

print("\n===== DEBUG: OVERALL =====")
print(overall)


# -----------------------------
# WRITE BACK TO SAME FILE
# -----------------------------

with pd.ExcelWriter(FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    school_totals.to_excel(writer, sheet_name="schl_wise", index=False)
    overall.to_excel(writer, sheet_name="grade_wise", index=False)

print("\n✅ DONE. Sheets written:")
print(" - schl_wise")
print(" - grade_wise")