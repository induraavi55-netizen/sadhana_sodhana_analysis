import pandas as pd
from datetime import time

# ==============================
# 1. Define Sheets You Want
# ==============================
input_file = "Grade_7.xlsx"
sheets_to_read = ["English", "Mathematics", "Science"]
output_file = "Rule_Based_Difficulty_Comparison_allsub.xlsx"

# ==============================
# 2. Manager Presence Rules
# ==============================
def get_manager_status(row):
    school = row['SchoolName']
    date = row['Date']
    t = row['Time']

    if date == pd.to_datetime("2026-02-09").date():
        if "Vikarabad" in school and t <= time(12, 30):
            return "Manager Present"
        if "Tandur" in school and time(13, 0) <= t <= time(17, 0):
            return "Manager Present"

    if date == pd.to_datetime("2026-02-10").date():
        if "Digwal" in school:
            return "Manager Present"

    if date == pd.to_datetime("2026-02-12").date():
        if "Siddipet" in school and t <= time(12, 30):
            return "Manager Present"

    return "Manager Absent"

# ==============================
# 3. Write Output
# ==============================
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

    for sheet in sheets_to_read:

        df = pd.read_excel(input_file, sheet_name=sheet)
        df.columns = df.columns.str.strip()

        # Convert Date
        df['Assessment Date'] = pd.to_datetime(df['Assessment Date'], errors='coerce')
        df = df.dropna(subset=['Assessment Date'])

        df['Date'] = df['Assessment Date'].dt.date
        df['Time'] = df['Assessment Date'].dt.time

        # Manager status
        df["Manager_Status"] = df.apply(get_manager_status, axis=1)

        # Keep only school-date groups having both statuses
        group_counts = df.groupby(['SchoolName', 'Date'])['Manager_Status'].nunique().reset_index()
        valid_groups = group_counts[group_counts['Manager_Status'] > 1][['SchoolName', 'Date']]
        df = df.merge(valid_groups, on=['SchoolName', 'Date'])

        # ==============================
        # Convert Wide to Long (Difficulty-wise)
        # ==============================
        records = []

        for _, row in df.iterrows():
            for i in range(1, 16):
                difficulty = row.get(f"Q{i}_difficulty_level")
                credit = row.get(f"Q{i}_credit")

                if pd.notna(difficulty):
                    records.append({
                        "School": row["SchoolName"],
                        "Date": row["Date"],
                        "Difficulty_Level": difficulty,
                        "Credit": pd.to_numeric(credit, errors="coerce"),
                        "Manager_Status": row["Manager_Status"]
                    })

        long_df = pd.DataFrame(records)

        # ==============================
        # Difficulty-wise performance
        # ==============================
        summary = (
            long_df
            .groupby(["School", "Date", "Difficulty_Level", "Manager_Status"])
            .agg(Total=("Credit", "count"),
                 Correct=("Credit", "sum"))
            .reset_index()
        )

        summary["Performance_%"] = round(
            (summary["Correct"] / summary["Total"]) * 100, 2
        )

        pivot = summary.pivot_table(
            index=["School", "Date", "Difficulty_Level"],
            columns="Manager_Status",
            values="Performance_%"
        ).reset_index()

        pivot = pivot.rename(columns={
            "Manager Present": "Manager Present (%)",
            "Manager Absent": "Manager Absent (%)"
        })

        # Write to Excel
        pivot.to_excel(writer, sheet_name=f"{sheet}_Difficulty", index=False)

print("Difficulty-wise comparison written per sheet successfully.")