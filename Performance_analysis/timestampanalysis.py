import pandas as pd
from datetime import time

# ===== 1. File Names =====
input_file = "Grade_7.xlsx"
output_file = "time stamp perf analysis.xlsx"

# ===== 2. Specify Sheet Names =====
sheets_to_read = ["English", "Mathematics", "Science"]   # change as needed

all_data = []

# ===== 3. Read Selected Sheets =====
for sheet in sheets_to_read:
    df = pd.read_excel(input_file, sheet_name=sheet)
    
    df.columns = df.columns.str.strip()
    
    # If Subject column missing, assign sheet name
    if 'Subject' not in df.columns:
        df['Subject'] = sheet
    
    all_data.append(df)

# Combine all sheets
df = pd.concat(all_data, ignore_index=True)

print("Sheets loaded:", sheets_to_read)

# ===== 4. Convert Assessment Date =====
df['Assessment Date'] = pd.to_datetime(df['Assessment Date'], errors='coerce')
df = df.dropna(subset=['Assessment Date'])

df['Date'] = df['Assessment Date'].dt.date
df['Time'] = df['Assessment Date'].dt.time

# ===== 5. Clean perf % (or calculate if missing) =====
df.columns = df.columns.str.strip()

if 'perf %' in df.columns:
    if df['perf %'].dtype == object:
        df['perf %'] = df['perf %'].str.replace('%', '', regex=False)
    df['perf %'] = pd.to_numeric(df['perf %'], errors='coerce')
else:
    # Auto-calculate from credit columns
    credit_cols = [col for col in df.columns if col.endswith('_credit')]
    df[credit_cols] = df[credit_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
    df['total_correct'] = df[credit_cols].sum(axis=1)
    total_questions = len(credit_cols)
    df['perf %'] = (df['total_correct'] / total_questions) * 100

# ===== 6. Assign Time Ranges =====
def assign_time_range(t):
    if time(9, 0) <= t <= time(12, 30):
        return "9:00–12:30"
    elif time(12, 31) <= t <= time(17, 0):
        return "12:31–5:00"
    else:
        return None

df['Time Range'] = df['Time'].apply(assign_time_range)
df = df.dropna(subset=['Time Range'])

# ===== 7. Create Pivot =====
pivot = (
    df.pivot_table(
        index=['SchoolName', 'Date', 'Subject'],
        columns='Time Range',
        values='perf %',
        aggfunc='mean'
    )
    .reset_index()
)

pivot = pivot.rename(columns={'SchoolName': 'School'})

# ===== 8. Save Output =====
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    pivot.to_excel(writer, index=False, sheet_name="Summary")

print("File created successfully:", output_file)