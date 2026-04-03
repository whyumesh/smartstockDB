import os
import pandas as pd
from datetime import datetime
from collections import defaultdict

# ==============================
# CONFIGURATION
# ==============================

STOCKIST_LOG_PATH = "Stockist_Log.xlsx"
BASE_FILE_PATH = "Base_File.xlsx"
ATTACHMENTS_FOLDER = r"YOUR_FOLDER_PATH_HERE"   # <-- YOU WILL UPDATE

START_DATE = "2026-03-01 00:00:00"
END_DATE = "2026-03-31 23:59:59"

OUTPUT_FILE = "dry_run_output.xlsx"


# ==============================
# HELPER FUNCTIONS
# ==============================

def find_column(df, keywords):
    """Auto-detect column based on keywords"""
    for col in df.columns:
        for key in keywords:
            if key.lower() in col.lower():
                return col
    return None


def parse_datetime(dt_str):
    return pd.to_datetime(dt_str, errors='coerce')


# ==============================
# STEP 1: LOAD FILES
# ==============================

print("Loading files...")

stockist_df = pd.read_excel(STOCKIST_LOG_PATH)
base_df = pd.read_excel(BASE_FILE_PATH)

# ==============================
# STEP 2: AUTO DETECT COLUMNS
# ==============================

print("Detecting columns...")

# Stockist log columns
date_col = find_column(stockist_df, ["date", "time"])
email_col_log = find_column(stockist_df, ["email"])
stockist_code_col_log = find_column(stockist_df, ["code", "stockist"])

# Base file columns
stockist_code_col_base = find_column(base_df, ["code", "stockist"])
distributor_email_col = find_column(base_df, ["email"])
abm_email_col = find_column(base_df, ["abm"])
tbm_email_col = find_column(base_df, ["tbm"])

# Validation
required_cols = [
    date_col, stockist_code_col_log,
    stockist_code_col_base, distributor_email_col,
    abm_email_col, tbm_email_col
]

if any(col is None for col in required_cols):
    raise Exception("❌ Column detection failed. Please verify column names.")

# ==============================
# STEP 3: FILTER BY DATE RANGE
# ==============================

print("Filtering by date range...")

stockist_df[date_col] = pd.to_datetime(stockist_df[date_col], errors='coerce')

start_dt = parse_datetime(START_DATE)
end_dt = parse_datetime(END_DATE)

filtered_df = stockist_df[
    (stockist_df[date_col] >= start_dt) &
    (stockist_df[date_col] <= end_dt)
]

print(f"Filtered rows: {len(filtered_df)}")

# ==============================
# STEP 4: GET UNIQUE STOCKIST CODES
# ==============================

stockist_codes = filtered_df[stockist_code_col_log].dropna().unique()

print(f"Unique stockists found: {len(stockist_codes)}")

# ==============================
# STEP 5: CREATE MAPPING FROM BASE FILE
# ==============================

print("Mapping emails from base file...")

base_df[stockist_code_col_base] = base_df[stockist_code_col_base].astype(str)

mapping = {}

for _, row in base_df.iterrows():
    code = str(row[stockist_code_col_base]).strip()

    mapping[code] = {
        "distributor_email": str(row[distributor_email_col]).strip(),
        "abm_email": str(row[abm_email_col]).strip(),
        "tbm_email": str(row[tbm_email_col]).strip()
    }

# ==============================
# STEP 6: MATCH ATTACHMENTS
# ==============================

print("Scanning attachments folder...")

attachments_map = defaultdict(list)

for file in os.listdir(ATTACHMENTS_FOLDER):
    if "_" in file:
        stockist_code = file.split("_")[0].strip()
        attachments_map[stockist_code].append(file)

# ==============================
# STEP 7: PREPARE DRY RUN OUTPUT
# ==============================

print("Preparing dry run report...")

output_rows = []

for code in stockist_codes:
    code_str = str(code).strip()

    emails = mapping.get(code_str, {
        "distributor_email": "NOT FOUND",
        "abm_email": "NOT FOUND",
        "tbm_email": "NOT FOUND"
    })

    files = attachments_map.get(code_str, [])

    output_rows.append({
        "Stockist Code": code_str,
        "Distributor Email": emails["distributor_email"],
        "ABM Email": emails["abm_email"],
        "TBM Email": emails["tbm_email"],
        "Attachment Count": len(files),
        "Attachment Names": ", ".join(files) if files else "No Files Found"
    })

# ==============================
# STEP 8: SAVE OUTPUT
# ==============================

output_df = pd.DataFrame(output_rows)

output_df.to_excel(OUTPUT_FILE, index=False)

print(f"\n✅ Dry run completed. Output saved to: {OUTPUT_FILE}")
