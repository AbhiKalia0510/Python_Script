"""
POS Sales Report Processor
FINAL PRODUCTION VERSION – Auto Split Enabled
"""

import pandas as pd
import numpy as np
import os
import glob
import sys
from datetime import datetime

# ==========================================================
# CONFIGURATION
# ==========================================================

source_folder = r"C:\Users\abhishek.a\Downloads\POS Sale Report"
base_output_name = os.path.join(source_folder, "POS Sales Report_Final")
SPLIT_LIMIT = 1_000_000  # 10 lakh

pos_file_patterns = {
    "DCC": "Aggregated POS Sales Report (DCC)*",
    "EMI": "Aggregated POS Sales Report (EMI)*",
    "PART1": "Aggregated POS Sales Report Part-1*",
    "PART2": "Aggregated POS Sales Report Part-2*",
    "PART3": "Aggregated POS Sales Report Part-3*"
}

lookup_file_pattern = "PC Dump New For TRM Store name*"

print("\n🚀 Starting Processing...\n")

# ==========================================================
# SAFE FILE READER
# ==========================================================

def read_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path, encoding="utf-8-sig", low_memory=False)
    else:
        return pd.read_excel(path, engine="openpyxl")

# ==========================================================
# STEP 1 – LOAD POS FILES + PAYMENT REMARKS
# ==========================================================

pos_list = []

for key, pattern in pos_file_patterns.items():

    matches = glob.glob(os.path.join(source_folder, pattern))
    if not matches:
        print(f"⚠ File not found: {key}")
        continue

    file_path = matches[0]
    df = read_file(file_path)
    df.columns = df.columns.str.strip()

    print(f"Loaded {os.path.basename(file_path)} → {len(df):,} rows")

    required_cols = ["Payment Method", "Acquirer", "StoreID", "Store", "Txn Count"]
    for col in required_cols:
        if col not in df.columns:
            print(f"❌ Missing column '{col}' in {file_path}")
            sys.exit(1)

    df["Payment Method"] = df["Payment Method"].astype(str).str.strip()
    df["Acquirer"] = df["Acquirer"].astype(str).str.strip()

    if key == "DCC":
        df["Payment Remarks"] = "DCC"

    elif key == "EMI":
        df["Payment Remarks"] = "EMI"

    elif key in ["PART1", "PART2","PART3"]:
        df["Payment Remarks"] = np.where(
            df["Payment Method"].str.lower() == "paper pos",
            "Paper POS",
            np.where(
                df["Acquirer"]
                    .astype(str)
                    .str.replace(r'["\s]+', '', regex=True)
                    .str.upper()
                    .eq("PHONEPE"),
                "PhonePe",
                df["Payment Method"]
            )
        )

    pos_list.append(df)

if not pos_list:
    print("❌ No POS files found.")
    sys.exit(1)

pos_df = pd.concat(pos_list, ignore_index=True)

print(f"\n✅ Combined Rows: {len(pos_df):,}")

pos_df["StoreID"] = pos_df["StoreID"].astype(str).str.strip()
pos_df["Store"] = pos_df["Store"].astype(str).str.strip()

# ==========================================================
# STEP 2 – LOAD LOOKUP FILE
# ==========================================================

lookup_matches = glob.glob(os.path.join(source_folder, lookup_file_pattern))
if not lookup_matches:
    print("❌ Lookup file not found.")
    sys.exit(1)

lookup_df = read_file(lookup_matches[0])
lookup_df.columns = lookup_df.columns.str.strip()

storeid_lookup = "Store Id"
store_lookup = "STORE"
gi_col = "GI Status"
bpcl_col = "BPCL Remark"

lookup_df[storeid_lookup] = lookup_df[storeid_lookup].astype(str).str.strip()
lookup_df[store_lookup] = lookup_df[store_lookup].astype(str).str.strip()

print(f"Loaded Lookup File → {len(lookup_df):,} rows")

# ==========================================================
# STEP 3 – LOOKUP (StoreID → fallback Store)
# ==========================================================

merge_id = pos_df.merge(
    lookup_df[[storeid_lookup, gi_col, bpcl_col]],
    how="left",
    left_on="StoreID",
    right_on=storeid_lookup
)

merge_store = pos_df.merge(
    lookup_df[[store_lookup, gi_col, bpcl_col]],
    how="left",
    left_on="Store",
    right_on=store_lookup
)

merge_id[gi_col] = merge_id[gi_col].combine_first(merge_store[gi_col])
merge_id[bpcl_col] = merge_id[bpcl_col].combine_first(merge_store[bpcl_col])

merged_df = merge_id

merged_df[gi_col] = merged_df[gi_col].fillna("")
merged_df[bpcl_col] = merged_df[bpcl_col].fillna("")

print("✅ Lookup Completed")

# ==========================================================
# STEP 4 – FINAL TXN & POINTER REMARKS
# ==========================================================

today = pd.Timestamp.today()
days_elapsed = (today - pd.offsets.Day(1)).day
days_in_month = today.days_in_month

txn = pd.to_numeric(merged_df["Txn Count"], errors="coerce").fillna(0)

pm = merged_df["Payment Method"].astype(str).str.strip()
gi = merged_df[gi_col].astype(str).str.strip()

bpcl_clean = merged_df[bpcl_col].fillna("").astype(str).str.strip()
bpcl_lower = bpcl_clean.str.lower()

cond_gi = gi.eq("GI Enabled")

cond_paper_50 = (
    pm.str.lower().eq("paper pos") &
    ((bpcl_clean == "") | (bpcl_lower == "not applicable"))
)

cond_paper_75 = (
    pm.str.lower().eq("paper pos") &
    (bpcl_clean != "") &
    (bpcl_lower != "not applicable")
)

merged_df["Final Txn"] = np.select(
    [cond_gi, cond_paper_50, cond_paper_75],
    [txn / 7.3, txn * 0.50, txn * 1],
    default=txn
)

merged_df["Final Txn"] = (
    merged_df["Final Txn"] / days_elapsed
) * days_in_month

merged_df["Final Txn"] = merged_df["Final Txn"].round(2)

payment_remarks = (
    merged_df["Payment Remarks"]
    .fillna("")
    .astype(str)
    .str.strip()
)

payment_lower = payment_remarks.str.lower()

# Group for UPI mapping
upi_group = [
    "paper pos", "card", "wallet", "reward", "cod",
    "ncmc", "pay_by_link", "cardless", "pgatpos",
    "sms_pay", "nbfc"
]

merged_df["Pointer Remarks"] = np.select(
    [
        cond_gi,

        cond_paper_50,
        cond_paper_75,

        payment_lower.isin(["emi", "dcc"]),

        payment_lower.eq("phonepe"),

        payment_lower.eq("card"),

        payment_lower.isin(upi_group)
    ],
    [
        "GI Enabled 550 Transaction",
        "Paper POS 50%",
        "For BPCL/IGL/ 100% Paper POS Transaction Captured",
        "EMI/DCC",
        "PhonePe",
        "CARD",
        "UPI"
    ],
    default=payment_remarks
)


print("✅ Final Txn & Pointer Remarks Calculated")

# ==========================================================
# STEP 5 – AUTO SPLIT EXPORT
# ==========================================================

total_rows = len(merged_df)

if total_rows <= SPLIT_LIMIT:
    output_path = base_output_name + ".csv"
    merged_df.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"\n📁 File saved: {output_path}")

else:
    print("\n⚡ Large dataset detected → Splitting file...")

    parts = (total_rows // SPLIT_LIMIT) + 1

    for i in range(parts):
        start = i * SPLIT_LIMIT
        end = start + SPLIT_LIMIT
        chunk = merged_df.iloc[start:end]

        output_path = f"{base_output_name}_Part_{i+1}.csv"
        chunk.to_csv(output_path, index=False, encoding="utf-8-sig")

        print(f"📁 Created: {output_path}")

print("\n🎯 Process Completed Successfully\n")