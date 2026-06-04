
"""""@author: abhishek.a with the help to AI
"""""

import pandas as pd
import re
from pathlib import Path
import logging
import sys

# =========================
# LOGGING
# =========================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    stream=sys.stdout,
)
logger = logging.getLogger(__name__)

# =========================
# CONFIG
# =========================
CHUNK_SIZE = 200000  # Safe chunk size (adjust if needed)

# =========================
# CLEAN FUNCTION
# =========================
ILLEGAL_CHARACTERS_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')

def clean_dataframe(df):
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.replace(ILLEGAL_CHARACTERS_RE, '', regex=True)
    return df

# =========================
# SAFE COLUMN HANDLER
# =========================
def ensure_columns(df, columns):
    for col in columns:
        if col not in df.columns:
            df[col] = ''
    return df

# =========================
# LOAD INVENTORY
# =========================
def load_inventory():
    inv_path = Path(r'D:\PC Dump\Input\Stock Inventory Report-Current Month.xlsb')
    logger.info("Loading Inventory...")

    sf = pd.read_excel(inv_path, engine='pyxlsb')

    inv1 = sf[['Right', 'No.']].copy()
    inv1.columns = ['8 Digit', 'Available Serial No.']
    inv1 = inv1.astype(str).drop_duplicates('8 Digit')

    inv2 = sf[['Right.1', 'No..1']].copy()
    inv2.columns = ['8 Digit', 'Available Serial']
    inv2 = inv2.astype(str).drop_duplicates('8 Digit')

    inv3 = sf[['No.', 'Model Type']].copy()
    inv3.columns = ['Full Serial No.', 'Model Name']
    inv3 = inv3.astype(str).drop_duplicates('Full Serial No.')

    inv4 = sf[['No..1', 'Model Type.1']].copy()
    inv4.columns = ['Full Serial No.', 'ModelName']
    inv4 = inv4.astype(str).drop_duplicates('Full Serial No.')

    return inv1, inv2, inv3, inv4

# =========================
# TRANSFORM LOGIC
# =========================
def transform_chunk(df, inv1, inv2, inv3, inv4):

    df = clean_dataframe(df)

    df = ensure_columns(df, [
        'FULL_SERIAL_NUMBER', 'HARDWARE_ID', 'IMEI_NUMBER',
        'IS_BLOCKED', 'IS_DELETED'
    ])

    # Serial logic
    df['Full Serial No.'] = df['FULL_SERIAL_NUMBER'].astype(str).str.replace("'", "")
    df['8 Digit'] = df['Full Serial No.'].str[-8:]

    df['HARDWARE_ID'] = df['HARDWARE_ID'].astype(str)

    # Match
    df['Match Hardware No.'] = (df['8 Digit'] == df['HARDWARE_ID']).map({True: 'True', False: 'False'})

    # Merge inventory
    df = df.merge(inv1, how='left', on='8 Digit')
    df['Available Serial No.'] = df['Available Serial No.'].fillna('NA')

    df = df.merge(inv2, how='left', on='8 Digit')
    df['Available Serial'] = df['Available Serial'].fillna('NA')

    # Serial fallback
    df['Available Serial No.'] = df['Available Serial No.'].mask(
        df['Available Serial No.'] == 'NA',
        df['Available Serial']
    )

    # Model merge
    df = df.merge(inv3, how='left', on='Full Serial No.')
    df['Model Name'] = df['Model Name'].fillna('NA')

    df = df.merge(inv4, how='left', on='Full Serial No.')
    df['ModelName'] = df['ModelName'].fillna('NA')

    df['Model Name'] = df['Model Name'].mask(
        df['Model Name'].str.upper().isin(['NA', '', 'NONE']),
        df['ModelName']
    )

    # Final status
    df['IS_BLOCKED'] = pd.to_numeric(df['IS_BLOCKED'], errors='coerce').fillna(0)
    df['IS_DELETED'] = pd.to_numeric(df['IS_DELETED'], errors='coerce').fillna(0)

    df['Final Status'] = 'In-Active'

    df.loc[(df['IS_BLOCKED'] == 0) & (df['IS_DELETED'] == 0), 'Final Status'] = 'Active'
    df.loc[(df['IS_BLOCKED'] == 1) & (df['IS_DELETED'] == 0), 'Final Status'] = 'Active_Blocked'

    return df

# =========================
# PROCESS FILE (CHUNK MODE)
# =========================
def process_file(file_path, sheet_name, writer, inv1, inv2, inv3, inv4):

    logger.info(f"Processing {file_path.name}")

    row_start = 0

    for i, chunk in enumerate(pd.read_csv(
        file_path,
        chunksize=CHUNK_SIZE,
        encoding='utf-8',
        encoding_errors='ignore',
        low_memory=False
    )):

        logger.info(f"Chunk {i+1} - rows: {len(chunk)}")

        processed = transform_chunk(chunk, inv1, inv2, inv3, inv4)

        processed.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=row_start,
            index=False,
            header=(row_start == 0)
        )

        row_start += len(processed)

# =========================
# MAIN
# =========================
def main():

    base = Path(r"C:\Users\abhishek.a\Downloads\1st May'26")

    files = [
        "POS_Hardware_PVM_ClientAppVersion_Details.csv",
        "POS_Hardware_PVM_ClientAppVersion_Details (1).csv",
        "POS_Hardware_PVM_ClientAppVersion_Details (2).csv",
        "POS_Hardware_PVM_ClientAppVersion_Details (3).csv",
        "POS_Hardware_PVM_ClientAppVersion_Details (4).csv",
    ]

    sheets = [
        "PC-Dump-1",
        "PC-Dump-2",
        "PC-Dump-3",
        "PC-Dump-4",
        "PC-Dump-5",
    ]

    output = base / "Enterprise_PC_Dump_Output.xlsx"

    inv1, inv2, inv3, inv4 = load_inventory()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        for file, sheet in zip(files, sheets):

            path = base / file

            if not path.exists():
                logger.warning(f"Missing: {file}")
                continue

            process_file(path, sheet, writer, inv1, inv2, inv3, inv4)

    logger.info("✅ ENTERPRISE JOB COMPLETED")

# =========================
# RUN
# =========================
if __name__ == "__main__":
    main()