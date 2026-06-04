# -*- coding: utf-8 -*-
import os
import re
import gc
import pandas as pd

folder_path = r"D:\Store -Summery Folder\Pending For Dispatch _Petro & Plutus\Pending on Stock\Pending\New folder"
output_csv = os.path.join(folder_path, "Combined_Report.csv")

def extract_date_from_name(name):
    m = re.search(r'(\d{1,2})\s([A-Za-z]{3})\s(\d{4})', name)
    return f"{m.group(1).zfill(2)} {m.group(2)} {m.group(3)}" if m else None

def clean_df(df):
    df = df.dropna(axis=0, how='all')
    df = df.dropna(axis=1, how='all')
    df.columns = [str(c).strip() for c in df.columns]

    dup = df.columns.duplicated()
    if dup.any():
        counts, new_cols = {}, []
        for c in df.columns:
            counts[c] = counts.get(c, 0) + 1
            new_cols.append(c if counts[c] == 1 else f"{c}_{counts[c]}")
        df.columns = new_cols

    return df

def try_read_excel(path, report_sheet):
    try:
        return pd.read_excel(path, sheet_name=report_sheet, engine="openpyxl", dtype=str)
    except Exception:
        tmp = pd.read_excel(path, sheet_name=report_sheet, engine="openpyxl", header=None, dtype=str)

        head_span = min(20, len(tmp))
        nn = tmp.iloc[:head_span].notna().sum(axis=1)
        header_idx = int(nn.idxmax())

        headers = tmp.iloc[header_idx].fillna('').astype(str).tolist()
        df = tmp[(header_idx+1):].copy()
        df.columns = headers

        return df

canonical_cols = None
first_write = True

for file in os.listdir(folder_path):

    file_path = os.path.join(folder_path, file)

    if file.startswith("~$"):
        continue

    try:

        df = None

        if file.lower().endswith(".xlsx"):

            xls = pd.ExcelFile(file_path, engine="openpyxl")

            sheet_name = next(
                (s for s in xls.sheet_names if str(s).lower().startswith("report")),
                None
            )

            if not sheet_name:
                print(f"Skipping {file}: no sheet starting with 'report'")
                continue

            print(f"Reading {file} -> {sheet_name}")
            df = try_read_excel(file_path, sheet_name)

        elif file.lower().endswith(".csv"):

            print(f"Reading CSV file {file}")

            df = pd.read_csv(
                file_path,
                dtype=str,
                low_memory=False,
                on_bad_lines='skip',
                encoding_errors='ignore'
            )

        else:
            continue

        df = clean_df(df)

        if df.shape[1] > 300:
            print(f"Warning: {file} has {df.shape[1]} columns. Skipping.")
            continue

        df["Source_File"] = file
        df["Report_Date"] = extract_date_from_name(file)

        if canonical_cols is None:

            canonical_cols = df.columns.tolist()

        else:

            missing = [c for c in canonical_cols if c not in df.columns]
            extra = [c for c in df.columns if c not in canonical_cols]

            df = df.drop(columns=extra, errors='ignore')

            for c in missing:
                df[c] = pd.NA

            df = df[canonical_cols]

        df.to_csv(
            output_csv,
            mode="w" if first_write else "a",
            header=first_write,
            index=False
        )

        first_write = False

        del df
        gc.collect()

    except Exception as e:
        print(f"Error reading {file}: {e}")

print(f"\nAll files streamed into: {output_csv}")