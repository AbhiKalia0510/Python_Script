# -*- coding: utf-8 -*-
"""
Created on Thu Sep 11 18:50:03 2025

@author: abhishek.a
"""
import pandas as pd

def split_csv_to_excel(csv_file, excel_file, rows_per_sheet=1000000):
    """
    Split a large CSV into multiple sheets in one Excel workbook.
    Each sheet will contain headers.
    """
    # Read CSV in chunks
    chunk_iter = pd.read_csv(csv_file, chunksize=rows_per_sheet)

    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        for i, chunk in enumerate(chunk_iter, start=1):
            sheet_name = f"Sheet{i}"
            print(f"Writing {len(chunk)} rows to {sheet_name}")
            chunk.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Saved split Excel file: {excel_file}")


# Usage
csv_file = "POS_Hardware_PVM_ClientAppVersion_Details.csv"   # your input CSV
excel_file = "POS_Hardware_PVM_ClientAppVersion_Details.xlsx"  # output workbook

split_csv_to_excel(csv_file, excel_file, rows_per_sheet=1000000)
