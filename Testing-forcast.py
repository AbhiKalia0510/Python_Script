#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Optimized Inventory Forecast Pipeline
Original: Abhishek.a with the of AI | Optimized for speed — no logic or sheets removed.

Key optimizations:
  1. stock_inventory.xlsx read once, shared across working/non_working/l1_l2_repair
  2. Forecast Model HUB.xlsb read once (all sheets) instead of 3x
  3. Forecast.xlsx written in one pass per stage instead of open/close per function
  4. VLOOKUP lambda replaced with vectorized map via pre-built dict
  5. unmatch_code uses vectorized isin() instead of row-wise apply
  6. Doable mapping HUB.xlsb read once instead of twice
  7. Location Mapping lookups batched (single set_index, multiple .map calls)
"""

import os
import datetime
import numpy as np
import pandas as pd

print("job started at:", datetime.datetime.now().time())

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _read_csv_safe(path: str) -> pd.DataFrame:
    """Read CSV with UTF-8 fallback to ISO-8859-1."""
    try:
        return pd.read_csv(path, encoding="utf-8")
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="ISO-8859-1")


def _write_excel_sheets(path: str, sheets: dict, mode: str = "w", engine: str = "openpyxl"):
    """Write multiple sheets to an Excel file in one ExcelWriter session."""
    kwargs = {"engine": engine, "mode": mode}
    if mode == "a":
        kwargs["if_sheet_exists"] = "replace"
    with pd.ExcelWriter(path, **kwargs) as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)


# ---------------------------------------------------------------------------
# 1. stock_inventory  (returns cleaned DataFrame — avoids re-read downstream)
# ---------------------------------------------------------------------------

def stock_inventory(script_dir: str) -> pd.DataFrame:
    output_dir = os.path.join(script_dir, "Output_File")
    os.makedirs(output_dir, exist_ok=True)

    # --- load CSV → Excel (kept for compatibility) ---
    csv_path = os.path.join(script_dir, "Inventory.csv")
    excel_path = os.path.join(script_dir, "stock_inventory.xlsx")

    csv_data = _read_csv_safe(csv_path)
    csv_data.to_excel(excel_path, index=False)

    data = pd.read_excel(excel_path)

    emp_col   = "Location: Emp Code"
    loc_col   = "Location: Location Name"
    ltype_col = "Location: Location Type"
    lreg_col  = "Location: Region"

    # --- strip emp code ---
    if emp_col in data.columns:
        data[emp_col] = data[emp_col].astype(str).str.strip()

    # --- hub flag for Stock/Warehouse/Pranjal ---
    if loc_col in data.columns:
        mask_hub = data[loc_col].str.contains("Stock|Warehouse|Pranjal", case=False, na=False)
        data.loc[mask_hub, ltype_col] = "Hub"

        # Pranjal → Stock Guwahati
        data[loc_col] = data[loc_col].str.replace(r".*Pranjal.*", "Stock Guwahati", regex=True)
        guw_rows = data[loc_col] == "Stock Guwahati"
        if guw_rows.any() and emp_col in data.columns:
            pranjal_code = data.loc[guw_rows, emp_col].iloc[0]
            data.loc[guw_rows, emp_col] = pranjal_code

    # --- Regional/State Warehouse → Hub; TSE Van → FSR ---
    if ltype_col in data.columns:
        data.loc[data[ltype_col].str.contains("Regional Warehouse|State Warehouse", case=False, na=False), ltype_col] = "Hub"
        data.loc[data[ltype_col].str.contains("TSE Van", case=False, na=False), ltype_col] = "FSR"

    # --- Noida HO / Sharwan Kumar fixes ---
    if loc_col in data.columns and ltype_col in data.columns and emp_col in data.columns:
        noida_mask = data[loc_col].str.contains("Noida HO", case=False, na=False) & (data[emp_col] == "99999")
        data.loc[noida_mask, emp_col] = "10033"

        specific = ["10033-Sharwan Kumar - HO", "Noida HO"]
        ho_mask = data[loc_col].isin(specific) & (data[ltype_col] == "HO")
        data.loc[ho_mask, [loc_col, ltype_col]] = ["10033-Sharwan Kumar - HO", "Hub"]

    # --- Gujarat emp code fix ---
    if loc_col in data.columns and emp_col in data.columns:
        guj_mask = data[loc_col].str.contains("Gujarat", case=False, na=False) & (data[emp_col] == "10796")
        print("Rows matching Gujarat and emp code 10796:")
        print(data[guj_mask])
        data.loc[guj_mask, emp_col] = "18156"

    # --- Kerala fixes ---
    if loc_col in data.columns and ltype_col in data.columns and emp_col in data.columns:
        kerala_mask = data[loc_col].str.contains("State Warehouse - Kerala", case=False, na=False)
        data.loc[kerala_mask & (data[emp_col] == "11395"), emp_col] = "3718"
        data.loc[kerala_mask, loc_col] = "Stock Kerala"

    # --- Repair → In House Repair Centre ---
    if loc_col in data.columns and ltype_col in data.columns:
        data.loc[data[loc_col].str.contains("Repair", case=False, na=False), ltype_col] = "In House Repair Centre"

    # --- Delhi NCR → North ---
    if lreg_col in data.columns:
        data.loc[data[lreg_col].str.contains("Delhi NCR", case=False, na=False), lreg_col] = "North"

    data.to_excel(excel_path, index=False)
    print(f"stock_inventory saved: {excel_path}")

    # --- fsr_hub: in-transit working ---
    required = ["In Transit?", "Quantity On Hand", "Product Working Condition Status"]
    if all(c in data.columns for c in required):
        filtered = data[
            (data["In Transit?"] == 1) &
            (data["Quantity On Hand"] == 1) &
            (data["Product Working Condition Status"] == "Working")
        ]
        fsr_hub_path = os.path.join(script_dir, "Output_File", "fsr_hub.xlsx")
        filtered.to_excel(fsr_hub_path, index=False)
        print(f"fsr_hub saved: {fsr_hub_path}")

    return data   # <-- returned so downstream functions skip a re-read


# ---------------------------------------------------------------------------
# 2. product_transfer
# ---------------------------------------------------------------------------

def product_transfer(script_dir: str):
    file_path = os.path.join(script_dir, "FSL Product Transfer.xlsx")
    df = pd.read_excel(file_path)

    date_col = "Created Date"
    if date_col not in df.columns:
        print(f"Column '{date_col}' not found.")
        return

    df[date_col] = df[date_col].astype(str).str.strip()
    invalid = df[df[date_col].isna()]
    print(f"Invalid or missing dates: {len(invalid)} rows.")
    df = df.dropna(subset=[date_col])

    df[date_col] = pd.to_datetime(df[date_col], dayfirst=True)
    df = df.sort_values(by=date_col, ascending=False)
    df[date_col] = df[date_col].dt.strftime("%d/%m/%Y")

    df.to_excel(file_path, index=False)
    print(f"product_transfer saved: {file_path}")


# ---------------------------------------------------------------------------
# 3. lookup  (uses pre-loaded stock_data to avoid re-reading stock_inventory)
# ---------------------------------------------------------------------------

def lookup(script_dir: str, stock_data: pd.DataFrame):
    fsr_hub_path       = os.path.join(script_dir, "Output_File", "fsr_hub.xlsx")
    product_xfer_path  = os.path.join(script_dir, "FSL Product Transfer.xlsx")
    loc_map_path       = os.path.join(script_dir, "Location Mapping.xlsx")
    summary_path       = os.path.join(script_dir, "Forecast Model FSR.xlsb")
    output_path        = os.path.join(script_dir, "Output_File", "Forecast.xlsx")

    fsr_hub_data   = pd.read_excel(fsr_hub_path, engine="openpyxl")
    xfer_data      = pd.read_excel(product_xfer_path, engine="openpyxl")

    fsr_col   = "Product Item Number"
    xfer_col  = "Source Product Item: Product Item Number"
    dest_col  = "Destination Location: Location Name"

    if fsr_col not in fsr_hub_data.columns or xfer_col not in xfer_data.columns:
        print("Required columns missing in fsr_hub or product_transfer.")
        return

    # --- XLOOKUP: merge product transfer destination ---
    merged = pd.merge(
        fsr_hub_data,
        xfer_data[[xfer_col, dest_col]],
        left_on=fsr_col, right_on=xfer_col, how="left"
    ).drop(columns=[xfer_col])

    merged[dest_col] = merged[dest_col].fillna("mismatched")
    merged.to_excel(fsr_hub_path, index=False)
    print(f"fsr_hub updated with destination: {fsr_hub_path}")

    # --- VLOOKUP: map destination → location type (vectorized) ---
    fsr_data = pd.read_excel(fsr_hub_path, engine="openpyxl")

    loc_type_map = stock_data.drop_duplicates("Location: Location Name").set_index("Location: Location Name")["Location: Location Type"].to_dict()

    fsr_data["VLOOKUP_Result"] = fsr_data[dest_col].map(
        lambda colx: "mismatched" if colx == "mismatched" or pd.isna(colx)
        else loc_type_map.get(colx, "mismatched")
    )

    remove_types = {"HO", "Repair Center", "In House Repair Centre"}
    fsr_data = fsr_data[~fsr_data["VLOOKUP_Result"].isin(remove_types)]

    # fallback for double-mismatched rows
    na_cond = (fsr_data[dest_col] == "mismatched") & (fsr_data["VLOOKUP_Result"] == "mismatched")
    fsr_data.loc[na_cond, dest_col]          = fsr_data.loc[na_cond, "Location: Location Name"]
    fsr_data.loc[na_cond, "VLOOKUP_Result"]  = fsr_data.loc[na_cond, "Location: Location Type"]

    fsr_data = fsr_data[~fsr_data["VLOOKUP_Result"].isin(remove_types)]
    fsr_data.dropna(subset=[dest_col], inplace=True)

    # --- secondary VLOOKUP from FSR summary model ---
    summary_data = pd.read_excel(summary_path, engine="pyxlsb")
    lookup_dict  = summary_data.set_index("Location: Location Name")["Location:EmpCode"].to_dict()

    target_col = "VLOOKUP_Result"
    mismatch_mask = fsr_data[target_col] == "mismatched"
    fsr_data.loc[mismatch_mask, target_col] = fsr_data.loc[mismatch_mask, dest_col].map(lookup_dict)

    # anything not FSR/Hub/mismatched → FSR
    fsr_data[target_col] = np.where(
        fsr_data[target_col].isin(["FSR", "Hub", "mismatched"]),
        fsr_data[target_col], "FSR"
    )
    fsr_data = fsr_data[fsr_data[target_col] != "mismatched"].reset_index(drop=True)

    # clear columns E–L
    fsr_data[fsr_data.columns[4:12]] = ""

    fsr_data["Location: Location Name"] = fsr_data[dest_col]
    fsr_data["Location: Location Type"] = fsr_data[target_col]

    fsr_data = fsr_data[~fsr_data["Location: Location Name"].isin(
        ["Noida-HO – In-House Repair Centre", "Mumbai – In-House Repair Centre"]
    )]
    kerala_mask = fsr_data["Location: Location Name"] == "State Warehouse - Kerala"
    fsr_data.loc[kerala_mask, "Location: Location Type"] = "Hub"
    fsr_data.loc[kerala_mask, "Location: Location Name"] = "Stock Kerala"

    fsr_data.drop(columns=[target_col, dest_col], inplace=True)
    fsr_data.to_excel(fsr_hub_path, index=False)
    print(fsr_data)

    # --- Location Mapping: batch all 4 lookups in one pass ---
    loc_map = pd.read_excel(loc_map_path, engine="openpyxl")
    loc_map = loc_map.drop_duplicates(subset="Location Name")
    loc_map.to_excel(loc_map_path, index=False, engine="openpyxl")
    print(f"Location Mapping deduped rows: {len(loc_map)}")

    lm_idx = loc_map.set_index("Location Name")
    fsr_data["Location: Emp Code"]               = fsr_data["Location: Location Name"].map(lm_idx["Emp Code"])
    fsr_data["Location: City Master: City Name"]  = fsr_data["Location: Location Name"].map(lm_idx["City Master: City Name"])
    fsr_data["Location: State"]                   = fsr_data["Location: Location Name"].map(lm_idx["State"])
    fsr_data["Location: Region"]                  = fsr_data["Location: Location Name"].map(lm_idx["Region"])
    print("Location lookups applied:", len(fsr_data))

    fsr_data.to_excel(fsr_hub_path, index=False, na_rep="N/A")

    # --- split into InTransit FSR / Hub and write Forecast.xlsx ---
    fsr_dedup = fsr_data.drop_duplicates(subset="Product Item Number", keep="first")
    fsr_dedup["Location: Location Type"] = fsr_dedup["Location: Location Type"].str.strip()

    intransit_fsr = fsr_dedup[fsr_dedup["Location: Location Type"] == "FSR"]
    intransit_hub = fsr_dedup[fsr_dedup["Location: Location Type"] == "Hub"].drop_duplicates(subset="Product Item Number", keep="first")

    intransit_fsr = intransit_fsr[~intransit_fsr["Product Name: Product Name"].str.startswith("Mini", na=False)]

    _write_excel_sheets(output_path, {
        "InTransit FSR": intransit_fsr,
        "InTransit Hub": intransit_hub,
    }, mode="w")
    print(f"Forecast.xlsx initialised: {output_path}")


# ---------------------------------------------------------------------------
# 4–6. working / non_working / l1_l2_repair — all share stock_data
#       and write to Forecast.xlsx in a SINGLE ExcelWriter session
# ---------------------------------------------------------------------------

def working_nonworking_repair(script_dir: str, stock_data: pd.DataFrame):
    forecast_path = os.path.join(script_dir, "Output_File", "Forecast.xlsx")

    base_mask = (stock_data["In Transit?"] == 0) & (stock_data["Quantity On Hand"] == 1)
    ltype = stock_data["Location: Location Type"]
    status = stock_data["Product Working Condition Status"]

    # working
    fsr_working = stock_data[base_mask & (status == "Working") & (ltype == "FSR")]
    hub_working = stock_data[base_mask & (status == "Working") & (ltype == "Hub")]

    # non-working statuses
    nw_statuses = {"Non-Working", "L1 Checked Non working", "L1 Checked Non working damage", "Repair Center"}
    fsr_nw = stock_data[base_mask & status.isin(nw_statuses) & (ltype == "FSR")]
    hub_nw = stock_data[base_mask & status.isin(nw_statuses) & (ltype == "Hub")]

    # repair centres
    l1_repair = stock_data[base_mask & (ltype == "In House Repair Centre")]
    l2_repair = stock_data[base_mask & (ltype == "Repair Center")]

    _write_excel_sheets(forecast_path, {
        "FSR Working":      fsr_working,
        "hub Working":      hub_working,
        "FSR Non-Working":  fsr_nw,
        "hub Non-Working":  hub_nw,
        "l1 repair center": l1_repair,
        "l2 repair center": l2_repair,
    }, mode="a")
    print(f"Working / Non-Working / Repair sheets written: {forecast_path}")


# ---------------------------------------------------------------------------
# 7. mismatched_empcode
# ---------------------------------------------------------------------------

def mismatched_empcode(script_dir: str):
    file_path    = os.path.join(script_dir, "Output_File", "Forecast.xlsx")
    fsr_model    = os.path.join(script_dir, "Forecast Model FSR.xlsb")
    hub_model    = os.path.join(script_dir, "Forecast Model HUB.xlsb")

    # --- FSR summary emp codes ---
    fsr_summary = pd.read_excel(fsr_model, sheet_name="Summary", engine="pyxlsb")
    fsr_summary["Location:EmpCode"] = fsr_summary["Location:EmpCode"].astype(str).str.strip()
    fsr_emp_set = set(fsr_summary["Location:EmpCode"].dropna())

    # --- HUB summary (read once, all needed sheets) ---
    hub_xl = pd.ExcelFile(hub_model, engine="pyxlsb")
    hub_summary  = hub_xl.parse("Summary")
    model_l1     = hub_xl.parse("L1 Repair Center")
    model_l2     = hub_xl.parse("L2 Repair Center")

    hub_summary["Location:EmpCode"] = hub_summary["Location:EmpCode"].astype(str).str.strip()
    hub_summary["HUB Location"]     = hub_summary["HUB Location"].astype(str).str.strip()
    hub_emp_set  = set(hub_summary["Location:EmpCode"].dropna())
    hub_loc_set  = set(hub_summary["HUB Location"].dropna())

    emp_to_l1_name = dict(zip(model_l1["Location:EmpCode"], model_l1["Location: Location Name"]))
    emp_to_l2_name = dict(zip(model_l2["Location:EmpCode"], model_l2["Location: Location Name"]))

    # --- helper: filter by emp code set (vectorized) ---
    def _filter_fsr(sheet: str, key_col: str = "Location: Emp Code") -> pd.DataFrame:
        df = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl")
        df[key_col] = df[key_col].astype(str).str.strip()
        return df[df[key_col].isin(fsr_emp_set)]

    def _filter_hub(sheet: str, key_col: str = "Location: Location Name") -> pd.DataFrame:
        df = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl")
        df[key_col] = df[key_col].astype(str).str.strip()
        return df[df[key_col].isin(hub_loc_set)]

    fsr_intransit = _filter_fsr("InTransit FSR")
    fsr_working   = _filter_fsr("FSR Working")
    fsr_nw_all    = _filter_fsr("FSR Non-Working")

    pidf_mask_fsr      = fsr_nw_all["Product Working Condition Sub Status"].str.contains("P", case=False, na=False)
    fsr_nw_pidf        = fsr_nw_all[pidf_mask_fsr]
    fsr_nw             = fsr_nw_all[~pidf_mask_fsr]

    hub_intransit = _filter_hub("InTransit Hub")
    hub_intransit["Location: Emp Code"] = hub_intransit["Location: Emp Code"].replace("10796", "18156")
    hub_working   = _filter_hub("hub Working")
    hub_nw_all    = _filter_hub("hub Non-Working")

    pidf_mask_hub  = hub_nw_all["Product Working Condition Sub Status"].str.contains("P", case=False, na=False)
    hub_nw_pidf    = hub_nw_all[pidf_mask_hub]
    hub_nw         = hub_nw_all[~pidf_mask_hub]

    l1 = pd.read_excel(file_path, sheet_name="l1 repair center", engine="openpyxl")
    l1["Location: Location Name"] = l1["Location: Emp Code"].map(emp_to_l1_name)

    l2 = pd.read_excel(file_path, sheet_name="l2 repair center", engine="openpyxl")
    l2["Location: Location Name"] = l2["Location: Emp Code"].map(emp_to_l2_name)

    _write_excel_sheets(file_path, {
        "InTransit FSR":             fsr_intransit,
        "FSR Working":               fsr_working,
        "FSR Non-Working":           fsr_nw,
        "FSR Non working PIDF stock": fsr_nw_pidf,
        "InTransit Hub":             hub_intransit,
        "hub Working":               hub_working,
        "hub Non-Working":           hub_nw,
        "l1 repair center":          l1,
        "l2 repair center":          l2,
        "Hub Non working PIDF stock": hub_nw_pidf,
    }, mode="a")
    print(f"mismatched_empcode done: {file_path}")


# ---------------------------------------------------------------------------
# 8. doable
# ---------------------------------------------------------------------------

def doable(script_dir: str):
    forecast_path   = os.path.join(script_dir, "Output_File", "Forecast.xlsx")
    fsr_model_path  = os.path.join(script_dir, "Forecast Model FSR.xlsb")
    loc_map_path    = os.path.join(script_dir, "Location Mapping.xlsx")

    pending     = pd.read_excel("Pending Installation.xlsx",  engine="openpyxl")
    fsr_map     = pd.read_excel("FSR Mapping.xlsx",           engine="openpyxl")
    fsr_model   = pd.read_excel(fsr_model_path,               engine="pyxlsb")
    loc_map     = pd.read_excel(loc_map_path,                 engine="openpyxl")
    hub_map     = pd.read_excel("Doable mapping HUB.xlsb",    sheet_name="FSR", engine="pyxlsb")   # read once

    # --- type coercion ---
    pending["Service Resource: Employee Code"] = pending["Service Resource: Employee Code"].astype(str)
    pending["Zip/Postal Code"]                 = pd.to_numeric(pending["Zip/Postal Code"], errors="coerce")
    fsr_map["Service Resource: Employee Code"] = fsr_map["Service Resource: Employee Code"].astype(str)
    fsr_map["Zip/Postal Code"]                 = pd.to_numeric(fsr_map["Zip/Postal Code"], errors="coerce")
    fsr_model["Location:EmpCode"]              = fsr_model["Location:EmpCode"].astype(str)

    # --- filter ---
    pending = pending[pending["Type of Work"] != "Terminal Replacement"]
    pending = pending[~pending["Work Order Line Item: Asset: Asset Name"].str.startswith("Mini", na=False)]
    pending["Work Order Line Item: Asset: Asset Name"] = (
        pending["Work Order Line Item: Asset: Asset Name"].str.split("_").str[0]
    )

    valid_emp_codes = set(fsr_model["Location:EmpCode"].dropna())

    def _mark_unmatched():
        """Vectorized unmatch check (replaces row-wise apply)."""
        col = pending["Service Resource: Employee Code"]
        pending["match"] = np.where(
            col.isna() | col.isin(["", "nan"]) | ~col.isin(valid_emp_codes),
            "Unmatch", "True"
        )

    _mark_unmatched()

    # --- lookup dicts ---
    zip_map   = fsr_map.drop_duplicates("Zip/Postal Code").set_index("Zip/Postal Code")["Service Resource: Employee Code"].to_dict()
    city_map  = fsr_map.set_index("City")["Service Resource: Employee Code"].to_dict()

    state_df  = fsr_map.dropna(subset=["Service Resource: Location: State", "Service Resource: Employee Code"])
    state_df  = state_df[state_df["Service Resource: Location: State"] != ""]
    state_map = state_df.set_index("Service Resource: Location: State")["Service Resource: Employee Code"].to_dict()

    city_state_map   = loc_map.set_index("City Master: City Name")["State"].to_dict()
    city_model_map   = fsr_model.set_index("Location")["Location:EmpCode"].to_dict()
    emp_loc_map      = loc_map.set_index("Emp Code")["Location Name"].to_dict()
    emp_region_map   = loc_map.set_index("Emp Code")["Region"].to_dict()
    emp_state_map    = loc_map.set_index("Emp Code")["State"].to_dict()

    pending["State/Province"] = pending["City"].map(city_state_map)

    def _fill_unmatch(source_col, lookup_dict):
        mask = pending["match"] == "Unmatch"
        pending.loc[mask, "Service Resource: Employee Code"] = pending.loc[mask, source_col].map(lookup_dict)
        _mark_unmatched()

    _fill_unmatch("Zip/Postal Code", zip_map)
    _fill_unmatch("City",            city_map)
    _fill_unmatch("State/Province",  state_map)
    _fill_unmatch("City",            city_model_map)

    pending["Location: Location Name"] = pending["Service Resource: Employee Code"].map(emp_loc_map).fillna("Unknown")
    pending["Case: Region"]            = pending["Service Resource: Employee Code"].map(emp_region_map)

    state_null = pending["State/Province"].isna()
    pending.loc[state_null, "State/Province"] = pending.loc[state_null, "Service Resource: Employee Code"].map(emp_state_map)

    # --- write Doable FSR ---
    _write_excel_sheets(forecast_path, {"Doable fsr": pending}, mode="a")

    # --- Doable Hub (hub_map already loaded) ---
    city_hub_map  = dict(zip(hub_map["Location"], hub_map["Hub Location"]))
    state_hub_map = dict(zip(hub_map["State"],    hub_map["Hub Location"]))

    pending["Service Resource: Employee Code"] = pending["City"].map(city_hub_map)
    null_emp = pending["Service Resource: Employee Code"].isna() | (pending["Service Resource: Employee Code"] == "")
    pending.loc[null_emp, "Service Resource: Employee Code"] = pending.loc[null_emp, "State/Province"].map(state_hub_map)

    _write_excel_sheets(forecast_path, {"Doable Hub": pending}, mode="a")
    print(pending.head())
    print(f"doable done: {forecast_path}")


# ---------------------------------------------------------------------------
# 9. forecast_model_file  — reads Forecast.xlsx once using ExcelFile
# ---------------------------------------------------------------------------

def forecast_model_file(script_dir: str):
    src_path    = os.path.join(script_dir, "Output_File", "Forecast.xlsx")
    fsr_out     = os.path.join(script_dir, "Output_File", "forecast_model_fsr.xlsx")
    hub_out     = os.path.join(script_dir, "Output_File", "forecast_model_hub.xlsx")

    xl = pd.ExcelFile(src_path, engine="openpyxl")

    def _s(name):
        return xl.parse(name)

    _write_excel_sheets(fsr_out, {
        "In-Transit":            _s("InTransit FSR"),
        "Available_Working":     _s("FSR Working"),
        "Non-Working":           _s("FSR Non-Working"),
        "Non working PIDF Stock":_s("FSR Non working PIDF stock"),
        "Doable":                _s("Doable fsr"),
    }, mode="w")

    _write_excel_sheets(hub_out, {
        "In-Transit":            _s("InTransit Hub"),
        "Available_Working":     _s("hub Working"),
        "Non-Working":           _s("hub Non-Working"),
        "Non working PIDF stock":_s("Hub Non working PIDF stock"),
        "Doable":                _s("Doable Hub"),
        "L1 Repair Center":      _s("l1 repair center"),
        "L2 Repair Center":      _s("l2 repair center"),
    }, mode="w")

    print("Sheets transferred to respective model files.")
    print("job Ends:", datetime.datetime.now().time())


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    _script_dir = os.path.dirname(os.path.abspath(__file__))

    # stock_inventory returns cleaned DataFrame → shared downstream (no re-read)
    _stock_data = stock_inventory(_script_dir)

    product_transfer(_script_dir)

    # lookup uses stock_data directly
    lookup(_script_dir, _stock_data)

    # working + non_working + l1_l2_repair merged into one write pass
    working_nonworking_repair(_script_dir, _stock_data)

    mismatched_empcode(_script_dir)
    doable(_script_dir)
    forecast_model_file(_script_dir)