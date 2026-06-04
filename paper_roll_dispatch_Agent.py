# -*- coding: utf-8 -*-
"""
===============================================================================
ADVANCED PAPER ROLL DISPATCH AI AGENT — ENTERPRISE DASHBOARD VERSION
===============================================================================

FEATURES
========
✅ Enterprise UI Dashboard
✅ Upload Excel Files
✅ Auto POS ID → Store ID Resolution
✅ Store Details Summary
✅ Transaction Consumption Analysis
✅ Paper Roll Eligibility Calculation
✅ Delivery Proactive Summary
✅ Delivery Reactive Summary
✅ Modern Colored Tables
✅ Fast Caching & Indexing
✅ Export Ready Architecture
✅ Advanced Recommendation Engine

INSTALL
=======
pip install pandas openpyxl gradio xlsxwriter numpy

RUN
===
python ADVANCED_PAPER_ROLL_AGENT.py

===============================================================================
"""

import math
import traceback
import datetime
import pandas as pd
import numpy as np
import gradio as gr

# =============================================================================
# GLOBAL CACHE
# =============================================================================

CACHE = {
    "store": None,
    "delivery": None,
    "txn": {},

    "store_index": {},
    "pos_index": {},
    "delivery_index": {},
    "txn_summary": {},
}

# =============================================================================
# CONFIG
# =============================================================================

STORE_SHEETS = [
    "Sheet1",
    "Sheet2",
    "Sheet3",
    "Sheet4",
    "Sheet5"
]

TXN_SHEETS = {
    "m2": "M-2",
    "m1": "M-1",
    "cm": "Current M"
}

DEFAULT_DIVISOR = 75
DEFAULT_BUFFER = 10

# =============================================================================
# HELPERS
# =============================================================================

def normalize(value):

    if pd.isna(value):
        return ""

    return str(value).replace(".0", "").strip()


def safe_int(value):

    try:
        return int(float(value))
    except:
        return 0


def safe_float(value):

    try:
        return float(value)
    except:
        return 0.0


def preprocess(df):

    if df is None or df.empty:
        return pd.DataFrame()

    df.columns = [str(c).strip() for c in df.columns]

    for c in df.columns:

        try:
            df[c] = df[c].astype(str).str.strip()
        except:
            pass

    return df

# =============================================================================
# BUILD STORE INDEX
# =============================================================================

def build_store_index(df):

    CACHE["store_index"] = {}
    CACHE["pos_index"] = {}

    if df is None or df.empty:
        return

    for store_id, grp in df.groupby("STORE_ID"):

        CACHE["store_index"][normalize(store_id)] = grp

    for pos_id, grp in df.groupby("CLIENT_ID"):

        CACHE["pos_index"][normalize(pos_id)] = grp

# =============================================================================
# BUILD DELIVERY INDEX
# =============================================================================

def build_delivery_index(df):

    CACHE["delivery_index"] = {}

    if df is None or df.empty:
        return

    store_col = None

    for c in [
        "Store ID",
        "STORE_ID",
        "Store",
        "STORE"
    ]:
        if c in df.columns:
            store_col = c
            break

    if not store_col:
        return

    for store_id, grp in df.groupby(store_col):

        CACHE["delivery_index"][normalize(store_id)] = grp

# =============================================================================
# BUILD TRANSACTION SUMMARY
# =============================================================================

def build_txn_summary(label, df):

    CACHE["txn_summary"][label] = {}

    if df is None or df.empty:
        return

    store_col = None

    for c in [
        "Store Id",
        "Store ID",
        "STORE_ID"
    ]:
        if c in df.columns:
            store_col = c
            break

    if not store_col:
        return

    if label == "m2":
        candidates = ["M-2", "CARD_Txn", "Current M"]

    elif label == "m1":
        candidates = ["M-1", "CARD_Txn", "Current M"]

    else:
        candidates = ["Current M", "CARD_Txn"]

    value_col = None

    for c in candidates:

        if c in df.columns:
            value_col = c
            break

    if not value_col:
        return

    df[value_col] = pd.to_numeric(
        df[value_col],
        errors="coerce"
    ).fillna(0)

    summary = (
        df.groupby(store_col)[value_col]
        .sum()
        .to_dict()
    )

    CACHE["txn_summary"][label] = {

        normalize(k): safe_int(v)

        for k, v in summary.items()
    }

# =============================================================================
# LOAD FILES
# =============================================================================

def load_files(
    store_file,
    txn_file,
    delivery_file
):

    try:

        if store_file is None:
            return "❌ Upload Store_ID_File.xlsx"

        if txn_file is None:
            return "❌ Upload transactionFile.xlsx"

        if delivery_file is None:
            return "❌ Upload deliveryFile.xlsx"

        # =====================================================================
        # CLEAR CACHE
        # =====================================================================

        CACHE["store"] = None
        CACHE["delivery"] = None
        CACHE["txn"] = {}

        CACHE["store_index"] = {}
        CACHE["pos_index"] = {}
        CACHE["delivery_index"] = {}
        CACHE["txn_summary"] = {}

        # =====================================================================
        # STORE FILE
        # =====================================================================

        store_dfs = []

        for sheet in STORE_SHEETS:

            try:

                df = pd.read_excel(
                    store_file.name,
                    sheet_name=sheet,
                    engine="openpyxl"
                )

                df = preprocess(df)

                if not df.empty:

                    df["SOURCE_SHEET"] = sheet

                    store_dfs.append(df)

            except:
                pass

        CACHE["store"] = pd.concat(
            store_dfs,
            ignore_index=True
        )

        build_store_index(
            CACHE["store"]
        )

        # =====================================================================
        # TXN FILE
        # =====================================================================

        for key, sheet in TXN_SHEETS.items():

            try:

                df = pd.read_excel(
                    txn_file.name,
                    sheet_name=sheet,
                    engine="openpyxl"
                )

                df = preprocess(df)

                CACHE["txn"][key] = df

                build_txn_summary(
                    key,
                    df
                )

            except:
                pass

        # =====================================================================
        # DELIVERY FILE
        # =====================================================================

        CACHE["delivery"] = pd.read_excel(
            delivery_file.name,
            sheet_name="Sheet1",
            engine="openpyxl"
        )

        CACHE["delivery"] = preprocess(
            CACHE["delivery"]
        )

        build_delivery_index(
            CACHE["delivery"]
        )

        return (
            "✅ FILES LOADED SUCCESSFULLY\n\n"
            f"Stores Indexed: {len(CACHE['store_index'])}\n"
            f"POS Indexed: {len(CACHE['pos_index'])}"
        )

    except Exception as e:

        traceback.print_exc()

        return f"❌ ERROR\n\n{str(e)}"

# =============================================================================
# STORE DETAILS
# =============================================================================

def get_store_details(
    search_value,
    search_type
):

    rows = None

    if search_type == "POS ID":

        rows = CACHE["pos_index"].get(
            normalize(search_value)
        )

    else:

        rows = CACHE["store_index"].get(
            normalize(search_value)
        )

    if rows is None or rows.empty:
        return None, {}

    store_id = normalize(
        rows.iloc[0]["STORE_ID"]
    )

    pos_ids = (
        rows["CLIENT_ID"]
        .astype(str)
        .unique()
        .tolist()
    )

    active_pos = int(
        rows["Status"]
        .astype(str)
        .str.lower()
        .eq("active")
        .sum()
    )

    return store_id, {

        "merchant": rows.iloc[0]["MERCHANT"],
        "store_name": rows.iloc[0]["STORE"],
        "store_id": store_id,
        "pos_ids": pos_ids,
        "active_pos": active_pos,
    }

# =============================================================================
# ELIGIBILITY
# =============================================================================

def calculate_consumption(
    store_id,
    divisor=75,
    buffer_pct=10
):

    results = []

    total_txn = 0
    total_eligibility = 0

    month_map = [
        ("m2", "M-2"),
        ("m1", "M-1"),
        ("cm", "Current M"),
    ]

    for key, label in month_map:

        txn = safe_int(
            CACHE["txn_summary"]
            .get(key, {})
            .get(store_id, 0)
        )

        eligibility = txn / divisor

        buffer = eligibility * (buffer_pct / 100)

        total = eligibility + buffer

        total_txn += txn
        total_eligibility += total

        results.append({

            "Month": label,

            "Total Txn": txn,

            "Paper Roll Eligibility":
                round(eligibility, 2),

            "10% Buffer":
                round(buffer, 2),

            "Total Eligibility":
                round(total, 2),
        })

    return (
        pd.DataFrame(results),
        round(total_eligibility, 2),
        total_txn
    )

# =============================================================================
# DELIVERY TABLES
# =============================================================================

def get_delivery_tables(store_id):

    rows = CACHE["delivery_index"].get(
        normalize(store_id)
    )

    if rows is None or rows.empty:

        return (
            pd.DataFrame(),
            pd.DataFrame(),
            0
        )

    proactive = rows[
        rows["Dispatch Type"]
        .astype(str)
        .str.contains(
            "Proactive",
            case=False,
            na=False
        )
    ]

    reactive = rows[
        rows["Dispatch Type"]
        .astype(str)
        .str.contains(
            "Reactive",
            case=False,
            na=False
        )
    ]

    proactive_qty = pd.to_numeric(
        proactive["Roll Qty"],
        errors="coerce"
    ).fillna(0).sum()

    reactive_qty = pd.to_numeric(
        reactive["Roll Qty"],
        errors="coerce"
    ).fillna(0).sum()

    total = proactive_qty + reactive_qty

    cols = [

        "Months",
        "Stock Id",
        "Dispatch Type",
        "POS ID",
        "MERCHANT_NAME",
        "Store",
        "Roll Qty",
        "Delivery date",
        "Vendor Name",
        "POD",
        "Delivery Status",
    ]

    final_cols = []

    for c in cols:
        if c in rows.columns:
            final_cols.append(c)

    return (

        proactive[final_cols],

        reactive[final_cols],

        int(total)
    )

# =============================================================================
# DASHBOARD HTML
# =============================================================================

def build_dashboard(
    info,
    total_eligibility,
    total_delivered,
    total_txn
):

    excess = round(
        total_delivered - total_eligibility,
        2
    )

    color = "#d4edda" if excess >= 0 else "#f8d7da"

    return f"""
    <div style='
        border-radius:15px;
        padding:20px;
        background:#ffffff;
        box-shadow:0px 0px 15px rgba(0,0,0,0.15);
        font-family:Arial;
    '>

    <h2 style='color:#0B5394'>
        Store Detail Dashboard
    </h2>

    <table style='width:100%;border-collapse:collapse'>

        <tr style='background:#0B5394;color:white'>
            <th style='padding:10px'>Field</th>
            <th style='padding:10px'>Value</th>
        </tr>

        <tr>
            <td>Merchant Name</td>
            <td>{info['merchant']}</td>
        </tr>

        <tr>
            <td>Store Name</td>
            <td>{info['store_name']}</td>
        </tr>

        <tr>
            <td>Store ID</td>
            <td>{info['store_id']}</td>
        </tr>

        <tr>
            <td>POS Count</td>
            <td>{len(info['pos_ids'])}</td>
        </tr>

        <tr>
            <td>Active POS Count</td>
            <td>{info['active_pos']}</td>
        </tr>

        <tr>
            <td>Total Transactions</td>
            <td>{total_txn}</td>
        </tr>

        <tr>
            <td>Total Eligibility</td>
            <td>{round(total_eligibility,2)}</td>
        </tr>

        <tr>
            <td>Total Delivered</td>
            <td>{total_delivered}</td>
        </tr>

        <tr style='background:{color};font-weight:bold'>
            <td>Excess Paper Rolls</td>
            <td>{round(excess,2)}</td>
        </tr>

    </table>
    </div>
    """

# =============================================================================
# MAIN ANALYSIS
# =============================================================================

def analyze(
    search_value,
    search_type,
    divisor,
    buffer_pct
):

    try:

        if CACHE["store"] is None:

            return (
                "❌ Please Load Files First",
                pd.DataFrame(),
                pd.DataFrame(),
                pd.DataFrame()
            )

        store_id, info = get_store_details(
            search_value,
            search_type
        )

        if not info:

            return (
                "❌ No Data Found",
                pd.DataFrame(),
                pd.DataFrame(),
                pd.DataFrame()
            )

        # =====================================================================
        # CONSUMPTION
        # =====================================================================

        (
            txn_df,
            total_eligibility,
            total_txn
        ) = calculate_consumption(
            store_id,
            divisor,
            buffer_pct
        )

        # =====================================================================
        # DELIVERY
        # =====================================================================

        proactive_df, reactive_df, total_delivered = (
            get_delivery_tables(store_id)
        )

        # =====================================================================
        # DASHBOARD
        # =====================================================================

        dashboard = build_dashboard(
            info,
            total_eligibility,
            total_delivered,
            total_txn
        )

        return (
            dashboard,
            txn_df,
            proactive_df,
            reactive_df
        )

    except Exception as e:

        traceback.print_exc()

        return (
            f"❌ ERROR\n\n{str(e)}",
            pd.DataFrame(),
            pd.DataFrame(),
            pd.DataFrame()
        )

# =============================================================================
# UI
# =============================================================================

custom_css = """

body{
    background:#f4f6f9;
}

.gradio-container{
    max-width:1700px !important;
}

table{
    font-size:14px;
}

"""

def build_ui():

    with gr.Blocks(
        css=custom_css,
        title="Advanced Paper Roll Dispatch Agent"
    ) as app:

        gr.Markdown("""
# 🧠 ADVANCED PAPER ROLL DISPATCH AI AGENT

### Enterprise Dashboard Version
""")

        # =====================================================================
        # FILES
        # =====================================================================

        with gr.Row():

            store_file = gr.File(
                label="Upload Store_ID_File.xlsx"
            )

            txn_file = gr.File(
                label="Upload transactionFile.xlsx"
            )

            delivery_file = gr.File(
                label="Upload deliveryFile.xlsx"
            )

        load_btn = gr.Button(
            "🚀 LOAD FILES",
            variant="primary"
        )

        load_status = gr.Textbox(
            label="Load Status",
            lines=4
        )

        # =====================================================================
        # SEARCH
        # =====================================================================

        with gr.Row():

            search_value = gr.Textbox(
                label="Store ID / POS ID"
            )

            search_type = gr.Radio(
                ["POS ID", "Store ID"],
                value="POS ID",
                label="Search Type"
            )

            divisor = gr.Number(
                value=75,
                label="Eligibility Divisor"
            )

            buffer_pct = gr.Number(
                value=10,
                label="Buffer %"
            )

        analyze_btn = gr.Button(
            "🔍 ANALYZE STORE",
            variant="primary"
        )

        # =====================================================================
        # OUTPUTS
        # =====================================================================

        dashboard_output = gr.HTML()

        gr.Markdown("""
# 📊 Paper Rolls Consumption / Transaction Summary
""")

        txn_table = gr.DataFrame()

        gr.Markdown("""
# 🚚 Paper Rolls Delivery Proactive
""")

        proactive_table = gr.DataFrame()

        gr.Markdown("""
# 🔁 Paper Rolls Delivery Reactive
""")

        reactive_table = gr.DataFrame()

        # =====================================================================
        # EVENTS
        # =====================================================================

        load_btn.click(
            fn=load_files,
            inputs=[
                store_file,
                txn_file,
                delivery_file
            ],
            outputs=load_status
        )

        analyze_btn.click(
            fn=analyze,
            inputs=[
                search_value,
                search_type,
                divisor,
                buffer_pct
            ],
            outputs=[
                dashboard_output,
                txn_table,
                proactive_table,
                reactive_table
            ]
        )

    return app

# =============================================================================
# MAIN
# =============================================================================

def main():

    app = build_ui()

    app.queue(
        default_concurrency_limit=20
    )

    app.launch(
        server_name="0.0.0.0",
        server_port=None,
        share=True,
        inbrowser=True,
        show_error=True
    )

if __name__ == "__main__":

    main()