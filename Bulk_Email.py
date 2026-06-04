# -*- coding: utf-8 -*-
"""
Created on Wed Mar 25 18:44:43 2026

@author: abhishek.a
"""

# -*- coding: utf-8 -*-

import requests
import pandas as pd
import os
import win32com.client as win32
from datetime import datetime

# ==============================
# CONFIG
# ==============================
CONFIG_FILE = r"C:\Users\abhishek.a\Downloads\E-Mail Automation Template(Bulk).xlsx"
OUTPUT_FOLDER = r"C:\Users\abhishek.a\Downloads"
MAX_ATTACHMENT_MB = 30

# ==============================
# SALESFORCE AUTH
# ==============================
client_id = '3MVG9pe2TCoA1Pf4AUXJnyzPmstC0CcF5l_8P5uP77_GADjrPL2kiZCmvc2zQpxs6anlBwbQa0byW5iaf0M23'
client_secret = '96A440D33246A148DC3FBDE4915F512FC9CC91CD08AC00DABD540B0249AECDE4'
username = 'integration.user.maple@pinelabs.com.prod'
password = 'Windows98'

auth_url = (
    "https://login.salesforce.com/services/oauth2/token"
    f"?grant_type=password&client_id={client_id}"
    f"&client_secret={client_secret}"
    f"&username={username}&password={password}"
)

auth_response = requests.post(auth_url)
auth_response.raise_for_status()
auth_data = auth_response.json()

access_token = auth_data["access_token"]
instance_url = auth_data["instance_url"]

print("✅ Authenticated")

# ==============================
# LOAD CONFIG EXCEL
# ==============================
config_df = pd.read_excel(CONFIG_FILE)

# Clean column names
config_df.columns = config_df.columns.str.strip()

# ==============================
# PIVOT HTML FUNCTION (KEEP SAME)
# ==============================
def build_pivot_html(df, to_email):

    location = df["Location.Name"].iloc[0]

    order = [
        "Terminal(POS)", "Material", "Adapter", "Cable", "Battery",
        "Biometric", "Soundbox", "Stand", "Base", "Paper POS",
        "SIM", "Tool Kit", "Paper Rolls", "POSM Kit Stickers",
        "Back Cover", "Sticker", "Pendrive", "POS Stickers", "Tent Card"
    ]

    df["Product_Type__c"] = df["Product_Type__c"].fillna("Others")

    html = f"""
    <p style="font-family:Calibri;font-size:14px;">
    Dear {to_email},<br><br>
    Please find the stock inventory with <b>{location}</b>.<br><br>
    </p>

    <table style="border-collapse:collapse;font-family:Calibri;font-size:14px;width:900px;">
    <tr style="background-color:#B7DEE8;font-weight:bold;border-bottom:3px solid #1F4E79;">
        <td colspan="3" style="padding:10px;">
            <b>Location: Location Name</b> &nbsp;&nbsp;&nbsp;&nbsp; {location}
        </td>
    </tr>

    <tr style="background-color:#D9E1F2;font-weight:bold;">
        <th style="border:1px solid #000;padding:8px;">Product Type</th>
        <th style="border:1px solid #000;padding:8px;">Product Name</th>
        <th style="border:1px solid #000;padding:8px;text-align:right;">Quantity</th>
    </tr>
    """

    grand_total = 0

    for category in order:

        group = df[df["Product_Type__c"] == category]
        if group.empty:
            continue

        group = group.sort_values(by="Product2.Name")

        category_total = 0

        html += f"""
        <tr style="background:#BFBFBF;font-weight:bold;">
            <td style="border:1px solid #000;">{category}</td>
            <td style="border:1px solid #000;"></td>
            <td style="border:1px solid #000;"></td>
        </tr>
        """

        for _, row in group.iterrows():
            qty = int(row["QuantityOnHand"])
            category_total += qty

            html += f"""
            <tr>
                <td style="border-left:1px solid #000;border-right:1px solid #000;font-weight:bold;">
                    {category}
                </td>
                <td style="border-right:1px solid #000;">
                    {row['Product2.Name']}
                </td>
                <td style="border-right:1px solid #000;text-align:right;">
                    {qty}
                </td>
            </tr>
            """

        html += f"""
        <tr style="background:#6F6FAE;color:white;font-weight:bold;">
            <td style="border:1px solid #000;">{category} Total</td>
            <td style="border:1px solid #000;"></td>
            <td style="border:1px solid #000;text-align:right;">{category_total}</td>
        </tr>
        """

        html += "<tr><td colspan='3' style='height:10px;'></td></tr>"

        grand_total += category_total

    html += f"""
    <tr style="background:#A6A6A6;font-weight:bold;">
        <td style="border:1px solid #000;">Grand Total</td>
        <td style="border:1px solid #000;"></td>
        <td style="border:1px solid #000;text-align:right;">{grand_total}</td>
    </tr>
    </table>
    """

    return html

# ==============================
# OUTLOOK INIT
# ==============================
outlook = win32.Dispatch("Outlook.Application")

# ==============================
# MAIN LOOP (CONFIG DRIVEN)
# ==============================
for _, cfg in config_df.iterrows():

    try:
        loc_code = str(cfg["Location.Emp_Code__c"]).strip()
        to_email = cfg["TO_EMAIL"]
        cc_email = cfg.get("CC_EMAIL", "")
        subject = cfg.get("SUBJECT", f"Stock Report - {loc_code}")

        print(f"\n📍 Processing Location: {loc_code}")

        # --------------------------
        # FETCH DATA
        # --------------------------
        query = f"""
        SELECT Location.Emp_Code__c, Location.Name,
               Product_Type__c, Product2.Name, QuantityOnHand
        FROM ProductItem
        WHERE Location.Emp_Code__c = '{loc_code}' AND QuantityOnHand > 0
        """

        url = f"{instance_url}/services/data/v58.0/query"
        headers = {"Authorization": f"Bearer {access_token}"}

        res = requests.get(url, headers=headers, params={"q": query})
        res.raise_for_status()

        records = res.json().get("records", [])

        if not records:
            print(f"⚠️ No data for {loc_code}")
            continue

        df = pd.json_normalize(records)

        # --------------------------
        # BUILD EMAIL BODY
        # --------------------------
        html_body = build_pivot_html(df, to_email)

        # --------------------------
        # EXPORT FILE
        # --------------------------
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = os.path.join(OUTPUT_FOLDER, f"Stock_{loc_code}_{timestamp}.xlsx")
        df.to_excel(file_path, index=False)

        # --------------------------
        # SEND EMAIL
        # --------------------------
        mail = outlook.CreateItem(0)

        mail.To = to_email
        mail.CC = cc_email
        mail.Subject = subject
        mail.HTMLBody = html_body

        if os.path.exists(file_path):
            size_mb = os.path.getsize(file_path)/(1024*1024)
            if size_mb <= MAX_ATTACHMENT_MB:
                mail.Attachments.Add(file_path)

        mail.Send()

        print(f"✅ Email sent for {loc_code}")

    except Exception as e:
        print(f"❌ Error for {loc_code}: {str(e)}")