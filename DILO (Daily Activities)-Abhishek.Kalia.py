# -*- coding: utf-8 -*-
"""
Created on Sat Dec  6 18:27:20 2025

@author: abhishek.a
"""

import pandas as pd
from datetime import datetime, timedelta

# =====================================================================
# SHEET 1: WEEKLY WISE ACTIVITIES
# =====================================================================

weekly_data = [
    ["Week 1 (1–5 Apr)", "IMSPL Stock Summary Report", "Daily", 45, 6, 4.5, "Daily stock updates completed"],
    ["Week 1 (1–5 Apr)", "Product Transfer Pending Report", "Twice a Week", 210, 2, 7.0, "Completed as per plan"],
    ["Week 1 (1–5 Apr)", "MCC Deployment", "Twice a Week", 120, 2, 4.0, "Completed"],
    ["Week 1 (1–5 Apr)", "Preparing PC Dump for Active POS", "Once a Week", 240, 1, 4.0, "Completed"],

    ["Week 2 (7–12 Apr)", "Salesforce Inventory Report", "Once a Week", 180, 1, 3.0, "Completed"],
    ["Week 2 (7–12 Apr)", "Product Consumption Report (Non-Serialized)", "Twice a Week", 120, 2, 4.0, "Completed"],
    ["Week 2 (7–12 Apr)", "Water Fall Report", "Once a Week", 120, 1, 2.0, "Completed"],

    ["Week 3 (14–19 Apr)", "JIO SIM Inventory", "Once a Week", 120, 1, 2.0, "Completed"],
    ["Week 3 (14–19 Apr)", "PIDF Deployment", "Once a Week", 45, 1, 0.8, "Completed"],

    ["Week 4 (21–26 Apr)", "SIM Activated but Not Consumed by FSR", "Once a Week", 150, 1, 2.5, "Completed"],
    ["Week 4 (21–26 Apr)", "MTD Report – Paper Rolls", "Twice a Week", 120, 2, 4.0, "Completed"],
]

df_weekly = pd.DataFrame(
    weekly_data,
    columns=[
        "Week",
        "Activity Name",
        "Frequency",
        "Average Time (Mins)",
        "Occurrences",
        "Total Time (Hrs)",
        "Remarks"
    ]
)

# =====================================================================
# SHEET 2: MONTHLY ACTIVITIES SUMMARY
# =====================================================================

monthly_data = [
    ["April", "IMSPL Stock Summary Report", "Daily", 18.0, "Completed", "Completed all days"],
    ["April", "Product Transfer Pending Report", "Twice a Week", 8.0, "Completed", "On schedule"],
    ["May", "Salesforce Inventory Report", "Once a Week", 12.0, "Completed", "4 cycles done"],
    ["June", "SIM Inventory", "Monthly", 2.0, "Pending", "Awaiting update"],
    ["July", "Water Fall Report", "Once a Week", 8.0, "Completed", "Completed every Friday"],
    ["Aug", "MTD Report – Paper Rolls", "Twice a Week", 8.0, "Completed", "Regular"],
    ["Sept", "MCC Deployment", "Twice a Week", 16.0, "Completed", "Full"],
    ["Oct", "Reactive Doable (MTD & Till Date)", "Once a Week", 4.0, "Completed", "Ongoing"],
    ["Nov", "Preparing PC Dump for Active POS", "Once a Week", 16.0, "Completed", "Completed"],
    ["Dec", "Summary & Closure Reports", "Once a Week", 12.0, "Completed", "End of year closure"],
]

df_monthly = pd.DataFrame(
    monthly_data,
    columns=[
        "Month",
        "Activity Name",
        "Frequency",
        "Total Time (Hrs)",
        "Task Status",
        "Remarks"
    ]
)

# =====================================================================
# SHEET 3: DAY WISE ACTIVITIES
# =====================================================================

start_date = datetime(2025, 4, 1)
end_date = datetime(2025, 4, 10)

activities = [
    ("IMSPL Stock Summary Report", 45),
    ("Others Activity (CTQ/WDV/SIM Check)", 90),
    ("Product Transfer Pending Report", 210),
    ("MCC Deployment", 120),
    ("Preparing PC Dump for Active POS", 240),
    ("Salesforce Inventory Report", 180),
]

day_rows = []
current_date = start_date

while current_date <= end_date:
    # Monday–Saturday → 0–5
    if current_date.weekday() < 6:

        # Time variation logic: rotates between 6–9 hours
        total_time = 6 + (current_date.day % 3)

        status = "Completed" if total_time <= 8 else "Pending"
        remark = "Smooth day" if status == "Completed" else "Overloaded"

        # Sample first 3 activities
        activity_list = ", ".join([a[0] for a in activities[:3]])

        day_rows.append([
            current_date.strftime("%d-%b-%y"),
            current_date.strftime("%a"),
            activity_list,
            total_time,
            status,
            remark
        ])

    current_date += timedelta(days=1)

df_daily = pd.DataFrame(
    day_rows,
    columns=["Date", "Day", "Activities", "Total Time (Hrs)", "Status", "Remarks"]
)

# =====================================================================
# EXPORT TO EXCEL
# =====================================================================

output_file = "DILO_Report_Apr-Dec_2025.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df_weekly.to_excel(writer, sheet_name="Weekly Wise Activities", index=False)
    df_monthly.to_excel(writer, sheet_name="Monthly Activities", index=False)
    df_daily.to_excel(writer, sheet_name="Day Wise Activities", index=False)

print(f"✅ Excel file '{output_file}' has been created successfully.")
