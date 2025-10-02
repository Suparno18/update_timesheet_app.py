import streamlit as st
import pandas as pd

st.title("Timesheet Updater App")

st.markdown("""
Upload the two Excel files:
- Monthly Output: "Bizx Murari monthly output-Suparno team.xlsx"
- Timesheet: "FIG - Bizx timesheet - Sep 2025.xlsx"

The app will update daily hours and recalculate totals.
""")

monthly_file = st.file_uploader("Upload Monthly Output Excel", type=["xlsx"])
timesheet_file = st.file_uploader("Upload Timesheet Excel", type=["xlsx"])

if monthly_file and timesheet_file:
    # Read files
    df_monthly = pd.read_excel(monthly_file, sheet_name="Sheet1", header=6)  # Header at row7 (0-index 6)
    df_timesheet = pd.read_excel(timesheet_file, sheet_name="Timesheet Report Day Wise (28)")

    # Clean columns
    df_monthly.columns = [col.replace('\n', ' ') for col in df_monthly.columns]
    daily_cols = [col for col in df_monthly.columns if col.isdigit() and 45901 <= int(col) <= 45930]

    # Process timesheet
    df_timesheet['EmployeeNo'] = df_timesheet['EmployeeNo'].astype(str)
    df_timesheet['Date'] = df_timesheet['Date'].astype(int)
    df_timesheet['Hours'] = df_timesheet['Hours'].astype(float)
    df_filtered = df_timesheet[df_timesheet['Category'].isin(['Worked', 'Holiday'])]
    df_grouped = df_filtered.groupby(['EmployeeNo', 'Date'])['Hours'].sum().reset_index()

    # Update monthly
    df_monthly['Emp No'] = df_monthly['Emp No'].astype(str)
    for idx, row in df_monthly.iterrows():
        emp_no = row['Emp No']
        if pd.isna(emp_no) or not emp_no:
            continue
        emp_hours = df_grouped[df_grouped['EmployeeNo'] == emp_no]
        for col in daily_cols:
            serial = int(col)
            day_hours = emp_hours[emp_hours['Date'] == serial]['Hours'].sum()
            df_monthly.at[idx, col] = day_hours if day_hours > 0 else ''

    # Recalculate totals
    df_monthly['Total Hours'] = df_monthly[daily_cols].sum(axis=1, numeric_only=True)

    # Weekly (assuming columns positions; adjust if needed)
    week1 = df_monthly[daily_cols[0:5]].sum(axis=1, numeric_only=True)
    week2 = df_monthly[daily_cols[7:12]].sum(axis=1, numeric_only=True)
    week3 = df_monthly[daily_cols[14:19]].sum(axis=1, numeric_only=True)
    week4 = df_monthly[daily_cols[21:26]].sum(axis=1, numeric_only=True)
    week5 = df_monthly[daily_cols[28:30]].sum(axis=1, numeric_only=True)

    weekly_col_names = ['Week1', 'Week2', 'Week3', 'Week4', 'Week5']  # Rename as per your sheet
    df_monthly[weekly_col_names[0]] = week1
    df_monthly[weekly_col_names[1]] = week2
    df_monthly[weekly_col_names[2]] = week3
    df_monthly[weekly_col_names[3]] = week4
    df_monthly[weekly_col_names[4]] = week5

    df_monthly['Total Hours.1'] = df_monthly['Total Hours']  # Duplicate if needed
    df_monthly['Billable Hours'] = df_monthly['Total Hours']
    df_monthly['Total No of Working Days'] = (df_monthly[daily_cols] > 0).sum(axis=1)

    # Download button
    output = pd.ExcelWriter('updated_monthly_output.xlsx', engine='openpyxl')
    df_monthly.to_excel(output, index=False, sheet_name='Sheet1')
    output.close()

    with open('updated_monthly_output.xlsx', 'rb') as f:
        st.download_button("Download Updated Excel", f, file_name="updated_monthly_output.xlsx")
