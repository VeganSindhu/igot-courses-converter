import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="RMS TP Pending Courses", layout="wide")
st.title("📘 RMS TP – Pending Course Completion Dashboard")

uploaded_file = st.file_uploader(
    "Upload Excel file (each sheet = pending list of a course)",
    type=["xlsx"]
)

if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)

    pending_records = []
    employee_master = {}

    # Column names as per your Excel
    name_cols = ["Employee_name", "Employee Name"]
    office_cols = ["Office of working", "Office of Working"]
    division_cols = ["Division"]

    # ---------------- READ EACH SHEET ----------------
    for sheet in xls.sheet_names:

        # Header is in 2nd row
        df = pd.read_excel(uploaded_file, sheet_name=sheet, header=1)

        # Clean columns
        df.columns = df.columns.astype(str).str.strip()
        df = df.dropna(axis=1, how="all")
        df = df[[c for c in df.columns if not c.lower().startswith("unnamed")]]

        emp_name_col = next((c for c in name_cols if c in df.columns), None)
        office_col = next((c for c in office_cols if c in df.columns), None)
        division_col = next((c for c in division_cols if c in df.columns), None)

        if not emp_name_col or not division_col:
            continue

        # ---------------- FILTER ONLY RMS TP ----------------
        df = df[
            df[division_col]
            .astype(str)
            .str.strip()
            .str.upper()
            .eq("RMS TP")
        ]

        if df.empty:
            continue

        # ---------------- PROCESS PENDING EMPLOYEES ----------------
        for _, row in df.iterrows():
            emp_name = str(row[emp_name_col]).strip()

            if not emp_name:
                continue

            # Master employee info (stored once)
            if emp_name not in employee_master:
                employee_master[emp_name] = {
                    "Employee Name": emp_name,
                    "Office of Working": row.get(office_col, "")
                }

            pending_records.append({
                "Employee Name": emp_name,
                "Course": sheet,
                "Pending": 1
            })

    if not pending_records:
        st.error("No RMS TP pending data found in the uploaded Excel.")
        st.stop()

    pending_df = pd.DataFrame(pending_records)
    master_df = pd.DataFrame(employee_master.values())

    # ---------------- BUILD COURSE MATRIX ----------------
    matrix_df = pending_df.pivot_table(
        index="Employee Name",
        columns="Course",
        values="Pending",
        aggfunc="max",
        fill_value=0
    ).reset_index()

    final_df = master_df.merge(matrix_df, on="Employee Name", how="left")
    final_df = final_df.fillna(0)

    # ---------------- TOTAL COURSES ----------------
    course_cols = [
        c for c in final_df.columns
        if c not in ["Employee Name", "Office of Working"]
    ]

    final_df["Total Courses"] = final_df[course_cols].sum(axis=1)

    # ---------------- SORT DESCENDING ----------------
    final_df = final_df.sort_values(
        by="Total Courses",
        ascending=False
    ).reset_index(drop=True)

    st.success("✅ RMS TP pending course matrix generated")
    st.dataframe(final_df)

    # ---------------- EXPORT TO EXCEL ----------------
    def export_to_excel(df):
        wb = Workbook()
        ws = wb.active
        ws.title = "RMS TP Pending Courses"

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(bold=True, color="FFFFFF")
        center = Alignment(horizontal="center", vertical="center")

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center

        for col in ws.columns:
            width = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(width + 3, 45)

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    excel_bytes = export_to_excel(final_df)

    st.download_button(
        "📥 Download RMS TP Pending Course Report",
        excel_bytes.getvalue(),
        "RMS_TP_Pending_Courses.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
