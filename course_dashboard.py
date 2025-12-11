import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Course Completion Dashboard", layout="wide")

st.title("📘 Course Completion Dashboard")

uploaded_file = st.file_uploader("Upload Excel with Multiple Sheets", type=["xlsx"])

if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    combined_df = pd.DataFrame()

    # Helper lists for flexible column detection
    possible_name_cols = ["Employee Name", "Employee Name ", "EmployeeName", "Name of the Official", "Name", "Employee", "Name of Official"]
    possible_empno_cols = ["Employee No.", "Employee No", "Employee No ", "EmployeeNo", "Emp No", "Employee Number"]
    # We'll detect division column by looking for 'division' or 'unit' keywords
    for sheet in xls.sheet_names:

        # IMPORTANT: use header=1 because column names are in the 2nd row
        df = pd.read_excel(uploaded_file, sheet_name=sheet, header=1)

        # strip column names and remove full-empty columns (this also clears many Unnamed: columns)
        df.columns = df.columns.astype(str).str.strip()
        df = df.dropna(axis=1, how="all")  # drop columns where every value is NaN

        # Also drop columns that are literally 'Unnamed: ...' or empty names
        drop_cols = [c for c in df.columns if (c.lower().startswith("unnamed") or c.strip() == "")]
        if drop_cols:
            df = df.drop(columns=drop_cols, errors="ignore")

        # --- find employee name column ---
        emp_name_col = None
        for col in possible_name_cols:
            if col in df.columns:
                emp_name_col = col
                break
        # fallback: any column containing both 'employee' and 'name' words
        if emp_name_col is None:
            for c in df.columns:
                if "employee" in c.lower() and "name" in c.lower():
                    emp_name_col = c
                    break
        # fallback: any column containing 'name' but not 'office' etc.
        if emp_name_col is None:
            for c in df.columns:
                if "name" in c.lower() and "office" not in c.lower():
                    emp_name_col = c
                    break

        # --- find employee no column ---
        emp_no_col = None
        for col in possible_empno_cols:
            if col in df.columns:
                emp_no_col = col
                break
        if emp_no_col is None:
            for c in df.columns:
                if ("emp" in c.lower() or "employee" in c.lower()) and ("no" in c.lower() or "number" in c.lower()):
                    emp_no_col = c
                    break

        # --- find division/unit column (for RMS TP) ---
        division_col = None
        for c in df.columns:
            if "division" in c.lower() or "division/" in c.lower() or "unit" in c.lower():
                division_col = c
                break

        # If no useful columns detected, skip sheet
        if emp_name_col is None or emp_no_col is None:
            # skip sheet if we cannot find names or emp no
            continue

        # Filter for RMS TP rows
        if division_col:
            # sometimes value may include extra spaces / non-breaking spaces — convert to str and search
            df_tp = df[df[division_col].astype(str).str.contains("RMS TP", case=False, na=False)]
        else:
            # as a fallback, search entire row for 'RMS TP'
            tp_mask = df.apply(lambda col: col.astype(str).str.contains("RMS TP", case=False, na=False))
            if tp_mask.any().any():
                df_tp = df[tp_mask.any(axis=1)]
            else:
                df_tp = pd.DataFrame()

        if df_tp.empty:
            continue

        # Add Course Name = sheet name
        df_tp["Course Name"] = sheet

        # Normalize Employee Name and Employee No column names in the combined df
        df_tp = df_tp.rename(columns={emp_name_col: "Employee Name", emp_no_col: "Employee No."})

        # Drop extra columns that are blank or not needed (already done above, but re-check)
        df_tp = df_tp.dropna(axis=1, how="all")

        combined_df = pd.concat([combined_df, df_tp], ignore_index=True)

    if combined_df.empty:
        st.error("No RMS TP data found in any sheet. Please check the file and that RMS TP rows exist.")
        st.stop()

    st.success("Data extracted successfully!")
    # Show a cleaned preview (drop fully-empty columns and the Unnamed columns if any remain)
    preview_df = combined_df.copy()
    preview_df = preview_df.dropna(axis=1, how="all")
    preview_df = preview_df[[c for c in preview_df.columns if not str(c).lower().startswith("unnamed")]]
    st.dataframe(preview_df)

    # -------------------------
    # Create pivot table
    # -------------------------
    st.subheader("📊 Pivot Table Summary")

    # Ensure required columns exist
    if "Employee Name" not in combined_df.columns or "Course Name" not in combined_df.columns:
        st.error("Required columns missing: 'Employee Name' or 'Course Name'.")
        st.stop()

    # For pivot values we use Employee No. if present otherwise use a placeholder column
    if "Employee No." not in combined_df.columns:
        # create a dummy employee no column for counting
        combined_df["Employee No."] = 1

    # Build pivot: rows = Employee Name, columns = Course Name, values = count of Employee No.
    pivot_df = combined_df.pivot_table(
        index="Employee Name",
        columns="Course Name",
        values="Employee No.",
        aggfunc="count",
        fill_value=0
    )

    # Add Grand Total column and row
    pivot_df["Grand Total"] = pivot_df.sum(axis=1)
    pivot_df.loc["Grand Total"] = pivot_df.sum(numeric_only=True)

    # Sort rows alphabetically (except Grand Total)
    if "Grand Total" in pivot_df.index:
        grand = pivot_df.loc["Grand Total"]
        pivot_df = pivot_df.drop(index="Grand Total").sort_index()
        pivot_df.loc["Grand Total"] = grand

    st.dataframe(pivot_df)

    # -------------------------
    # Export pivot to Excel with formatting
    # -------------------------
    def export_pivot_to_excel(df_to_export: pd.DataFrame) -> BytesIO:
        wb = Workbook()
        ws = wb.active
        ws.title = "Pivot Summary"

        # Write dataframe (dataframe_to_rows includes index)
        for r in dataframe_to_rows(df_to_export.reset_index(), index=False, header=True):
            ws.append(r)

        # Style header row (first row)
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        center_align = Alignment(horizontal="center", vertical="center")

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        # Bold the Grand Total row
        # find row index of "Grand Total"
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if row and str(row[0]).strip().lower() == "grand total":
                for cell in ws[r_idx]:
                    cell.font = Font(bold=True)
                    cell.alignment = center_align

        # Auto-size columns
        for column_cells in ws.columns:
            length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
            col_letter = column_cells[0].column_letter
            ws.column_dimensions[col_letter].width = min(max(length + 2, 10), 50)

        out = BytesIO()
        wb.save(out)
        out.seek(0)
        return out

    excel_bytes = export_pivot_to_excel(pivot_df)

    st.download_button(
        label="📥 Download Pivot Table (Excel)",
        data=excel_bytes.getvalue(),
        file_name="RMS_TP_Pivot_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
