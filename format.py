import streamlit as st
import pandas as pd
import io
import difflib
import os

st.set_page_config(page_title="Inventory Formatter", layout="wide")
st.title("📦 Inventory Excel Formatter")
st.write("Upload your Excel file to extract and format the inventory report.")

uploaded_file = st.file_uploader("Upload Excel (.xls, .xlsx, .xlsm)", type=["xls", "xlsx", "xlsm"])

if uploaded_file:
    try:
        # Detect extension
        ext = os.path.splitext(uploaded_file.name)[1]
        engine = "xlrd" if ext == ".xls" else "openpyxl"

        # Load file
        df = pd.read_excel(uploaded_file, engine=engine)
        df.columns = [col.strip() for col in df.columns]

        st.subheader("📥 Raw Data Preview")
        st.dataframe(df.head())

        # Required columns
        expected_cols = ["Warehouse", "Code", "Eng Name", "Quantity", "Cost"]
        col_map = {}

        # Fuzzy match columns if needed
        for col in expected_cols:
            match = difflib.get_close_matches(col, df.columns, n=1, cutoff=0.6)
            if match:
                col_map[col] = match[0]
        if len(col_map) != len(expected_cols):
            st.error(f"❌ Missing columns. Found: {list(col_map.values())}")
            st.stop()
        df = df.rename(columns=col_map)

        # Rename and keep only relevant columns
        df = df[["Warehouse", "Code", "Eng Name", "Quantity", "Cost"]]
        df = df.rename(columns={"Cost": "USD Each"})

        # Backup original cost
        df["Original Cost"] = df["USD Each"]

        # Constants
        special_warehouses = ["DONA - RGA Warehouse", "DONA - Scrap Warehouse", "NOT FOR SALE", "INVOICED"]
        adjust_codes = ["029813261", "NA029813261"]
        price_029813 = 1398.06
        price_91v2 = 7752.42

        # Logic: update cost
        df.loc[df["Code"].isin(adjust_codes), "USD Each"] = price_029813
        df.loc[df["Code"].astype(str).str.strip() == "91V2NU000001", "USD Each"] = price_91v2

        # Notes + Total computation
        notes_col = []
        total_col = []
        for idx, row in df.iterrows():
            code = row["Code"]
            qty = row["Quantity"]
            warehouse = row["Warehouse"]
            orig_cost = row["Original Cost"]
            usd_each = row["USD Each"]
            usd_val = float(usd_each if isinstance(usd_each, (float, int)) else str(usd_each).replace("$", "").replace(",", ""))

            # Total value or zero override
            if warehouse in special_warehouses:
                total_col.append("$0.00")
            else:
                total_val = qty * usd_val
                total_col.append("")  # Leave it empty — will insert formula later

            # Notes logic
            note = ""
            if code in adjust_codes:
                cost_text = f"${float(orig_cost):,.2f}"
                note = f"Compiere previously showed {cost_text}"
            elif str(code).strip() == "91V2NU000001":
                cost_text = f"${float(orig_cost):,.2f}"
                note = f"Compiere is showing a cost of {cost_text}; See DONA-PO-10001758"
            notes_col.append(note)

        df["Total USD Each"] = total_col
        df["Notes (Ex. Rate 1.12)"] = notes_col

        # Final formatting
        df = df[["Warehouse", "Code", "Eng Name", "Quantity", "USD Each", "Total USD Each", "Notes (Ex. Rate 1.12)"]]

        st.subheader("✅ Processed Preview")
        st.dataframe(df.head())

        # Excel export
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Formatted", startrow=1, header=False)

            workbook = writer.book
            worksheet = writer.sheets["Formatted"]

            # Write headers manually
            for col_num, header in enumerate(df.columns):
                worksheet.write(0, col_num, header)

            # Column indexes
            qty_col = df.columns.get_loc("Quantity")
            usd_col = df.columns.get_loc("USD Each")
            total_col_idx = df.columns.get_loc("Total USD Each")

            # Formats
            center_format = workbook.add_format({'align': 'center'})
            currency_format = workbook.add_format({'num_format': '$#,##0.00', 'align': 'center'})

            # Write rows with formatting and formula logic
            for row_num, row in df.iterrows():
                excel_row = row_num + 1

                worksheet.write_number(excel_row, qty_col, row["Quantity"], center_format)
                worksheet.write_string(excel_row, usd_col, f"${float(row['USD Each']):,.2f}", center_format)

                # Special warehouse → $0.00 string
                if row["Total USD Each"] == "$0.00":
                    worksheet.write_string(excel_row, total_col_idx, "$0.00", center_format)
                else:
                    # Insert Excel formula: =D2*E2
                    q_col_letter = chr(ord("A") + qty_col)
                    u_col_letter = chr(ord("A") + usd_col)
                    t_formula = f"={q_col_letter}{excel_row+1}*{u_col_letter}{excel_row+1}"
                    worksheet.write_formula(excel_row, total_col_idx, t_formula, currency_format)

            # Add Excel table with formatting
            last_col = chr(ord('A') + len(df.columns) - 1)
            worksheet.add_table(f"A1:{last_col}{len(df)+1}", {
                'columns': [{'header': col} for col in df.columns],
                'style': 'Table Style Medium 15',
                'header_row': True
            })

            # Auto-adjust column widths based on max content length
            for i, col in enumerate(df.columns):
                series = df[col].astype(str)
                max_length = max(series.map(len).max(), len(col)) + 2  # Add padding
                worksheet.set_column(i, i, max_length)


        st.download_button("📥 Download Formatted Excel", output.getvalue(),
                           file_name="formatted_inventory.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
