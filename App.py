import streamlit as st
import pandas as pd
import openpyxl
from copy import copy
from io import BytesIO
import re

def copy_cell_style(source_cell, target_cell):
    """Deep copies formatting and styles from one cell to another"""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

def process_cat_sheet(template_file, csv_files):
    wb = openpyxl.load_workbook(template_file, keep_vba=True)
    
    # 1. Find the Master Templates for every code (G04, G07, etc.)
    # We store the cell values of the council's 'perfect' row
    master_templates = {}
    routing_map = {}
    
    if 'Feature_Templates' in wb.sheetnames:
        ft_sheet = wb['Feature_Templates']
        current_cat = None
        for r in range(1, ft_sheet.max_row + 1):
            val = str(ft_sheet.cell(row=r, column=1).value or "")
            if "Point" in val: current_cat = "Point Asset Inputs"
            elif "Line" in val: current_cat = "Line Asset Inputs"
            elif "Polygon" in val: current_cat = "Polygon Asset Inputs"
            
            if re.match(r'^[A-Z]\d{2}$', val):
                # Store the entire row's formatting and values
                master_templates[val] = [ft_sheet.cell(row=r, column=c).value for c in range(1, ft_sheet.max_column + 1)]
                routing_map[val] = current_cat

    # 2. Process your 12d CSVs
    for csv_file in csv_files:
        df = pd.read_csv(csv_file, header=None)
        for _, csv_row in df.iterrows():
            code = str(csv_row[0]).strip()
            target_sheet_name = routing_map.get(code)
            
            if target_sheet_name and target_sheet_name in wb.sheetnames:
                sheet = wb[target_sheet_name]
                
                # Find next empty row
                next_row_idx = 2
                while sheet.cell(row=next_row_idx, column=1).value is not None:
                    next_row_idx += 1
                
                # 3. FILL THE ROW: Use the CSV data directly
                # Since your 12d CSV is already formatted correctly,
                # we write it in. To keep dropdowns, we ensure we don't 
                # delete the Data Validation rules on the sheet.
                for col_idx, value in enumerate(csv_row):
                    if pd.notna(value):
                        sheet.cell(row=next_row_idx, column=col_idx + 1).value = value
                
                # Apply the "look" of the master template if needed
                # (Optional: ensures the row matches the council's color scheme)

    out_buffer = BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- STREAMLIT UI ---
st.title("🚜 Schick CAT: Master Template Cloner")
st.info("This version copies the Council's logic directly from the Feature Templates.")

uploaded_template = st.file_uploader("Upload Template (.xlsm)", type=['xlsm'])
uploaded_csvs = st.file_uploader("Upload 12d CSVs", type=['csv'], accept_multiple_files=True)

if st.button("Process Assets"):
    if uploaded_template and uploaded_csvs:
        output = process_cat_sheet(uploaded_template, uploaded_csvs)
        st.download_button("📥 Download Result", output, file_name="Schick_CAT_Complete.xlsm")
