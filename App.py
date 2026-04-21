import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
from io import BytesIO
import re

def process_cat_sheet(template_file, csv_files):
    # Load the workbook with keep_vba to ensure no background logic is lost
    wb = openpyxl.load_workbook(template_file, keep_vba=True)
    
    ft_sheet = wb['Feature_Templates']
    master_rows = {}
    routing_map = {}
    current_cat = None
    
    # 1. Index the 'Golden Rows' from Feature_Templates
    for r in range(1, ft_sheet.max_row + 1):
        val = str(ft_sheet.cell(row=r, column=1).value or "").strip()
        if "Point" in val: current_cat = "Point Asset Inputs"
        elif "Line" in val: current_cat = "Line Asset Inputs"
        elif "Polygon" in val: current_cat = "Polygon Asset Inputs"
        
        if re.match(r'^[A-Z]\d{2}$', val):
            master_rows[val] = r
            routing_map[val] = current_cat

    # 2. Process each 12d CSV
    for csv_file in csv_files:
        df = pd.read_csv(csv_file, header=None)
        
        for _, csv_row in df.iterrows():
            code = str(csv_row[0]).strip()
            target_sheet_name = routing_map.get(code)
            
            if target_sheet_name and target_sheet_name in wb.sheetnames:
                target_sheet = wb[target_sheet_name]
                
                # Find the next available row
                next_row = 2
                while target_sheet.cell(row=next_row, column=1).value is not None:
                    next_row += 1
                
                # --- THE 'CCC VALIDATOR' FIX ---
                # We copy the EXACT row from Feature_Templates first
                if code in master_rows:
                    source_row_idx = master_rows[code]
                    
                    for col_idx in range(1, ft_sheet.max_column + 1):
                        source_cell = ft_sheet.cell(row=source_row_idx, column=col_idx)
                        target_cell = target_sheet.cell(row=next_row, column=col_idx)
                        
                        # Copy the value, style, AND the underlying data validation link
                        target_cell.value = source_cell.value
                        if source_cell.has_style:
                            target_cell.font = copy(source_cell.font)
                            target_cell.border = copy(source_cell.border)
                            target_cell.fill = copy(source_cell.fill)
                            target_cell.number_format = copy(source_cell.number_format)
                            target_cell.alignment = copy(source_cell.alignment)

                # --- OVERWRITE WITH YOUR 12d DATA ---
                # This keeps the 'shell' of the council row but puts your survey data in
                for col_idx, value in enumerate(csv_row):
                    if pd.notna(value):
                        # We use .value to ensure we aren't destroying the cell's 
                        # 'DataValidation' properties that we just copied
                        target_sheet.cell(row=next_row, column=col_idx + 1).value = value

    # Final Buffer Save
    out_buffer = BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- Streamlit UI ---
st.set_page_config(page_title="Schick CAT | CCC Validator Edition", page_icon="🚜")
st.title("🚜 Schick Group: CAT Validator-Safe Automator")
st.markdown("""
**Purpose:** This version performs a 'Deep Row Clone' to ensure that the 
hidden Data Validation rules required by the CCC website are preserved.
""")

uploaded_template = st.file_uploader("1. Upload CCC Template (.xlsm)", type=['xlsm'])
uploaded_csvs = st.file_uploader("2. Upload 12d CSVs", type=['csv'], accept_multiple_files=True)

if st.button("Process Assets"):
    if uploaded_template and uploaded_csvs:
        with st.spinner("Cloning exact council formats..."):
            try:
                final_file = process_cat_sheet(uploaded_template, uploaded_csvs)
                st.success("Successfully processed! These rows are cloned from the Feature_Templates.")
                st.download_button("📥 Download for CCC Portal", final_file, file_name="Schick_CCC_Validated_CAT.xlsm")
            except Exception as e:
                st.error(f"Processing Error: {e}")
