import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
from io import BytesIO
import re

# --- 1. THE CLONING ENGINE ---
def clone_row_with_rules(source_sheet, source_row_idx, target_sheet, target_row_idx):
    """
    Mimics a 'Copy Row' -> 'Paste Row' action, including styles and values.
    """
    for col_idx in range(1, source_sheet.max_column + 1):
        source_cell = source_sheet.cell(row=source_row_idx, column=col_idx)
        target_cell = target_sheet.cell(row=target_row_idx, column=col_idx)
        
        # Copy Value & Style
        target_cell.value = source_cell.value
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)

def process_cat_sheet(template_file, csv_files):
    # Load with keep_vba=True to preserve any Excel macros
    wb = openpyxl.load_workbook(template_file, keep_vba=True)
    
    # 1. Map every CAT Code to its 'Golden Row' in Feature_Templates
    ft_sheet = wb['Feature_Templates']
    master_rows = {} # Code -> Row Index
    routing_map = {}
    current_cat = None
    
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
                
                # Find the next empty row in the input sheet
                next_row = 2
                while target_sheet.cell(row=next_row, column=1).value is not None:
                    next_row += 1
                
                # STEP A: CLONE the row from Feature_Templates (Copy/Paste mimic)
                if code in master_rows:
                    clone_row_with_rules(ft_sheet, master_rows[code], target_sheet, next_row)
                
                # STEP B: OVERWRITE placeholders with actual 12d survey data
                for col_idx, value in enumerate(csv_row):
                    if pd.notna(value):
                        target_sheet.cell(row=next_row, column=col_idx + 1).value = value

    # Save to memory
    out_buffer = BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- 2. INTERFACE ---
st.set_page_config(page_title="Schick Group | Master Cloner", page_icon="🚜")
st.title("🚜 Schick Group: Master CAT Cloner")
st.markdown("This version **clones the Council's Master Rows** from the Feature Templates to preserve all dropdowns and formatting.")

uploaded_template = st.file_uploader("1. Upload CCC Template (.xlsm)", type=['xlsm'])
uploaded_csvs = st.file_uploader("2. Upload 12d Exports", type=['csv'], accept_multiple_files=True)

if st.button("Generate Master CAT Sheet"):
    if uploaded_template and uploaded_csvs:
        with st.spinner("Cloning master rows and injecting survey data..."):
            try:
                final_file = process_cat_sheet(uploaded_template, uploaded_csvs)
                st.success("Done! The master rules have been copied into your new rows.")
                st.download_button("📥 Download Result", final_file, file_name="Schick_Master_CAT.xlsm")
            except Exception as e:
                st.error(f"Error: {e}")
