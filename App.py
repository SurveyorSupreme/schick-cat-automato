import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re

# --- 1. DYNAMIC ROUTING LOGIC ---
def build_routing_map(wb):
    """
    Scans the 'Feature_Templates' sheet to find all CAT codes
    and maps them to Point, Line, or Polygon sheets.
    """
    routing_map = {}
    if 'Feature_Templates' not in wb.sheetnames:
        return routing_map

    sheet = wb['Feature_Templates']
    current_category = None
    
    # Iterate through Column A to find headers and codes
    for row in range(1, sheet.max_row + 1):
        cell_value = str(sheet.cell(row=row, column=1).value or "").strip()
        
        # Identify the section type
        if "Type of Point feature" in cell_value:
            current_category = "Point Asset Inputs"
        elif "Type of Line feature" in cell_value:
            current_category = "Line Asset Inputs"
        elif "Type of Polygon feature" in cell_value:
            current_category = "Polygon Asset Inputs"
        
        # Check if the cell is a CAT code (e.g., G04, D12, E01)
        if re.match(r'^[A-Z]\d{2}$', cell_value) and current_category:
            routing_map[cell_value] = current_category
            
    return routing_map

def process_cat_sheet(template_file, csv_files):
    # Open the CCC Template (keep_vba=True preserves macros)
    wb = openpyxl.load_workbook(template_file, keep_vba=True)
    
    # Dynamically build the map (D, E, F, G, H, etc.)
    routing_map = build_routing_map(wb)
    
    if not routing_map:
        st.error("Could not find any feature codes in the 'Feature_Templates' sheet.")
        return None

    for csv_file in csv_files:
        df = pd.read_csv(csv_file, header=None)
        
        for index, row in df.iterrows():
            feature_code = str(row[0]).strip()
            target_sheet_name = routing_map.get(feature_code)
            
            if target_sheet_name and target_sheet_name in wb.sheetnames:
                sheet = wb[target_sheet_name]
                
                # Find next empty row
                next_row = 2 
                while sheet.cell(row=next_row, column=1).value is not None:
                    next_row += 1
                
                # Inject data cell-by-cell to preserve dropdowns
                for col_index, value in enumerate(row):
                    if pd.notna(value):
                        sheet.cell(row=next_row, column=col_index + 1).value = value

    out_buffer = BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- 2. STREAMLIT INTERFACE ---
st.set_page_config(page_title="Schick CAT Automator Pro", page_icon="🚜")

st.title("🚜 Schick Group: Universal CAT Automator")
st.markdown("""
This version automatically detects **all council codes** (Water, Storm, Waste, etc.) 
by scanning the template's Feature Templates.
""")

st.divider()

uploaded_template = st.file_uploader("1. Upload CCC CAT Template (.xlsm)", type=['xlsm', 'xlsx'])
uploaded_csvs = st.file_uploader("2. Upload 12d CSV Exports", type=['csv'], accept_multiple_files=True)

if st.button("Generate Processed CAT Sheet"):
    if uploaded_template and uploaded_csvs:
        with st.spinner("Scanning template and processing assets..."):
            try:
                final_file = process_cat_sheet(uploaded_template, uploaded_csvs)
                if final_file:
                    st.success("Success! All assets (G, E, D, F, H) have been routed.")
                    st.download_button(
                        label="📥 Download Processed CAT Sheet",
                        data=final_file,
                        file_name="Processed_Schick_CAT_Sheet.xlsm",
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12"
                    )
            except Exception as e:
                st.error(f"An error occurred: {e}")
    else:
        st.warning("Please upload both the template and your CSVs.")
