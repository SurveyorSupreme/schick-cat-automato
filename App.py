import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

def process_batch(uploaded_files, template_file):
    # Load template into memory
    template_bytes = template_file.read()
    wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    
    for file in uploaded_files:
        df = pd.read_csv(file, header=None)
        code = str(df.iloc[0, 0]).upper() # E, D, F, or G
        
        # Logic: Determine tab based on code or column count
        # Pipes/Lines usually have more columns in 12d exports
        if len(df.columns) > 15: 
            ws = wb["Line Asset Inputs"]
        else:
            ws = wb["Point Asset Inputs"]
            
        start_row = ws.max_row + 1
        # Fill the data
        for i, row in df.iterrows():
            for col_num, value in enumerate(row, start=1):
                cell_val = value if pd.notnull(value) else "LEAVE BLANK"
                ws.cell(row=start_row + i, column=col_num).value = cell_val

    # Save to a buffer to allow download
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

st.title("🚜 Schick Group: CCC CAT Batch Mapper")
st.write("Upload your 12d CSVs for Stormwater (E), Sewer (D), or Water (G).")

# Select files
csv_files = st.file_uploader("Upload 12d Exports", type="csv", accept_multiple_files=True)
template = st.file_uploader("Upload CCC Template (.xlsm)", type="xlsm")

if csv_files and template:
    if st.button("Process & Combine"):
        result = process_batch(csv_files, template)
        st.download_button("📥 Download Completed CAT Sheet", result, "Combined_CAT_Sheet.xlsm")
