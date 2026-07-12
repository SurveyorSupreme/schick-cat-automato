import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy
from io import BytesIO
import re

# --- 1. THE CORE ENGINE ---
def process_cat_sheet(template_file, csv_files, is_xlsm):
    # keep_vba only matters for macro-enabled (CCC) templates; forcing it on a
    # plain .xlsx (e.g. the IDS Reticulation CAT template) mislabels the output.
    wb = openpyxl.load_workbook(template_file, keep_vba=is_xlsm)

    ft_sheet = wb['Feature_Templates']
    routing_map = {}
    master_row_map = {}  # Stores the row index for each code
    current_cat = None

    # 1. Build a map of all codes and their "Golden Row" index
    for r in range(1, ft_sheet.max_row + 1):
        val = str(ft_sheet.cell(row=r, column=1).value or "").strip()
        if "Point" in val: current_cat = "Point Asset Inputs"
        elif "Line" in val: current_cat = "Line Asset Inputs"
        elif "Polygon" in val: current_cat = "Polygon Asset Inputs"

        if re.match(r'^[A-Z]\d{2}$', val):
            routing_map[val] = current_cat
            master_row_map[val] = r

    # 2. Index every dropdown rule on Feature_Templates ONCE, by (row, col).
    # The old per-CSV-row scan of all ~766 rules was slow and produced one
    # DataValidation object per cell, which bloats the file until Excel
    # "repairs" it — and a repair strips every dropdown in the workbook.
    dv_by_cell = {}
    for dv in ft_sheet.data_validations.dataValidation:
        for rng in dv.sqref.ranges:
            for row, col in rng.cells:
                dv_by_cell[(row, col)] = dv

    # One shared DataValidation per unique rule per target sheet; we append
    # cells to it instead of creating a new rule for every single cell.
    shared_dvs = {}

    def target_dv_for(sheet, dv):
        key = (sheet.title, dv.type, dv.operator, dv.formula1, dv.formula2,
               dv.allow_blank, dv.showDropDown)
        if key not in shared_dvs:
            new_dv = DataValidation(
                type=dv.type, operator=dv.operator,
                formula1=dv.formula1, formula2=dv.formula2,
                allow_blank=dv.allow_blank,
                showErrorMessage=dv.showErrorMessage,
                showInputMessage=dv.showInputMessage,
                showDropDown=dv.showDropDown,
                error=dv.error, errorTitle=dv.errorTitle,
                prompt=dv.prompt, promptTitle=dv.promptTitle,
            )
            sheet.add_data_validation(new_dv)
            shared_dvs[key] = new_dv
        return shared_dvs[key]

    # Track the next empty row per sheet instead of rescanning from row 2
    # for every CSV row.
    next_row_map = {}

    def next_empty_row(sheet):
        if sheet.title not in next_row_map:
            r = 2
            while sheet.cell(row=r, column=1).value is not None:
                r += 1
            next_row_map[sheet.title] = r
        return next_row_map[sheet.title]

    # 3. Process the CSVs
    for csv_file in csv_files:
        df = pd.read_csv(csv_file, header=None)

        for _, csv_row in df.iterrows():
            code = str(csv_row[0]).strip()
            target_sheet_name = routing_map.get(code)

            if target_sheet_name and target_sheet_name in wb.sheetnames:
                target_sheet = wb[target_sheet_name]
                source_row_idx = master_row_map[code]
                next_row = next_empty_row(target_sheet)

                # --- STEP A: CLONE THE ROW (Styles & Values) ---
                for col in range(1, ft_sheet.max_column + 1):
                    source_cell = ft_sheet.cell(row=source_row_idx, column=col)
                    target_cell = target_sheet.cell(row=next_row, column=col)

                    target_cell.value = source_cell.value
                    if source_cell.has_style:
                        target_cell.font = copy(source_cell.font)
                        target_cell.border = copy(source_cell.border)
                        target_cell.fill = copy(source_cell.fill)
                        target_cell.number_format = source_cell.number_format
                        target_cell.alignment = copy(source_cell.alignment)

                    # --- STEP B: MIGRATE DROPDOWNS (Data Validation) ---
                    dv = dv_by_cell.get((source_row_idx, col))
                    if dv is not None:
                        target_dv_for(target_sheet, dv).add(target_cell.coordinate)

                # --- STEP C: OVERWRITE WITH CSV VALUES ---
                # This ensures we keep the dropdowns we just pasted but put in the real data
                for col_idx, value in enumerate(csv_row):
                    if pd.notna(value):
                        target_sheet.cell(row=next_row, column=col_idx + 1).value = value

                next_row_map[target_sheet.title] = next_row + 1

    # Save to buffer
    out_buffer = BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- 2. STREAMLIT UI ---
st.set_page_config(page_title="Schick CAT | Deep Clone", page_icon="🚜")
st.title("🚜 Schick Group: Master Row Cloner")
st.markdown("Clones **Feature_Templates** rows (including dropdowns) into the Input sheets.")

template = st.file_uploader("1. Upload CAT Template (.xlsx or .xlsm)", type=['xlsx', 'xlsm'])
csvs = st.file_uploader("2. Upload 12d CSVs", type=['csv'], accept_multiple_files=True)

if st.button("Generate Validated CAT"):
    if template and csvs:
        with st.spinner("Deep cloning council rules..."):
            try:
                is_xlsm = template.name.lower().endswith('.xlsm')
                output = process_cat_sheet(template, csvs, is_xlsm)
                out_name = "Schick_Validated_CAT" + (".xlsm" if is_xlsm else ".xlsx")
                st.success("Done! Your survey data is now inside official Council rows.")
                st.download_button("📥 Download for Council Portal", output, file_name=out_name)
            except Exception as e:
                st.error(f"Error: {e}")
