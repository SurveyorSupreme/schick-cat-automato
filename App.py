import pandas as pd
import openpyxl

def process_cat_sheet(template_path, csv_files):
    # 1. Open the CCC Template safely so we don't destroy dropdowns
    wb = openpyxl.load_workbook(template_path)
    
    # Define our routing logic based on the CAT Feature Number (Column A)
    # You can add more codes here as you export them from 12d
    routing_map = {
        'G07': 'Point Asset Inputs',     # Fittings, Valves, etc.
        'G04': 'Line Asset Inputs',      # Pipes, Submains, etc.
        'G18': 'Polygon Asset Inputs'    # Thrust Blocks, etc.
    }

    for csv_file in csv_files:
        # Read the CSV without headers (since 12d exports raw data)
        df = pd.read_csv(csv_file, header=None)
        
        for index, row in df.iterrows():
            feature_code = str(row[0]).strip() # E.g., 'G04' or 'G07'
            
            # Find which sheet this belongs to
            target_sheet_name = routing_map.get(feature_code)
            
            if target_sheet_name and target_sheet_name in wb.sheetnames:
                sheet = wb[target_sheet_name]
                
                # Find the next empty row in this specific sheet
                # We check Column A to see if it has data
                next_row = 2 
                while sheet.cell(row=next_row, column=1).value is not None:
                    next_row += 1
                
                # Write the CSV data into the cells one by one
                # This PRESERVES the dropdowns and data validation!
                for col_index, value in enumerate(row):
                    # Only write if it's not a NaN/Blank from the CSV
                    if pd.notna(value):
                        # Excel columns are 1-indexed, so we add 1
                        sheet.cell(row=next_row, column=col_index + 1).value = value

    # Save the new file (don't overwrite the original template!)
    output_filename = "Processed_Schick_CAT_Sheet.xlsm"
    wb.save(output_filename)
    return output_filename
