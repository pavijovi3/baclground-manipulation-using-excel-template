import os
import pandas as pd
import openpyxl
import xlwings as xw

# Define the directory paths
raw_dir = r"H:\FTIR_data\DOCUMENTS\background-manupulation-pyprogram\New folder\excel_raw"
ready_dir = r"H:\FTIR_data\DOCUMENTS\background-manupulation-pyprogram\New folder\excel_format"
output_dir = r"H:\FTIR_data\DOCUMENTS\background-manupulation-pyprogram\New folder\excel_output"

# Change the working directory to the raw directory
os.chdir(raw_dir)

# Get a list of all the Excel files in the raw directory (excluding temporary files)
files = [f for f in os.listdir(raw_dir) if f.endswith('.xlsx') and not f.startswith('~$')]

for filename in files:
    # Read the Excel file into a pandas dataframe
    df = pd.read_excel(os.path.join(raw_dir, filename))

    # Load workbook and select worksheet
    app = xw.App(visible=False)
    wb = xw.Book(os.path.join(ready_dir, 'Col_D_0.3V.xlsx'))
    ws = wb.sheets['Sheet3']

    # Update workbook at specified range
    ws.range('A1').options(index=False).value = df

    # Define the output file path and name
    output_path = os.path.join(output_dir, f"{os.path.splitext(filename)[0]}_0.3V_output.xlsx")

    # Save and close the workbook with the output file name
    wb.save(output_path)
    wb.close()
    app.quit()