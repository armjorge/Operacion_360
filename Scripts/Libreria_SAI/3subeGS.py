import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe
import glob
import os

# Define a function to load the most recent file with a certain prefix
def load_latest_excel(prefix, usecols):
    files = glob.glob(f"{prefix}*.xlsx")
    if not files:
        raise FileNotFoundError(f"No files found with prefix: {prefix}")
    latest_file = max(files, key=os.path.getctime)
    return pd.read_excel(latest_file, usecols=usecols)



# Load the data from Excel files
altas = load_latest_excel('altas_export_', 'A:N')
ordenes = load_latest_excel('ordenes_export_', 'A:O')
insabi = pd.read_excel('ordenesSuministro.xlsx', usecols='A:O')
zoho = pd.read_excel('zoho2024.xlsx', usecols='A:N')
#facturacionimss = pd.read_excel('C:/Users/armjorge/Dropbox/FACT 2024/Generacion facturas IMSS 2024.xlsx', sheet_name='Reporte Paq', usecols='A:L')




print("Exceles 2024 ubicados")

# Define the scope
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# Add your service account file
creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)

# Authorize the clientsheet 
client = gspread.authorize(creds)

# Get the instance of the Spreadsheet
spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1KN4XwXQlZ5jhyErxdpA_R6kNk2tKwwFeDnCDuHCr0fo/edit#gid=2033397596')

# Get the sheets by their names
sheet1 = spreadsheet.worksheet('RAW_ALTA2024')
sheet2 = spreadsheet.worksheet('RAW_IMSS2024')
sheet3 = spreadsheet.worksheet('RAW_INSABI2024')
sheet4 = spreadsheet.worksheet('ZOHO2024')
#sheet5 = spreadsheet.worksheet('PAQ_IMSS2024')


print("Hojas en el Gsheet ubicadas")

# Clear the contents of the sheets
sheet1.clear()
sheet2.clear()
sheet3.clear()
sheet4.clear()
print("Hojas en el Gsheet sin contenido")

# Appending the dataframes to the sheets
set_with_dataframe(sheet1, altas)
set_with_dataframe(sheet2, ordenes)
set_with_dataframe(sheet3, insabi)
set_with_dataframe(sheet4, zoho)

"""
# Step 1: Extract the Last 'Folio' Value from sheet5
existing_values = sheet5.col_values(2)  # Assuming 'Folio' is in column B, which is column index 2 in gspread
if existing_values:  # Ensure there's at least one value
    last_folio = existing_values[-1]  # Get the last 'Folio' value
else:
    last_folio = None

# Step 2: Identify New Rows in facturacionimss
if last_folio:
    # Find the row in facturacionimss that contains the last_folio
    folio_idx = facturacionimss[facturacionimss['Folio'] == last_folio].index.max()
    if pd.isna(folio_idx):  # If last_folio is not found, prepare to append the entire DataFrame
        new_rows_df = facturacionimss
    else:
        # Select rows that are after the last_folio row
        new_rows_df = facturacionimss.iloc[folio_idx + 1:]
else:
    # If sheet5 is empty, prepare to append the entire DataFrame
    new_rows_df = facturacionimss

# Step 3: Append the New Rows to sheet5
if not new_rows_df.empty:
    # Appending the new rows to the sheet starting from the row after the last existing row
    set_with_dataframe(sheet5, new_rows_df, include_index=False, include_column_header=False, row=len(existing_values)+1)
    print("\nNuevas filas agregadas a la consola de facturación IMSS 2024 \n")
else:
    print("\nNo hay nuevas filas para agregar a PAQ_IMSS2024\n")
"""
print("\n*******************************\nDatos en el google sheet 2024, recuerda correr scripts adicionales\n*******************************\n")