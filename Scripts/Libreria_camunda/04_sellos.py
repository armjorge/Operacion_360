import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe
import glob
import os
from openpyxl import load_workbook

def convert_dates(date_str):
    # Check and fix the hour part if it's missing a leading zero (for "2024-04-01 0:00:00" format)
    if '-' in date_str and ':' in date_str and date_str[-5] == ' ':
        date_str = date_str[:11] + '0' + date_str[11:]
    
    # Try converting using the default format (yyyy-mm-dd HH:MM:SS)
    try:
        return pd.to_datetime(date_str, format='%Y-%m-%d %H:%M:%S', errors='raise')
    except ValueError:
        # If the default conversion fails, try the day-first format (dd/mm/yyyy)
        try:
            return pd.to_datetime(date_str, dayfirst=True, errors='raise')
        except ValueError:
            # Return the original string if all conversions fail (shouldn't happen if all formats are accounted for)
            return date_str

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# Add your service account file
creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)

# Authorize the clientsheet 
client = gspread.authorize(creds)

# Get the instance of the Spreadsheet
spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1KN4XwXQlZ5jhyErxdpA_R6kNk2tKwwFeDnCDuHCr0fo/edit#gid=2033397596')

# Get the sheets by their names
worksheet = spreadsheet.worksheet('TYCSA_logistica')
worksheet_RAWINSABI = spreadsheet.worksheet('RAW_INSABI')
# Convert the worksheet data to a pandas DataFrame
df_datoslogisticos = pd.DataFrame(worksheet.get_all_records())
print(df_datoslogisticos.columns)
df_ordenes = pd.DataFrame(worksheet_RAWINSABI.get_all_records())

# Load df_sellos from Excel file
df_sellos = pd.read_excel("./INSABI_ordenes-contratos-sellos-remisiones.xlsx", sheet_name="Sellos")

# Step 2: Load df_sellos from the Excel file
df_ordenes_list = df_ordenes['NÚMERO DE ORDEN DE SUMINISTRO'].tolist()
df_filtered = df_datoslogisticos[df_datoslogisticos['INSTITUTO'].isin(['INSABI', 'IMSS BIENESTAR'])]
dropped_pedidos = df_filtered[~df_filtered['PEDIDO CLIENTE'].isin(df_ordenes_list)]['PEDIDO CLIENTE'].tolist()
df_filtered = df_filtered[df_filtered['PEDIDO CLIENTE'].isin(df_ordenes_list)]
df_filtered['FECHA DE ENTREGA AL CLIENTE'] = df_filtered['FECHA DE ENTREGA AL CLIENTE'].apply(lambda x: convert_dates(x))
print("Dropped PEDIDO CLIENTE values:", dropped_pedidos)

# Step 3: Update df_sellos with non-existing PEDIDO CLIENTE values
for index, row in df_filtered.iterrows():
    if row['PEDIDO CLIENTE'] not in df_sellos['Suministro'].values:
        new_row = {'Suministro': row['PEDIDO CLIENTE'], 'Fecha de sello': row['FECHA DE ENTREGA AL CLIENTE']}
        df_sellos = pd.concat([df_sellos, pd.DataFrame([new_row])], ignore_index=True)

# Step 4: Write back to the Excel file without removing other sheets
with pd.ExcelWriter("./INSABI_ordenes-contratos-sellos-remisiones.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    book = writer.book
    try:
        book.remove(book['Sellos'])
    except KeyError:
        pass  # The sheet does not exist, which is fine
    df_sellos.to_excel(writer, sheet_name='Sellos', index=False)

print("\n*******************************\n INSABI_ordenes-contratos-sellos-remisiones.xlsx hoja Sellos actualizada \n*******************************\n")