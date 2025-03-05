import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe
import os

os.startfile('.')

# Define the scope
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
# Add your service account file
creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)  # Ensure the correct path
# Authorize the client sheet
client = gspread.authorize(creds)
spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1KN4XwXQlZ5jhyErxdpA_R6kNk2tKwwFeDnCDuHCr0fo/edit#gid=2033397596')

def worksheet_to_df(worksheet_name, columns_range):
    worksheet = spreadsheet.worksheet(worksheet_name)
    data = worksheet.get(columns_range)
    print("INSABI cargado desde el google sheet")
    return pd.DataFrame(data[1:], columns=data[0])


# Ask the user for the file names
file_2023_2024 = input("File 2023-2024.xlsx: ")
file_2023 = input("File 2024.xlsx: ")

# Load the Excel files into DataFrames
df_2023_2024 = pd.read_excel(file_2023_2024)
df_2023 = pd.read_excel(file_2023)
df_INSABI_gsheet = worksheet_to_df('RAW_INSABI', 'A:AA')

# Merge the DataFrames
df_merged = pd.concat([df_2023_2024, df_2023], ignore_index=True)

# Drop rows where 'Estado de la factura' is 'Cancelada'
df_merged = df_merged[df_merged['Estado de la factura'] != 'Cancelado']

"""
Aquí tenemos que cargar la de INSABI del googlesheet y mapear las columnas
"""
# Rename the 'NÚMERO DE ORDEN DE SUMINISTRO' column in df_INSABI_gsheet to match df_merged for merging
# Add 'P-' prefix to 'Número de factura' and rename it to 'Factura'
df_merged['Factura'] = 'P-' + df_merged['Número de factura'].astype(str)

# Create a dictionary for fast lookup of Factura to Orden de suministro and vice versa
factura_to_orden = dict(zip(df_INSABI_gsheet['Factura'], df_INSABI_gsheet['NÚMERO DE ORDEN DE SUMINISTRO']))
orden_to_factura = dict(zip(df_INSABI_gsheet['NÚMERO DE ORDEN DE SUMINISTRO'], df_INSABI_gsheet['Factura']))

# Completa los falores faltantes desde el SAGI
def fill_missing_values(row):
    if pd.isna(row['Orden de suministro']) and not pd.isna(row['Factura']):
        row['Orden de suministro'] = factura_to_orden.get(row['Factura'])
    elif pd.isna(row['Factura']) and not pd.isna(row['Orden de suministro']):
        row['Factura'] = orden_to_factura.get(row['Orden de suministro'])
    return row

# Apply the function to the dataframe
df_merged = df_merged.apply(fill_missing_values, axis=1)

df_merged = df_merged.merge(df_INSABI_gsheet[['Factura','Contrato', 'Fuente de financiamiento', 'Fecha de entrega', 'Fecha de pago', 'Oficio de Sanción', 'sanción']], 
                            on='Factura', 
                            how='left')
df_merged.drop(columns=['Unnamed: 0', 'Número de oficio', 'Número de factura', 'Opciones'], inplace=True)

# Convert 'Total' column to integer after removing the currency sign
df_merged['Total'] = df_merged['Total'].replace('[\$,]', '', regex=True).astype(float).astype(int)

###########################################################################
# Ask the user for the prefix for the processed file
print(f"\n*******************************\n {file_2023_2024}\n{df_2023}\n*******************************\n")
mm_dd_preffix = input("Prefix processed: ")
print(f"\n*******************************\n{mm_dd_preffix}\n{file_2023_2024}\n{df_2023} \n*******************************\n")
# Create the output file name
output_filename = f"{mm_dd_preffix} 2023-2024 ESTATUS SAGI.xlsx"

# Save the merged DataFrame to the output Excel file
df_merged.to_excel(output_filename, index=False)

print(f"File saved as {output_filename}\n")

print("\n*******************************\nSe procede a subirlo en GoogleSheet\n*******************************\n")

# Access the specific worksheet 'PagosKM'
gsheet_estatusSAGI = spreadsheet.worksheet('EstatusSAGI')
# Clear existing data
gsheet_estatusSAGI.clear()
# Upload the DataFrame to the worksheet
set_with_dataframe(gsheet_estatusSAGI, df_merged)  # Ensure 'merged_df' is your DataFrame variable
print("\n*******************************\nEstatus SAGI en actualizado en el Google Sheet\n*******************************\n")