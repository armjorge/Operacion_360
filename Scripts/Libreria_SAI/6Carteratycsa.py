import pandas as pd
import numpy as np

# File paths
source_file_path = 'C:/Users/armjorge/Dropbox/FACT 2023/Generacion facturas IMSS VFinal.xlsx'
target_file_path = './M02 Cartera template.xlsx'  # Assuming this is now an unformatted Excel file
imss_2024_path = 'C:/Users/armjorge/Dropbox/FACT 2024/Generacion facturas IMSS 2024.xlsx'
df_2024 = pd.read_excel(imss_2024_path, sheet_name='Reporte Paq', usecols=['Folio', 'Fecha', 'Alta', 'Referencia', 'UUID','Total'])
df_source = pd.read_excel(source_file_path, sheet_name='Reporte Paq', usecols=['Folio', 'Fecha', 'Alta', 'Referencia', 'UUID','Total'])

# Filter out rows where "Alta" is empty, #N/A, or only spaces
df_filtered = df_source.dropna(subset=['Alta'])
df_filtered = df_filtered[~df_filtered['Alta'].astype(str).str.strip().isin(['', '#N/A'])]

# Assuming DETALLE sheet already exists and has headers but we will overwrite it with new data
# If DETALLE does not exist, this will create it
df_target = pd.read_excel(target_file_path, sheet_name='DETALLE')

# Ensure the target DataFrame has the correct structure to receive the data
# This step clears existing data below headers and prepares for new data
df_target = pd.DataFrame(columns=df_target.columns)

# Map source columns to target columns based on your specification
df_target['Factura'] = df_filtered['Folio']
df_target['Fecha factura'] = df_filtered['Fecha']
df_target['Alta'] = df_filtered['Alta']
df_target['base'] = df_filtered['Referencia']
df_target['UUID'] = df_filtered['UUID']
df_target['Importe'] = df_filtered['Total']
df_target['Nombre cliente'] = 'IMSS Bienestar ESEOTRES'


# Filter out rows where "Alta" is empty, #N/A, or only spaces
df_2024_filtered = df_2024.dropna(subset=['Alta'])
df_2024_filtered = df_2024_filtered[~df_2024_filtered['Alta'].astype(str).str.strip().isin(['', '#N/A', np.nan])]

# Create a new DataFrame for appending, matching the structure of df_target
df_new_rows = pd.DataFrame({
    "Factura": df_2024_filtered['Folio'],
    "Fecha factura": df_2024_filtered['Fecha'],
    "Alta": df_2024_filtered['Alta'],
    "base": df_2024_filtered['Referencia'],
    "UUID": df_2024_filtered['UUID'],
    "Importe": df_filtered['Total'],
    "Nombre cliente": "IMSS Bienestar ESEOTRES"  # Filling this for all rows
})

# Use pd.concat to append new rows to df_target
df_target = pd.concat([df_target, df_new_rows], ignore_index=True)

# Read "Órdenes" sheet from both files for columns A:O
df_orders_2023 = pd.read_excel(source_file_path, sheet_name='Órdenes', usecols='A:O')
df_orders_2024 = pd.read_excel(imss_2024_path, sheet_name='Órdenes', usecols='A:O')

# Merge both DataFrames into df_SAI2324
df_SAI2324 = pd.concat([df_orders_2023, df_orders_2024], ignore_index=True)

# Print the first 10 and the last 10 rows to confirm
print("First 10 rows:")
print(df_SAI2324.head(10))
print("\nLast 10 rows:")
print(df_SAI2324.tail(10))

# Add new columns to df_target
df_target["Unidad operativa"] = None
df_target["Ejercicio fiscal"] = None
df_target["Contrato"] = None
df_target["Nombre cliente"] = ["IMSS ESEOTRES" if alta else None for alta in df_target["Alta"].notnull()]

# Populate new columns based on values found in df_SAI2324
for index, row in df_target.iterrows():
    matching_rows = df_SAI2324[df_SAI2324["orden"] == row["base"]]
    if not matching_rows.empty:
        df_target.at[index, "Unidad operativa"] = matching_rows.iloc[0]["descripciónEntrega"]
        df_target.at[index, "Ejercicio fiscal"] = matching_rows.iloc[0]["fechaExpedicion"]  
        df_target.at[index, "Contrato"] = matching_rows.iloc[0]["contrato"]

# Keep only the last 4 digits from each value in "Ejercicio fiscal"
df_target["Ejercicio fiscal"] = df_target["Ejercicio fiscal"].astype(str).str[-4:]

# Filter out rows where "base" is empty or blank in df_target
df_target_filtered = df_target.dropna(subset=['base'])  # Remove rows where "base" is NaN
df_target_filtered = df_target_filtered[df_target_filtered['base'].astype(str).str.strip() != '']  # Remove rows with blank "base"

# Define the new Excel file path
imssb2024_path = 'C:/Users/armjorge/Dropbox/FACT 2024/Consola IMSSB.xlsx'

# Read the specified sheets from IMSSB2024
df_imssb2024_paq = pd.read_excel(imssb2024_path, sheet_name='Reporte PAQ', usecols='A:L')
df_imssb2024_ordenes = pd.read_excel(imssb2024_path, sheet_name='Órdenes', usecols='A:O')
df_imssb2024_contratos = pd.read_excel(imssb2024_path, sheet_name='Contratos', usecols='A:B')

# Initialize df_IMSSBIENESTAR2024 with df_imssb2024_paq's data
df_IMSSBIENESTAR2024 = df_imssb2024_paq.copy()

# Add new columns with default/empty values
df_IMSSBIENESTAR2024['Unidad operativa'] = ''
df_IMSSBIENESTAR2024['Ejercicio fiscal'] = ''
df_IMSSBIENESTAR2024['Contrato'] = ''
df_IMSSBIENESTAR2024['Nombre cliente'] = 'IMSS Bienestar ESEOTRES'  # Filling "Nombre cliente" for all rows

# Populate the new columns based on lookups
for index, row in df_IMSSBIENESTAR2024.iterrows():
    referencia = row['Referencia']
    
    # Lookup in Órdenes sheet for Unidad operativa
    match_ordenes = df_imssb2024_ordenes[df_imssb2024_ordenes['NÚMERO DE ORDEN DE SUMINISTRO'] == referencia]
    if not match_ordenes.empty:
        df_IMSSBIENESTAR2024.at[index, 'Unidad operativa'] = match_ordenes['CLUES'].iloc[0]
        
        # Extracting year from "FECHA EXPEDICIÓN DE LA ORDEN" and keeping last 4 digits
        fecha_expedicion = str(match_ordenes['FECHA EXPEDICIÓN DE LA ORDEN'].iloc[0])
        df_IMSSBIENESTAR2024.at[index, 'Ejercicio fiscal'] = fecha_expedicion[-4:]
    
    # Lookup in Contratos sheet for Contrato
    match_contratos = df_imssb2024_contratos[df_imssb2024_contratos['Orden'] == referencia]
    if not match_contratos.empty:
        df_IMSSBIENESTAR2024.at[index, 'Contrato'] = match_contratos['Contrato'].iloc[0]

# Print the first 15 rows of df_IMSSBIENESTAR2024 to verify
print(df_IMSSBIENESTAR2024.head(15))

# Assuming df_2024_filtered is already defined and ready for updates
print(df_target_filtered.columns)
print(df_IMSSBIENESTAR2024.columns)
