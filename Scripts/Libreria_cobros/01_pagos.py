import pandas as pd
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe

print("\n*******************************\nExtracción de fecha de pagos\n*******************************\n")

# Define the file paths
source_2023 = 'C:/Users/armjorge/Dropbox/FACT 2023/Facturación y pagos 2023.xlsx'
source_2024 = 'C:/Users/armjorge/Dropbox/FACT 2024/Facturación y pagos 2024.vfinal.xlsx'

def read_excel_verify_headers(file_path, year):
    # Read the file without header to manually verify and set the column names
    df = pd.read_excel(file_path, header=None)
    
    # Update expected column names to include "Producto", "Piezas", and "Precio" for both years
    expected_names_2023 = ["Folio", "UUID", "Receptor", "Total", "Fecha Pago", "Importe pago", "Fecha sanción", "Sanción", "Fecha", "Producto", "Piezas", "Precio"]
    expected_names_2024 = ["Folio", "UUID", "Receptor", "Total", "Fecha Pago", "Importe pago", "Fecha sanción", "Sanción", "Fecha", "Producto", "Piezas", "Precio"]
    
    # Update the column indices to include indices for columns F, G, H (5, 6, 7 in zero-based indexing)
    columns_indices = [2, 3, 4, 10, 15, 16, 31, 32, 1, 5, 6, 7]  # Including columns F, G, H as "Producto", "Piezas", "Precio"
    
    # Select only the relevant columns based on updated indices
    df = df.iloc[:, columns_indices]
    
    # Set column names based on the year with the updated lists
    column_names = expected_names_2023 if year == 2023 else expected_names_2024
    
    # If it's the 2024 DataFrame, adjust the 'Fecha Pago' column name for consistency
    # Note: This step may not be necessary with the updated structure since "Fecha Pago" is already correctly positioned
    # But it's left here for clarity and future adjustments
    if year == 2024:
        column_names[column_names.index("Fecha Pago")] = "Fecha Pago"
    
    # Apply the column names to the DataFrame
    df.columns = column_names
    return df

# Read and verify the Excel files
df_2023 = read_excel_verify_headers(source_2023, 2023)
df_2024 = read_excel_verify_headers(source_2024, 2024)



if df_2023 is not None and df_2024 is not None:
    # Ensure both DataFrames have exactly the same column names
    assert df_2023.columns.equals(df_2024.columns), "DataFrames have mismatched columns names"
    df_2023 = df_2023.iloc[1:]  # Drop the first row of df_2023
    df_2024 = df_2024.iloc[1:]  # Drop the first row of df_2024
    # Merge the dataframes
    merged_df = pd.concat([df_2023, df_2024], ignore_index=True)
    # Assuming 'merged_df' is the DataFrame created after merging df_2023 and df_2024

    # 1. Delete rows where all fields are empty
    merged_df.dropna(how='all', inplace=True)
    # 1.1 Fechas de sanción a fechas de pago. 
    merged_df.loc[merged_df['Fecha Pago'].isna() & merged_df['Fecha sanción'].notna(), 'Fecha Pago'] = merged_df.loc[merged_df['Fecha Pago'].isna() & merged_df['Fecha sanción'].notna(), 'Fecha sanción']

    # 2. Delete rows where 'Fecha de pago' is blank, empty, or doesn't have data
    merged_df.dropna(subset=['Fecha Pago'], inplace=True)

    # 4. Adjust 'Folio' to match "P-" followed by numbers, and print modified values
    # Regex pattern for 'Folio' validation and extraction
    folio_pattern = re.compile(r'(P-\d+)')
    # Find and adjust non-conforming 'Folio' values
    def adjust_folio(folio):
        match = folio_pattern.search(folio)
        return match.group(0) if match else folio  # Return the original folio if no match found

    # Apply the adjustment and capture changes
    original_folios = merged_df['Folio'].astype(str)
    merged_df['Folio'] = original_folios.apply(adjust_folio)
    # Find modified values by comparing the new 'Folio' column to the original values
    modified_folios = merged_df[merged_df['Folio'] != original_folios]
    if not modified_folios.empty:
        print("Modified 'Folio' values:")
        print(modified_folios[['Folio']])

    # Write the cleaned DataFrame to an Excel file
    merged_df.to_excel(r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Dataframes\2023-2024pagos.xlsx", index=False)
    print("Archivo con pagos registrados.xlsx")
else:
    print("Issue with reading files, please check the output for more details.")
# Define the scope
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
# Add your service account file
creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)  # Ensure the correct path
# Authorize the client sheet
client = gspread.authorize(creds)
# Get the instance of the Spreadsheet
spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1KN4XwXQlZ5jhyErxdpA_R6kNk2tKwwFeDnCDuHCr0fo/edit#gid=2033397596')
# Access the specific worksheet 'PagosKM'
gsheet_pagos = spreadsheet.worksheet('PagosKM')
# Clear existing data
gsheet_pagos.clear()
# Upload the DataFrame to the worksheet
set_with_dataframe(gsheet_pagos, merged_df)  # Ensure 'merged_df' is your DataFrame variable
print("\n*******************************\nPagos actualizados en el google sheet\n*******************************\n")

########## A PARTIR DE AQUÍ SALE LA LISTA DE FACTURAS CON Y SIN PAGOS #############
exclusion_list = [
    "Diagnoquim, S.A. de C.V.", "Nadro SAPI de CV", "Clínica de Corta Estancia Jatzibe SA de CV", "Vision & Links, S.A. de C.V.", "Servicios Médicos de Corta Estancia SC", "Administradora De Hospitales Grande De Oriente SRL de CV", "Ziba Medicina de Alta Especialidad SC", "Hospital Medex Roma SA de CV", "Cima Cirugía Mayor Ambulatoria Roma SAS de CV", "Artimédica SA de CV", "DLV Hospital SAPI de CV", "Distribuidora Lactymedic SA De CV", "Insumos Médicos Unión SRL de CV", "S.M.CH. Medicina y Hospitalización Chavez SC", "Fers-Med SA de CV", "Adra de México SA de CV", "Folio perdido", "Norma Araceli Pérez Martínez", "Deilou Andreina García Salas", "Clínica Quirúrgica Roma SRL de CV", "Medirep SA de CV", "Juan Antonio Alvarado Guevara", "Yarenhy Monroy Hernandez", "Ernesto Sanchez Cid", "Clínica Santa Lucía Toluca SA de CV", "Nayeli Isabel Trejo Bahena  ", "Ventas Mostrador", "Jose Antonio Jimenez Contreras", "Jose Daniel Lozada Leon", "Hospital Churubusco SA de CV", "Jesús Adrian Orozco Carnalla", "Sanatorios Unidos SA ", "Yesenia Estevez Estevez", "Andres Muciño Cortes", "Hospital Policlinica Lucano", "Miguel Ángel Martínez Aguilar", "Central Médica Paola", "Maritere Gonzalez Lama", "Endomedix", "Zerifar", "Ana Elena Gutiérrez Peredo", "Bernardo de Jesús Salgado Ruíz", "Vet Point", "Santek Health", "Gabriela Islas Lagunas", "Luis Felipe Cuellar Guzman", "Lilia Ivonne Pineda Castañeda", "Servicios Veterinarios Darwin", "Coloproctologia Diagnostica", "Cancelada", "Mauricio Fernando Gonzalez Costes", "Aldo Alain Diaz Garduño", "Clínica Veterinaria Casaubon", "Super Megapet", "Quálitas Compañía de Seguros", "Javier Guevara Cardoso", "GNK Logística", "Representaciones OPV", "Corporación Armo", "Waldra Medical", "Fers-Med", "FARMACEUTICA MEDIKAMENTA", "CLINICA QUIRURGICA ROMA", "C&M DISTRIBUIDORA DE MEDICAMENTOS Y MATERIAL DE CURACION", "NADRO SAPI DE CV","ZERIFAR","FERS-MED","ANDRES MUCIÑO CORTES","REPRESENTACIONES OPV","Distribuidora Lactymedic SA de CV", "CANCELADA"
]

# Filter out the rows based on 'Receptor' column
df_2023_conysinpago = df_2023[~df_2023['Receptor'].isin(exclusion_list)]
df_2024_conysinpago = df_2024[~df_2024['Receptor'].isin(exclusion_list)]

if df_2023_conysinpago is not None and df_2024_conysinpago is not None:
    # Ensure both DataFrames have exactly the same column names
    assert df_2023.columns.equals(df_2024.columns), "DataFrames have mismatched columns names"
    
    # Drop the first row of both DataFrames
    df_2023_conysinpago = df_2023_conysinpago.iloc[1:]
    df_2024_conysinpago = df_2024_conysinpago.iloc[1:]
    
    # Merge the dataframes
    merged_conysinpago_df = pd.concat([df_2024_conysinpago, df_2023_conysinpago], ignore_index=True)
    
    # Delete rows where all fields are empty
    merged_conysinpago_df.dropna(how='all', inplace=True)
    
    # Sort the DataFrame to prioritize rows with data in "Fecha Pago", "Importe pago", "Fecha sanción"
    # Rows with any of these fields filled are moved to the top
    merged_conysinpago_df['sort_priority'] = merged_conysinpago_df[['Fecha Pago', 'Importe pago', 'Fecha sanción']].notna().any(axis=1).astype(int)
    merged_conysinpago_df.sort_values(by=['sort_priority', 'Folio', 'UUID', 'Piezas', 'Precio'], ascending=[False, True, True, True, True], inplace=True)
    merged_conysinpago_df.drop('sort_priority', axis=1, inplace=True)
    
    # Remove duplicates based on 'Folio', 'UUID', 'Piezas', and 'Precio', keeping the first (filled data prioritized)
    merged_conysinpago_df = merged_conysinpago_df.drop_duplicates(subset=['Folio', 'UUID', 'Piezas', 'Precio'])
    
    # Update 'Fecha Pago' from 'Fecha sanción' where applicable
    merged_conysinpago_df.loc[merged_conysinpago_df['Fecha Pago'].isna() & merged_conysinpago_df['Fecha sanción'].notna(), 'Fecha Pago'] = merged_conysinpago_df.loc[merged_conysinpago_df['Fecha Pago'].isna() & merged_conysinpago_df['Fecha sanción'].notna(), 'Fecha sanción']
    
    # Adjust 'Folio' to match "P-" followed by numbers, and print modified values
    folio_pattern = re.compile(r'(P-\d+)')
    def adjust_folio(folio):
        match = folio_pattern.search(folio)
        return match.group(0) if match else folio
    
    original_folios = merged_conysinpago_df['Folio'].astype(str)
    merged_conysinpago_df['Folio'] = original_folios.apply(adjust_folio)
    modified_folios = merged_conysinpago_df[merged_conysinpago_df['Folio'] != original_folios]
    
    if not modified_folios.empty:
        print("Modified 'Folio' values:")
        print(modified_folios[['Folio']])
    
    # Write the cleaned DataFrame to an Excel file
    merged_conysinpago_df.to_excel(r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Dataframes\2023-2024conysinpagos.xlsx", index=False)
    print("Process completed. The cleaned data is written to 2023-2024 Con y Sin pagos.xlsx")
else:
    print("Issue with reading files, please check the output for more details.")