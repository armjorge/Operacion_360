import os
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe

# This function remains unchanged
def worksheet_to_df(worksheet_name, columns_range):
    worksheet = spreadsheet.worksheet(worksheet_name)
    data = worksheet.get(columns_range)
    print("INSABI cargado desde el google sheet")
    return pd.DataFrame(data[1:], columns=data[0])

def complete_missing_values(df_merged, df_insabi):
    """
    For each row in df_merged, if any of the fields
    'Número de contrato', 'Orden de suministro', or 'Número de factura'
    is missing, attempt to find a matching row in df_insabi using one of three pivot checks:
    
       1. df_merged['Folio fiscal'] == df_insabi['UUID']
       2. df_merged['Número de factura'] == df_insabi['Factura']
       3. df_merged['Orden de suministro'] == df_insabi['NÚMERO DE ORDEN DE SUMINISTRO']
    
    If a match is found, update the missing values as:
       - 'Orden de suministro' ← df_insabi['NÚMERO DE ORDEN DE SUMINISTRO']
       - 'Folio fiscal' ← df_insabi['UUID']
       - 'Número de factura' ← df_insabi['Factura']
       - 'Número de contrato' ← df_insabi['Número de contrato']
    
    Existing (non-missing) values are not overwritten.
    """
    # Iterate over each row in df_merged
    for idx, row in df_merged.iterrows():
        # Check if any target column is missing
        missing_targets = any(pd.isna(row[col]) or row[col] == "" 
                              for col in ['Número de contrato', 'Orden de suministro', 'Número de factura'])
        if not missing_targets:
            continue  # Nothing missing, skip row

        match = None
        pivot_used = None
        
        # First try matching using Folio fiscal (df_merged) vs UUID (df_insabi)
        if pd.notna(row.get('Folio fiscal')) and row['Folio fiscal'] != "":
            candidate = df_insabi[df_insabi['UUID'] == row['Folio fiscal']]
            if not candidate.empty:
                match = candidate.iloc[0]
                pivot_used = "Folio fiscal (UUID)"
        
        # If no match, try matching using Número de factura vs Factura
        if match is None and pd.notna(row.get('Número de factura')) and str(row['Número de factura']) != "":
            candidate = df_insabi[df_insabi['Factura'] == str(row['Número de factura'])]
            if not candidate.empty:
                match = candidate.iloc[0]
                pivot_used = "Número de factura (Factura)"
        
        # If still no match, try matching using Orden de suministro vs NÚMERO DE ORDEN DE SUMINISTRO
        if match is None and pd.notna(row.get('Orden de suministro')) and row['Orden de suministro'] != "":
            candidate = df_insabi[df_insabi['NÚMERO DE ORDEN DE SUMINISTRO'] == row['Orden de suministro']]
            if not candidate.empty:
                match = candidate.iloc[0]
                pivot_used = "Orden de suministro (NÚMERO DE ORDEN DE SUMINISTRO)"
        
        if match is None:
            print(f"Row {idx}: No matching INSABI row found using any pivot {row}.")
            continue
        
        # If a match is found, update only missing values:
        if (pd.isna(row['Número de contrato']) or row['Número de contrato'] == "") and 'Número de contrato' in match:
            df_merged.at[idx, 'Número de contrato'] = match['Número de contrato']
            print(f"Row {idx}: Filled 'Número de contrato' using pivot: {pivot_used}.")
        
        if (pd.isna(row['Orden de suministro']) or row['Orden de suministro'] == "") and 'NÚMERO DE ORDEN DE SUMINISTRO' in match:
            df_merged.at[idx, 'Orden de suministro'] = match['NÚMERO DE ORDEN DE SUMINISTRO']
            print(f"Row {idx}: Filled 'Orden de suministro' using pivot: {pivot_used}.")
        
        if (pd.isna(row['Número de factura']) or row['Número de factura'] == "") and 'Factura' in match:
            df_merged.at[idx, 'Número de factura'] = match['Factura']
            print(f"Row {idx}: Filled 'Número de factura' using pivot: {pivot_used}.")
        if (pd.isna(row['CLUES']) or row['CLUES'] == "") and 'CLUES' in match:
            df_merged.at[idx, 'CLUES'] = match['CLUES']
            print(f"Row {idx}: Filled 'CLUES' using pivot: {pivot_used}.")        
        # Optionally, you could also update 'Folio fiscal' if desired:
        if (pd.isna(row['Folio fiscal']) or row['Folio fiscal'] == "") and 'UUID' in match:
            df_merged.at[idx, 'Folio fiscal'] = match['UUID']
            print(f"Row {idx}: Filled 'Folio fiscal' using pivot: {pivot_used}.")
    
    return df_merged

    
def join_SAGI_files(json_key, file2023_2024, file_2024, output_joined_file):
    """
    Merges two Excel files (one for 2023-2024 and one for 2024), 
    cleans and enriches the data using a Google Sheet (INSABI), 
    saves the joined result to an output Excel file, and uploads the merged data 
    to the 'EstatusSAGI' worksheet in the target Google Sheet.
    
    Parameters:
      json_key: Path to the JSON key file for the service account.
      file2023_2024: Path to the Excel file for 2023-2024.
      file_2024: Path to the Excel file for 2024.
      output_joined_file: Path (including filename) where the joined Excel file will be saved.
    """
    global spreadsheet  # Ensure worksheet_to_df can access this global variable
    
    # Define the scope for Google Sheets and Drive
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    
    # Authorize using the provided JSON key
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_key, scope)
    client = gspread.authorize(creds)
    
    # Open the target Google Sheet by URL
    spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1KN4XwXQlZ5jhyErxdpA_R6kNk2tKwwFeDnCDuHCr0fo/edit#gid=2033397596')
    
    # Load the Excel files into DataFrames
    df_2023_2024 = pd.read_excel(file2023_2024)
    df_2024 = pd.read_excel(file_2024)
    
    # Load INSABI data from Google Sheet using worksheet_to_df (which uses the global variable "spreadsheet")
    df_INSABI_gsheet = worksheet_to_df('RAW_INSABI', 'A:AA')
    
    # Merge the Excel DataFrames
    df_merged = pd.concat([df_2023_2024, df_2024], ignore_index=True)
    
    # Drop rows where 'Estado de la factura' is 'Cancelado'
    df_merged = df_merged[df_merged['Estado de la factura'] != 'Cancelado']
    
    # Add a 'Factura' column by adding the 'P-' prefix to 'Número de factura'
    df_merged['Factura'] = 'P-' + df_merged['Número de factura'].astype(str)
    new_rows = [
        {'UUID': 'D58B6D34-45E7-4DB2-BF7C-C1FDDCB8EC0A', 'Factura': 'P-1830', 'NÚMERO DE ORDEN DE SUMINISTRO': 'U00-28-02-2023-284486-F7', 'CLUES': 'TSSSA017786'},
        {'UUID': '475546B7-19AF-471C-B6B1-D8A16850618E', 'Factura': 'P-1828', 'NÚMERO DE ORDEN DE SUMINISTRO': 'U00-28-02-2023-284485-F7', 'CLUES': 'TSSSA017786'}
    ]
    
    new_df = pd.DataFrame(new_rows)
    df_INSABI_gsheet = pd.concat([df_INSABI_gsheet, new_df], ignore_index=True) 
    
    print("\n DATAFRAME PREVIO")
    print(df_merged.info())
    print("\n *******")
    df_merged = complete_missing_values(df_merged, df_INSABI_gsheet)
    
    print("Extra columns merged based on 'Orden de suministro' match:")
    print(df_merged.head())    
    #print(df_merged.head(15))
    print("\n ******* dataframe completado con la función.")
    print(df_merged.info())
    
    # Drop unnecessary columns
    #df_merged.drop(columns=['Unnamed: 0', 'Número de oficio', 'Número de factura', 'Opciones'], inplace=True)
    
    # Convert the 'Total' column to integer after removing the currency sign
    df_merged['Total'] = df_merged['Total'].replace(r'[\$,]', '', regex=True).astype(float).astype(int)
    
    # Save the merged DataFrame to the specified output Excel file
    df_merged.to_excel(output_joined_file, index=False)
    print(f"File saved as {output_joined_file}")
    
    # Upload the merged DataFrame to the 'EstatusSAGI' worksheet in the Google Sheet
    gsheet_estatusSAGI = spreadsheet.worksheet('EstatusSAGI')
    gsheet_estatusSAGI.clear()
    set_with_dataframe(gsheet_estatusSAGI, df_merged)
    print("Estatus SAGI actualizado en el Google Sheet")