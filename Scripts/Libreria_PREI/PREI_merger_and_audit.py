import os
import sys
import glob
import subprocess
import datetime
import pandas as pd
import pyperclip
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe
import shutil
import re

def move_files(temp_folder, final_folder):
    """
    Deletes specific .xls files from the final_folder (PISP directory)
    and then moves all .xls files from the temp_folder (Descarga PISP) to final_folder.
    """
    # Deleting specific .xls files from final_folder (PISP)
    pisp_dir = final_folder
    exclude_files = ["[febrero pagado.xls", "[Veracuz pagada.xls"]
    for file in glob.glob(os.path.join(pisp_dir, "*.xls")):
        file_name = os.path.basename(file)
        if file_name not in exclude_files:
            try:
                os.remove(file)
                print(f"Deleted file: {os.path.basename(file)}")
            except Exception as e:
                print(f"Error deleting {os.path.basename(file)}: {e}")

    # Moving all .xls files from temp_folder (Descarga PISP) to final_folder (PISP)
    descarga_pisp_dir = temp_folder
    for file in glob.glob(os.path.join(descarga_pisp_dir, "*.xls")):
        try:
            shutil.move(file, pisp_dir)
            print(f"Moved file: {os.path.basename(file)} to {os.path.basename(pisp_dir)}")
        except Exception as e:
            print(f"Error moving {os.path.basename(file)}: {e}")

def merge_files(temp_folder, final_folder):
    """
    Reads all .xls files in temp_folder, merges them into a single DataFrame,
    updates a Google Sheet with the combined data, saves the merged Excel file in final_folder,
    prints a summary report, and then calls the audit function.
    """
    # Use temp_folder as the current directory for source files.
    current_directory = final_folder
    files_path = os.path.join(current_directory, "*.xls")
    files = glob.glob(files_path)
    print('\nAbriendo los exceles fuentes:')
    if not files:
        print(f"No se encontraron archivos .xls en {os.path.basename(current_directory)}")
        return

    # Read and merge data from all found Excel files.
    dfs = []
    headers = ["Documento", "Folio Fiscal", "Fecha Factura", "Importe", 
               "Fecha Carga", "Unidad Negocio", "Contra Recibo", "Estado C.R."]
    for file in files:
        try:
            df = pd.read_excel(file, header=None, skiprows=range(8), usecols='A:H')
            df.columns = headers
            dfs.append(df)
        except Exception as e:
            print(f"Error leyendo el archivo {file}: {e}")
    
    if not dfs:
        print("No se pudieron leer datos de los archivos.")
        return
    combined_data = pd.concat(dfs, ignore_index=True)
    print('Combinando exceles en la carpeta:', {os.path.basename(current_directory)}, "\n")
    
    # Check for duplicated 'Folio Fiscal'

    if combined_data.duplicated('Folio Fiscal').any():
        # Identify and print all duplicated rows before merging
        duplicates = combined_data[combined_data.duplicated('Folio Fiscal', keep=False)]
        print("Duplicated rows before merging:")
        print(duplicates)
        
        # For duplicated 'Folio Fiscal', join the 'Estado C.R.' values into a single string
        # and take the first value for the remaining columns.
        combined_data = combined_data.groupby('Folio Fiscal', as_index=False).agg({
            'Documento': 'first',
            'Fecha Factura': 'first',
            'Importe': 'first',
            'Fecha Carga': 'first',
            'Unidad Negocio': 'first',
            'Contra Recibo': 'first',
            'Estado C.R.': lambda x: ', '.join(x.astype(str).unique())
        })
        
        # Check again for any duplicated 'Folio Fiscal'
        if combined_data.duplicated('Folio Fiscal').any():
            print("'Estado C.R.' joined but still duplicated 'Folio Fiscal' exist!")
            
                
    now = datetime.datetime.now()
    # Save the merged file in the final_folder.
    filename = os.path.join(final_folder, now.strftime("%Y %m %d") + ' PISP ' + now.strftime("%H") + 'h.xlsx')
    
    # Add the current date to the DataFrame.
    combined_data["Fecha"] = now.strftime("%d/%m/%Y")
    
    # Update a Google Sheet with the combined data: No es necesario por ahora. 0
    """
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    print('\nIngresando a google sheet Management dashboard.')
    key_file_path = os.path.join(current_directory, 'key.json')  # Assumes key.json is here.
    try:
        credentials = ServiceAccountCredentials.from_json_keyfile_name(key_file_path, scope)
        gc = gspread.authorize(credentials)
        spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1KN4XwXQlZ5jhyErxdpA_R6kNk2tKwwFeDnCDuHCr0fo/edit#gid=1459325443'
        spreadsheet = gc.open_by_url(spreadsheet_url)
        sheet_raw_pisp = spreadsheet.worksheet('RAW_PISP2023')
        sheet_raw_pisp.clear()
        set_with_dataframe(sheet_raw_pisp, combined_data)
        print("Consola general actualizada")
    except Exception as e:
        print("Error updating Google Sheet:", e)
    """
    sorting_columns = ["Documento", "Folio Fiscal", "Fecha Factura", "Importe", "Fecha Carga", "Unidad Negocio", "Contra Recibo", "Estado C.R.", "Fecha"]
    combined_data = combined_data[sorting_columns]
    # Save the merged data to an Excel file.
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        combined_data.to_excel(writer, sheet_name='Sheet1', startrow=0, index=False)
    print('Reporte al', now.strftime("%d/%m/%y %H:%M"))
    
    # Calculate and print summaries.
    en_proceso_sum = combined_data.loc[combined_data['Estado C.R.'] == 'En proceso', 'Importe'].sum()
    aprobados_sum = combined_data.loc[combined_data['Estado C.R.'] == 'Aprobado', 'Importe'].sum()
    pagado_sum = combined_data.loc[combined_data['Estado C.R.'] == 'Pagado', 'Importe'].sum()
    total_sum = en_proceso_sum + aprobados_sum + pagado_sum
    result_string = f'''
Reporte Al: {now.strftime("%d/%m/%y %H:%M")}

En proceso: {en_proceso_sum:,.2f}
Aprobado: {aprobados_sum:,.2f}
Pagado: {pagado_sum:,.2f}
Total: {total_sum:,.2f}
'''
    print(result_string)
    pyperclip.copy(result_string)
    print('\nInformación copiada al portapapeles')
    """
    # Call the audit function.
    audit(final_folder)
    
    # Read and display a preview of the audit file.
    audit_file_path = os.path.join(final_folder, 'audit.xlsx')
    try:
        audit_df = pd.read_excel(audit_file_path, usecols='A:F')
        print("Archivo audit.xlsx generado. Preview:")
        print(audit_df.head())
    except Exception as e:
        print("Error leyendo audit.xlsx:", e)
    """

def audit(final_folder, year):
    """
    Processes all Excel files in final_folder with "PISP" in the name,
    extracts key columns, calculates the earliest dates for different statuses,
    and then writes an audit Excel file named "{year} audit.xlsx" to final_folder.
    
    Parameters:
        final_folder (str): The directory containing the Excel files.
        year (str or int): The year to include in the output filename.
    """
    # Directory containing the Excel files
    directory = final_folder

    # Dictionary to store the extracted data
    data_dict = {
        'Folio Fiscal': [],
        'Estado C.R.': [],
        'Fecha': [],
        'File Date': [],
        'Importe': [],
        'Unidad Negocio': []
    }

    # Loop through all files in the directory
    for filename in os.listdir(directory):
        # Check if the file has the correct naming pattern and is an Excel file
        if filename.endswith('.xlsx') and "PISP" in filename:
            # Extract the date from the filename.
            # Assumes filename format: "DD MM YYYY ...", adjust splitting if needed.
            parts = filename.split(' ')
            if len(parts) < 3:
                continue
            file_date = parts[2] + '/' + parts[1] + '/' + parts[0]
            
            # Construct the full file path and read the Excel file
            filepath = os.path.join(directory, filename)
            try:
                df = pd.read_excel(filepath, engine='openpyxl')
            except Exception as e:
                print(f"Error reading {filename}: {e}")
                continue
            
            # Append the required columns to the dictionary
            data_dict['Folio Fiscal'].extend(df['Folio Fiscal'].tolist())
            data_dict['Estado C.R.'].extend(df['Estado C.R.'].tolist())
            data_dict['Fecha'].extend(df['Fecha'].tolist())
            data_dict['File Date'].extend([file_date] * len(df))
            data_dict['Importe'].extend(df['Importe'].tolist())
            data_dict['Unidad Negocio'].extend(df['Unidad Negocio'].tolist())

    # Convert the dictionary to a DataFrame
    data_df = pd.DataFrame(data_dict)

    # Create a new DataFrame for the audit file with unique 'Folio Fiscal'
    audit_df = pd.DataFrame(data_df['Folio Fiscal'].unique(), columns=['Folio Fiscal'])
    audit_df['F. En proceso'] = None
    audit_df['F. Aprobado'] = None
    audit_df['F. Pagado'] = None
    audit_df['Importe'] = None
    audit_df['Unidad Negocio'] = None

    # For each unique Folio Fiscal:
    for index, row in audit_df.iterrows():
        folio = row['Folio Fiscal']
        # Extract associated Importe and Unidad Negocio values from the first matching row
        subset_values = data_df[data_df['Folio Fiscal'] == folio].iloc[0]
        audit_df.at[index, 'Importe'] = subset_values['Importe']
        audit_df.at[index, 'Unidad Negocio'] = subset_values['Unidad Negocio']
        
        # Find the earliest date with "Estado C.R." = "En proceso"
        subset_en_proceso = data_df[(data_df['Folio Fiscal'] == folio) & (data_df['Estado C.R.'] == 'En proceso')]
        if not subset_en_proceso.empty:
            earliest_date_en_proceso = min(pd.to_datetime(subset_en_proceso['Fecha'], dayfirst=True))
            audit_df.at[index, 'F. En proceso'] = earliest_date_en_proceso.strftime('%d/%m/%Y')
        
        # Find the earliest date with "Estado C.R." = "Aprobado"
        subset_aprobado = data_df[(data_df['Folio Fiscal'] == folio) & (data_df['Estado C.R.'] == 'Aprobado')]
        if not subset_aprobado.empty:
            earliest_date_aprobado = min(pd.to_datetime(subset_aprobado['Fecha'], dayfirst=True))
            audit_df.at[index, 'F. Aprobado'] = earliest_date_aprobado.strftime('%d/%m/%Y')

        # Find the earliest date with "Estado C.R." = "Pagado"
        subset_pagado = data_df[(data_df['Folio Fiscal'] == folio) & (data_df['Estado C.R.'] == 'Pagado')]
        if not subset_pagado.empty:
            earliest_date_pagado = min(pd.to_datetime(subset_pagado['Fecha'], dayfirst=True))
            audit_df.at[index, 'F. Pagado'] = earliest_date_pagado.strftime('%d/%m/%Y')

    # Save the audit DataFrame to a new Excel file
    output_filename = f'{year} audit.xlsx'
    output_path = os.path.join(directory, output_filename)
    try:
        audit_df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"Data extraction and processing complete. {os.path.basename(output_path)} has been updated.")
    except Exception as e:
        print(f"Error saving audit file: {e}")


def fusion_2023_2025(valid_years, folder_root):
    """
    Fusiona ciclos fiscales 2023, 2024, 2025.
    
    Busca y fusiona archivos Excel del día actual para cada año en valid_years. 
    El nombre de los archivos debe seguir el formato:
      "{today.year} {month:02d} {day:02d} PISP {hour:02d}h.xlsx"
    Si no se encuentra un archivo para el día de hoy en un ciclo, imprime un mensaje
    y continúa con el siguiente año.
    """
    # Define expected headers (adjust as needed)
    expected_headers = ["Documento", "Folio Fiscal", "Fecha Factura", "Importe", 
                        "Fecha Carga", "Unidad Negocio", "Contra Recibo", "Estado C.R.", "Fecha"]

    # List to collect DataFrames from each recent xlsx file
    dataframes = []
    
    # Get today's date
    today = datetime.datetime.now()
    # Format month and day with two digits each, separated by a space (e.g., "03 20")
    month_day_str = today.strftime("%m %d")
    print(f"DEBUG: Today's date for file search: {today.strftime('%Y %m %d')}")
    
    # Loop through each valid year folder
    for year in valid_years:
        final_folder = os.path.join(folder_root, 'Implementación', 'PREI', f"{year} Final")
        print(f"DEBUG: Looking in folder: {os.path.basename(final_folder)}")
        
        # Build glob pattern for today's file using today's year instead of the folder year.
        pattern = os.path.join(final_folder, f"{today.year} {month_day_str} PISP *h.xlsx")
        print(f"DEBUG: Using glob pattern: {os.path.basename(pattern)}")
        
        files = glob.glob(pattern)
        print("DEBUG: Files found:")
        for item in files:
            print(os.path.basename(item))
        
        if not files:
            print(f"\n**** No file found for today for folder {year} Final ****\n")
            continue
        
        # Extract the hour from the filename and choose the file with the highest hour.
        file_hour_pairs = []
        for file in files:
            # Regex to match pattern like "PISP 09h.xlsx" at the end of the filename.
            match = re.search(r"PISP (\d{2})h\.xlsx$", file)
            if match:
                hour = int(match.group(1))
                file_hour_pairs.append((hour, file))
                print(f"DEBUG: Found file: {os.path.basename(file)} with hour: {hour}")
            else:
                print(f"DEBUG: File {file} does not match expected pattern.")
        
        if not file_hour_pairs:
            print(f"\n**** No file found for today for folder {year} Final after pattern matching ****\n")
            continue
        
        # Choose the file with the maximum hour value (most recent)
        recent_file = max(file_hour_pairs, key=lambda x: x[0])[1]
        print(f"DEBUG: Selected recent file for folder {year} Final: {os.path.basename(recent_file)}")
        
        # Load the recent file
        try:
            df = pd.read_excel(recent_file, engine='openpyxl')
            print(f"DEBUG: Successfully read file: {os.path.basename(recent_file)}")
        except Exception as e:
            print(f"Error reading {recent_file}: {e}")
            continue
        
        # Check if the file's headers match the expected headers
        if list(df.columns) != expected_headers:
            print(f"Headers in {os.path.basename(recent_file)} do not match expected headers. Skipping this file.")
            print(f"DEBUG: File headers: {list(df.columns)}")
            continue
        
        dataframes.append(df)
    
    # Merge all valid DataFrames into one master DataFrame and save
    if dataframes:
        master_df = pd.concat(dataframes, ignore_index=True)
        master_file = os.path.join(folder_root, 'Implementación', 'PREI', "PREI.xlsx")
        master_df.to_excel(master_file, index=False, engine='openpyxl')
        print(f"Master file saved as: {master_file}")
    else:
        print("No valid files found to merge.")