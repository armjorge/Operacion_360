
import os
import glob
import pandas as pd
from openpyxl import load_workbook
import warnings
warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")
import yaml
"""def merge_dataframes(dataframes):"""
from datetime import datetime
from collections import defaultdict
import re

def xlsx_loading(download_directory, headers, output_name):
    """
    Load all .xlsx files in the given directory that match the expected header keys.
    
    Parameters:
        download_directory (str): Path to the directory containing the Excel files.
        headers (dict): Dictionary where the keys are the expected column names and 
                        the values represent the desired data types.
    
    Returns:
        pd.DataFrame: A DataFrame combining the valid Excel files, each appended with a new column "File_date"
                      that contains the creation date from the Excel file properties.
    """
    # Use glob to find all .xlsx files in the download_directory.
    xlsx_files = glob.glob(os.path.join(download_directory, "*.xlsx"))
    xlsx_files = [f for f in xlsx_files if os.path.basename(f) != output_name]
    expected_headers = set(headers.keys())
    
    # List to hold files with matching headers.
    valid_files = []
    
    # Filter files: only keep the ones where the file's headers contain the expected keys.
    for file in xlsx_files:
        try:
            # Read only the first row to get the column headers.
            temp_df = pd.read_excel(file, nrows=0)
            file_headers = set(temp_df.columns)
            
            # Check if the file contains at least the expected headers.
            if expected_headers.issubset(file_headers):
                valid_files.append(file)
            else:
                print(f"file {os.path.basename(file)} skipped due to headers mis match")
        except Exception as e:
            print(f"Error reading headers from {os.path.basename(file)}: {e}")
    
    dataframes = []
    
    # For each valid file load the data and add a "File_date" column.
    for file in valid_files:
        try:
            # Load the creation date from the Excel file metadata using openpyxl.
            wb = load_workbook(file, read_only=True, data_only=True)
            created_date = wb.properties.created
        except Exception as e:
            print(f"Could not get created date from file {os.path.basename(file)}: {e}")
            created_date = None  # Fallback to None if there is an error.
            
        try:
            # Load the full Excel data.
            df = pd.read_excel(file)
            # Append the creation date as a new column.
            date_column = "File_date"
            df[date_column] = created_date
            dataframes.append(df)
        except Exception as e:
            print(f"Error loading data from {os.path.basename(file)}: {e}")
    
    if dataframes:
        # Concatenate all loaded dataframes into one.
        combined_df = pd.concat(dataframes, ignore_index=True)
        # Iterate over the headers and convert columns with datetime dtype.
        for col, dtype in headers.items():
            if dtype == 'datetime64[ns]' and col in combined_df.columns:
                combined_df[col] = pd.to_datetime(combined_df[col], format='%d/%m/%Y', errors='coerce')
                
        #print(combined_df.head())
        return combined_df, date_column
    else:
        print("No valid dataframes loaded.")
        return None


def merge_dataframes(dataframes, duplicates_column, date_column):
    """
    Merge duplicate rows in a DataFrame by keeping the newest row for each duplicate entry.
    
    Parameters:
        dataframes (pd.DataFrame): The input DataFrame containing potential duplicates in duplicates_column.
        duplicates_column (str): The column name used to identify duplicate rows.
        date_column (str): The column name containing dates that determine the "newest" row for duplicates.
        
    Returns:
        pd.DataFrame: A DataFrame with duplicates removed and only the newest row retained.
    """
    # Sort the DataFrame by date_column so that the most recent date is last.
    sorted_df = dataframes.sort_values(by=date_column, ascending=True)
    
    # Drop duplicates based on duplicates_column, keeping the last occurrence (i.e. the newest row).
    merged_df = sorted_df.drop_duplicates(subset=duplicates_column, keep='last')
    
    return merged_df

def CAMUNDA_merging(download_directory, headers, columna_duplicados, output_name):
    print(message_print('Iniciando la fusión de xlsx para la versión 2021-2024 de CAMUNDA'))
    print("Loading Camunda Merging: merge all xlsx file while keeping recent row")
    #print(download_directory)
    #print(headers)
    download_directory = os.path.join(download_directory, "2023-2024")
    dataframe, date_column  = xlsx_loading(download_directory, headers, output_name)
    #print(date_column, "\n", dataframe.head() )
    print(message_print("A partir de la fecha de los archivos, conservaremos la columna más reciente sin duplicar los renglones"))
    final_dataframe = merge_dataframes(dataframe, columna_duplicados, date_column)
    #print(final_dataframe.head())
    # Save final_dataframe to the file in the download_directory.
    download_directory = os.path.join(download_directory, "..")
    final_path = os.path.join(download_directory, output_name)
    final_dataframe.to_excel(final_path, index=False)
    print(f"**********\nArchivo final guardado en {os.path.basename(download_directory)} / {output_name}\n *************")

def message_print(message): 
    message_highlights= '*' * len(message)
    message = f'\n{message_highlights}\n{message}\n{message_highlights}\n'
    return message

def csv_renaming(download_directory):
    # 1. Gather all .csv files
    pattern_csv = os.path.join(download_directory, '*.csv')
    all_csvs = glob.glob(pattern_csv)

    # 2. Exclude files starting with YYYY MM DD  
    date_prefix = re.compile(r'^\d{4}\s\d{2}\s\d{2}')
    filtered = [
        f for f in all_csvs
        if not date_prefix.match(os.path.basename(f))
    ]

    # 3. Map each file to its modification datetime
    file_dates = {
        f: datetime.fromtimestamp(os.path.getmtime(f))
        for f in filtered
    }

    # Sort filtered list by date for consistent grouping
    filtered.sort(key=lambda f: file_dates[f])

    # 4. Group by (year, month)
    groups = defaultdict(list)
    for f in filtered:
        dt = file_dates[f]
        groups[(dt.year, dt.month, dt.day)].append(f)

    # 5. Process each year-month group
    for (year, month, day), files in groups.items():
        # load into DataFrames, keeping track of filenames
        dfs = {f: pd.read_csv(f) for f in files}

        # detect duplicates: keep first, mark the rest for removal
        to_remove = set()
        file_list = list(dfs.keys())
        for i, f1 in enumerate(file_list):
            if f1 in to_remove:
                continue
            for f2 in file_list[i+1:]:
                if f2 in to_remove:
                    continue
                if dfs[f1].equals(dfs[f2]):
                    to_remove.add(f2)

        # remove duplicate files from disk
        for dup in to_remove:
            os.remove(dup)
            print(f"Removed duplicate: {os.path.basename(dup)}")

        # rename remaining files with suffixes if more than one
        remaining = [f for f in files if f not in to_remove]
        if len(remaining) > 1:
            for idx, f in enumerate(remaining, start=1):
                year, month, day = file_dates[f].year, file_dates[f].month, file_dates[f].day
                new_name = f"{year:04d} {month:02d} {day:02d}_{idx}.csv"
                new_path = os.path.join(download_directory, new_name)
                os.rename(f, new_path)
                print(f"Renamed {os.path.basename(f)} → {os.path.basename(new_name)}")

def merging_and_updating_camunda2025(implementacion):
    print(message_print('Iniciando la fusión de CSV para la versión 2025 de CAMUNDA'))
    download_directory = os.path.join(implementacion, "CAMUNDA", "Descargas", "2025")
    print(message_print('Renombrando archivos CSV'))
    csv_renaming(download_directory)
    # Step 1: Load all CSV files with a date extracted from filename or filesystem metadata
    all_files = [
        os.path.join(download_directory, f)
        for f in os.listdir(download_directory)
        if f.endswith('.csv')
    ]

    # Extract file date based on last modified time
    dated_dfs = []
    for file in all_files:
        file_date = datetime.fromtimestamp(os.path.getmtime(file))
        df = pd.read_csv(file)
        df['__file_date'] = file_date
        dated_dfs.append(df)

    # Step 2: Combine all into one dataframe
    combined_df = pd.concat(dated_dfs, ignore_index=True)

    # Step 3: Sort by 'numero_orden_suministro' and then by '__file_date' to get latest info
    combined_df.sort_values(by=['numero_orden_suministro', '__file_date'], ascending=[True, False], inplace=True)

    # Step 4: Keep only the latest version for each 'numero_orden_suministro'
    latest_rows = combined_df.drop_duplicates(subset='numero_orden_suministro', keep='first')

    # Step 5: Load historical merged file if it exists
    merged_file_path = os.path.join(implementacion,"CAMUNDA", 'Camunda.xlsx')
    if os.path.exists(merged_file_path):
        df_merged_updated = pd.read_excel(merged_file_path)
    else:
        df_merged_updated = pd.DataFrame(columns=latest_rows.columns.drop('__file_date'))

    # Step 6: Merge with latest info, updating `descripcion_estatus_orden_suministro` if changed
    df_merged_updated = df_merged_updated.set_index('numero_orden_suministro')
    latest_rows = latest_rows.set_index('numero_orden_suministro')

    for idx, new_row in latest_rows.iterrows():
        if idx not in df_merged_updated.index:
            # New row
            df_merged_updated.loc[idx] = new_row.drop('__file_date')
        else:
            # Existing row, check if estatus changed
            current_status = df_merged_updated.at[idx, 'descripcion_estatus_orden_suministro']
            new_status = new_row['descripcion_estatus_orden_suministro']
            if current_status != new_status:
                df_merged_updated.at[idx, 'descripcion_estatus_orden_suministro'] = new_status

    df_merged_updated.reset_index(inplace=True)
    # Save the merged file back for future usage
    df_merged_updated.to_excel(merged_file_path, index=False)

    return df_merged_updated

if __name__ == "__main__":
    # Esto solo se ejecuta cuando corres el script, no cuando lo importas desde otro módulo.
    implementacion =os.path.join( os.getcwd(), "Implementación")
    download_directory = os.path.join(implementacion, "CAMUNDA", "Descargas")
    
    print(download_directory)
    yaml_file = os.path.join(implementacion, "df_headers.yaml")
    with open(yaml_file, "r", encoding="utf-8") as file:
        data = yaml.load(file, Loader=yaml.FullLoader)
    INSABI_headers = data.get("columns_INSABI")
    duplicados = "NÚMERO DE ORDEN DE SUMINISTRO"
    archivo_final = "Camunda 2023-2025.xlsx"
    merging_and_updating_camunda2025(implementacion)
    CAMUNDA_merging(download_directory, INSABI_headers, duplicados, archivo_final)
