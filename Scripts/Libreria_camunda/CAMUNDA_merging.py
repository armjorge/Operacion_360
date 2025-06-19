
import os
import glob
import pandas as pd
from openpyxl import load_workbook
import warnings
warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")
"""def merge_dataframes(dataframes):"""
from datetime import datetime

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
    print("Loading Camunda Merging: merge all xlsx file while keeping recent row")
    #print(download_directory)
    #print(headers)
    dataframe, date_column  = xlsx_loading(download_directory, headers, output_name)
    #print(date_column, "\n", dataframe.head() )
    final_dataframe = merge_dataframes(dataframe, columna_duplicados, date_column)
    #print(final_dataframe.head())
    # Save final_dataframe to the file in the download_directory.
    final_path = os.path.join(download_directory, output_name)
    final_dataframe.to_excel(final_path, index=False)
    print(f"**********\nArchivo final guardado en {os.path.basename(download_directory)} / {output_name}\n *************")


def merging_and_updating_camunda2025(download_directory):
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
    merged_file_path = os.path.join(download_directory, '..', 'Camunda v2025.xlsx')
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
    download_directory = r"C:\Users\arman\Dropbox\3. Armando Cuaxospa\Adjudicaciones\Licitaciones 2025\E115 360\Implementación\CAMUNDA\Descargas"
    merging_and_updating_camunda2025(download_directory)
