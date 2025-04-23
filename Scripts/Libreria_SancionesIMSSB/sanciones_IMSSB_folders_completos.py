import os
import pandas as pd
import shutil
import PyPDF2

"""
# googlesheet
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe
import sys


scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

try:
    # Load credentials
    creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)
    client = gspread.authorize(creds)

    # Try to open the spreadsheet
    spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1KN4XwXQlZ5jhyErxdpA_R6kNk2tKwwFeDnCDuHCr0fo/edit#gid=2033397596'
    spreadsheet = client.open_by_url(spreadsheet_url)

    print("✅ Connected successfully to the spreadsheet:", spreadsheet.title)

except FileNotFoundError:
    print("❌ key.json file not found. Make sure it is in the correct folder.")
    sys.exit(1)

except gspread.exceptions.APIError as e:
    print("❌ Google API error:", e)
    print("🔐 Did you share the sheet with your service account email?")
    sys.exit(1)

except Exception as e:
    print("❌ Unexpected error:", e)
    import traceback
    traceback.print_exc()
    sys.exit(1)
"""

def check_oficio_en_cada_folder(root_directory):
    complete_folders = []
    incomplete_folders = []
    
    # List all items in the root_directory
    for item in os.listdir(root_directory):
        # Construct the path to the item
        path = os.path.join(root_directory, item)
        
        # Check if the item is a directory
        if os.path.isdir(path):
            # Look for a PDF with the same name as the directory
            pdf_name = item + '.pdf'
            pdf_path = os.path.join(path, pdf_name)
            
            # Check if the PDF file exists in the directory
            if os.path.isfile(pdf_path):
                complete_folders.append(item)
            else:
                incomplete_folders.append(item)
    
    # Return the sets of complete and incomplete folders
    return complete_folders, incomplete_folders

def check_excel_relations(root_directory):
    complete_folders = []
    incomplete_folders = []
    complete_files = []
    incomplete_files = []
    
    for item in os.listdir(root_directory):
        path = os.path.join(root_directory, item)
        if os.path.isdir(path):
            excel_name = item + '_relacion.xlsx'
            excel_path = os.path.join(path, excel_name)
            
            if os.path.isfile(excel_path):
                complete_folders.append(item)
                try:
                    # Load the Excel file
                    df = pd.read_excel(excel_path)
                    # Check if required columns exist
                    if 'ORDEN DE SUMINISTRO' in df.columns and 'PENA' in df.columns:
                        complete_files.append(excel_name)
                    else:
                        incomplete_files.append(excel_name)
                except Exception as e:
                    print(f"Error reading {excel_name}: {e}")
                    incomplete_files.append(excel_name)
            else:
                incomplete_folders.append(item)
    
    return complete_folders, incomplete_folders, complete_files, incomplete_files

def process_folders(base_path):
    all_data = []
    for folder in os.listdir(base_path):
        folder_path = os.path.join(base_path, folder)
        if os.path.isdir(folder_path):
            excel_file = os.path.join(folder_path, f"{folder}_relacion.xlsx")
            if os.path.exists(excel_file):
                # Read the specified columns from the Excel file
                data = pd.read_excel(excel_file, usecols=['ORDEN DE SUMINISTRO', 'PENA'])
                # Remove spaces from 'ORDEN DE SUMINISTRO'
                data['ORDEN DE SUMINISTRO'] = data['ORDEN DE SUMINISTRO'].str.replace(' ', '', regex=True)
                
                # Convert 'PENA' to numeric, set errors='coerce' to handle non-convertible values
                data['PENA'] = pd.to_numeric(data['PENA'], errors='coerce')                
                # Add the folder name as a new column
                data['OFICIO'] = folder
                
                # Append the DataFrame to the list
                all_data.append(data)

                # Recursively process subfolders
                process_folders(folder_path)

    # Merge all DataFrames into one
    """
    if all_data:
        final_data = pd.concat(all_data, ignore_index=True)
        # Save the merged DataFrame to a new Excel file
        final_data.to_excel(os.path.join(base_path, 'Penas-Oficios-Ordenes.xlsx'), index=False)
        """
    if all_data:
        final_data = pd.concat(all_data, ignore_index=True)

        # Define the file path
        file_path = os.path.join(base_path, 'Penas-Oficios-Ordenes.xlsx')

        # Create an ExcelWriter object with the target file path
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            # Write the DataFrame to a sheet named 'Desglose'
            final_data.to_excel(writer, sheet_name='Desglose', index=False)
        return final_data
    
    return None
            
def copy_files_with_suffix(source_directory, target_directory, suffixes):
    for root, dirs, files in os.walk(source_directory):
        folder_name = os.path.basename(root)  # Get the last part of the current path, which is the folder name
        for file in files:
            for suffix in suffixes:
                expected_filename = folder_name + suffix
                if file == expected_filename:  # Check if the file matches the expected name
                    source_file_path = os.path.join(root, file)
                    destination_directory = os.path.join(target_directory, os.path.relpath(root, source_directory))

                    # Create the destination directory if it does not exist
                    if not os.path.exists(destination_directory):
                        os.makedirs(destination_directory)

                    destination_file_path = os.path.join(destination_directory, file)
                    
                    # Copy file
                    shutil.copy2(source_file_path, destination_file_path)
                    print(f"Copied '{source_file_path}' to '{destination_file_path}'")


def extract_dates_PDF(root_directory):
    def extract_text_from_pdf(pdf_path):
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ''
            for page in reader.pages:
                text += page.extract_text()
        return text    
    all_data = []
    for folder in os.listdir(root_directory):
        folder_path = os.path.join(root_directory, folder)
        if os.path.isdir(folder_path):
            pdf_file = os.path.join(folder_path, f'{folder}.pdf')
            if os.path.exists(pdf_file):
                text = extract_text_from_pdf(pdf_file)                               
                # Searching for 'Ciudad de México*' and capturing until line break
                import re
                match = re.search(r"Ciudad de México.*$", text, re.M)
                if match:
                    extracted_text = match.group(0)
                    print(extracted_text)
                    # Creating a DataFrame for each folder and appending to list
                    df = pd.DataFrame({'Oficio': [folder], 'Fecha': [extracted_text]})
                    df['Fecha'] = df['Fecha'].apply(lambda x: re.sub(r'^.*?(\d)', r'\1', x))             
                    df['Fecha'] = df['Fecha'].apply(lambda x: x.replace(" de ", "/"))                    
                    all_data.append(df)

    return all_data 


def folders_completos(root_directory):
    # If the file doesn’t exist yet, create it with two empty sheets
    file_path = os.path.join(root_directory, 'Penas-Oficios-Ordenes.xlsx')    
    if not os.path.exists(file_path):
        # Create an empty DataFrame
        empty_df = pd.DataFrame()
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            empty_df.to_excel(writer, sheet_name='Desglose', index=False)
            empty_df.to_excel(writer, sheet_name='Resumen',   index=False)
        print(f"Created new workbook with sheets 'Desglose' and 'Resumen' at:\n  {file_path}")    
    #root_directory = '.'  # Current directory as root
    con_oficio, sin_oficio = check_oficio_en_cada_folder(root_directory)
    
    #print(f"\n*******************************\nCarpetas con oficio:", con_oficio)
    comp_folders, inc_folders, comp_files, inc_files = check_excel_relations(root_directory)
    print(f"\n*******************************\nOficios con excel:", comp_folders)
    print("Exceles válidos:", comp_files)
    print(f"\nCarpetas Sin oficio", sin_oficio,"\n*******************************\n")
    print("Exceles incompletos:", inc_files)    
    print("Carpetas sin relación en excel:", inc_folders)

    final_data_desglose = process_folders(root_directory)
    #if final_data_desglose is not None:
        #print("\n*******************************\nReemplazando información en google sheet\n*******************************\n")
        #gsheet_SANCIONESINSABI = spreadsheet.worksheet('Sanciones_INSABI')
       # gsheet_SANCIONESINSABI.clear()
        #set_with_dataframe(gsheet_SANCIONESINSABI, final_data_desglose)
        #print("\n*******************************\nReemplazadas\n*******************************\n")
    #else:
    #    print("\n*******************************\nNo hay datos para reemplazar en el google sheet\n*******************************\n")
 
    summary_data = extract_dates_PDF(root_directory)
    if summary_data:
        final_data = pd.concat(summary_data, ignore_index=True)
        
        
        if os.path.exists(file_path):
            # Append/replace sheet in existing file
            with pd.ExcelWriter(
                file_path,
                engine='openpyxl',
                mode='a',
                if_sheet_exists='replace'
            ) as writer:
                final_data.to_excel(writer, sheet_name='Resumen', index=False)
        else:
            # Create new file
            with pd.ExcelWriter(
                file_path,
                engine='openpyxl',
                mode='w'
            ) as writer:
                final_data.to_excel(writer, sheet_name='Resumen', index=False)