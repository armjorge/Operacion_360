import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import numpy as np
import os
import shutil
import re
from PyPDF2 import PdfMerger


# Setup Google Sheets API
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1KN4XwXQlZ5jhyErxdpA_R6kNk2tKwwFeDnCDuHCr0fo/edit#gid=2033397596')
# Read the worksheets into DataFrames
def worksheet_to_df(worksheet_name, columns_range):
    worksheet = spreadsheet.worksheet(worksheet_name)
    data = worksheet.get(columns_range)
    return pd.DataFrame(data[1:], columns=data[0])
    
def search_and_copy(source_folder, target_folder, file_list): #Esta función busca en función de una lista de archivos
    for root, dirs, files in os.walk(source_folder):
        for file in files:
            for item in file_list:
                prefix, extension = os.path.splitext(item)  # Extracts prefix and the extension
                # Check if the file in the directory starts with the prefix and ends with the same extension
                if file.startswith(prefix) and file.endswith(extension):
                    source_file = os.path.join(root, file)
                    target_file_name = prefix + extension  # Clean file name to just prefix + extension
                    target_file_path = os.path.join(target_folder, target_file_name)
                    shutil.copy(source_file, target_file_path)
                    print(f"Copied: {target_file_name}")
                    break  # Once the correct file is found and copied, no need to check further for this item

def copy_files(source_dir, target_dir): #Esta función copia y pega de un lado al otro. 
    # Ensure the target directory exists
    os.makedirs(target_dir, exist_ok=True)
    
    # List all files in the source directory
    source_files = {file for file in os.listdir(source_dir) if os.path.isfile(os.path.join(source_dir, file))}
    
    # Copy files from source to target if they don't already exist
    for file in source_files:
        source_file_path = os.path.join(source_dir, file)
        target_file_path = os.path.join(target_dir, file)
        
        if not os.path.exists(target_file_path):
            shutil.copy2(source_file_path, target_file_path)
            print(f"Copied {file} to {target_dir}")
        else:
            #print(f"File {file} already exists in {target_dir}")
            pass

def check_directory_sync(source_dir, target_dir):
    source_files = set(os.listdir(source_dir))
    target_files = set(os.listdir(target_dir))
    
    if source_files == target_files:
        print(f"\n{source_dir} \nY\n \n{target_dir}\ntienen los mismos archivos")
    else:
        print("The directories do not have the same files.")

def generate_upload_list():
    # Filter df_altasFacturadas where ['PISP'] == '#N/A'
    df_uploadXMLlist = df_altasFacturadas[df_altasFacturadas['PISP'] == '#N/A']['Factura'] + '.xml'
    
    # Write df_uploadXMLlist to .\uploadXMLlist.csv
    df_uploadXMLlist.to_csv('.\\uploadXMLlist.csv', index=False)
    
    # Verify each file exists in .\XMLs
    missing_files = []
    for file_name in df_uploadXMLlist:
        if not os.path.exists(os.path.join('.\\XMLs', file_name)):
            missing_files.append(file_name)
    
    if missing_files:
        print("\n*******************************")
        print("These files were not found:")
        for file in missing_files:
            print(file)
        print("Populate, clean the names and try again.")
        print("*******************************")
    else:
        target_dir = '.\\XMLS a subir'
        # Ensure the target directory exists
        if not os.path.exists(target_dir):
            os.makedirs(target_dir)
            print(f"Created directory: {target_dir}")

        # Copy files
        for file_name in df_uploadXMLlist:
            src = os.path.join('.\\XMLs', file_name)
            dst = os.path.join(target_dir, file_name)
            shutil.copy2(src, dst)
            print(f"Copied {file_name} to {dst}")
        
        print("\n*******************************")
        print("Tienes que correr el script 01_UploadXML.py, subir XML, descargar acuse, renombrar acuses y volver para fusionar factura + acuse")
        print("*******************************")

def merge_altas_teoricas(df_altasFacturadas):
    df_selected = df_altasFacturadas
    acuses_folder = '.\\Acuses'
    pdfs_folder = '.\\PDFs'
    output_folder = '.\\Facturas y acuses'
    os.makedirs(output_folder, exist_ok=True)  # Ensure the output directory exists

    existing_files = {f for f in os.listdir(output_folder) if os.path.isfile(os.path.join(output_folder, f))}
    missing_files_list = []  # List to collect missing file information

    for index, row in df_selected.iterrows():
        factura = row['Factura']
        folio = row['Folio']  # Ensure 'Folio' column exists in your DataFrame
        no_alta = row.get('noAlta', '')  # Using .get to avoid KeyError if column does not exist
        no_orden = row.get('noOrden', '')  # Same as above
        output_file_name = f"{factura} {no_alta} {no_orden}.pdf"
        output_file = os.path.join(output_folder, output_file_name)

        # Skip merging if output file already exists
        if output_file_name in existing_files:
            print(f"File already exists and was skipped: {output_file_name}")
            continue

        acuse_file = os.path.join(acuses_folder, factura + '_.pdf')
        pdf_file = os.path.join(pdfs_folder, factura + '.pdf')

        # Check if both required files exist
        if not os.path.exists(acuse_file) or not os.path.exists(pdf_file):
            # If either file is missing, add details to the missing_files_list
            if not os.path.exists(acuse_file):
                missing_files_list.append({"Factura": factura, "Missing File": "Acuse", "Folio": folio})
                print(f"Missing Acuse file for: {factura}")
            if not os.path.exists(pdf_file):
                missing_files_list.append({"Factura": factura, "Missing File": "PDF", "Folio": folio})
                print(f"Missing PDF file for: {factura}")
        else:
            # Proceed with merge if both files are present
            merger = PdfMerger()
            merger.append(pdf_file)
            merger.append(acuse_file)
            # Write the merged PDF to the output file
            with open(output_file, "wb") as fout:
                merger.write(fout)
                merger.close()
            print(f"Successfully merged and moved for: {factura}")

    # Optionally print a summary of all missing files at the end
    if missing_files_list:
        print("Summary of missing files:")
        for item in missing_files_list:
            print(item)


df_altas = worksheet_to_df('RAW_ALTA', 'A:AB')
# Assuming df_altas is your DataFrame
df_altas['Factura'] = df_altas['Factura'].astype(str)
# Creating df_altasFacturadas with the specified conditions
df_altasFacturadas = df_altas[(df_altas['Factura'].notna()) & (df_altas['Factura'] != '#N/A') &
                              (df_altas['Folio'].notna()) & (df_altas['Folio'] != '#N/A')].copy()
df_altasFacturadas = df_altasFacturadas[['Factura', 'noAlta', 'noOrden','PISP','Folio']]
print("\n*******************************\nDataframe cargado desde el googleSheet\n*******************************\n")
df_uploadMergelist = df_altasFacturadas[df_altasFacturadas['PISP'] == '#N/A'][['Factura', 'noAlta', 'noOrden', 'PISP','Folio']]
if df_uploadMergelist.empty:
    
    print("\nNo hay nuevos XMLS por subir")
else:
    print(f"\nHay {len(df_uploadMergelist)} xmls por subir, elije la opción 3, abre el script que sube los XMLs y regresa aquí a generar los PDFs.")
 
# Corrected paths
fact_2024_path = 'C:\\Users\\armjorge\\Dropbox\\FACT 2024'
fact_2023_path = 'C:\\Users\\armjorge\\Dropbox\\FACT 2023'
facturasPDF_path = '.\\PDFs'  
facturasXML_path = '.\\XMLs'                    
df_files_pdf = df_altasFacturadas['Factura'] + '.pdf'
df_files_xml = df_altasFacturadas['Factura'] + '.xml'

### List of existing files in the target directory
target_dir_pdfs = r".\PDFs"
target_dir_xmls = r".\XMLs"
raw_existing_files_pdfs = os.listdir(target_dir_pdfs)
raw_existing_files_xmls = os.listdir(target_dir_xmls)

# Suffixes to exclude from the initial list
exclude_suffixes = ('_SAT.pdf', '_TXT.pdf')
# Filter out files ending with specified suffixes
existing_files_pdfs = [file for file in raw_existing_files_pdfs if not file.endswith(exclude_suffixes)]
existing_files_xmls = [file for file in raw_existing_files_xmls if not file.endswith(exclude_suffixes)]

# Check which expected PDFs and XMLs are not in existing_files
filtered_pdf = [pdf for pdf in df_files_pdf if pdf not in existing_files_pdfs]
filtered_xml = [xml for xml in df_files_xml if xml not in existing_files_xmls]

# Print missing files
print("\n*******************************\nArchivos por copiar")
print("Missing PDF files:", filtered_pdf)
print("Missing XML files:", filtered_xml)
print("\n*******************************\n")
source_files_dir = r"C:\Users\armjorge\Dropbox\FACT 2024"
search_and_copy(source_files_dir, target_dir_pdfs, filtered_pdf)
search_and_copy(source_files_dir, target_dir_xmls, filtered_xml)

# Perform the operations


def main_menu():
    while True:
        print("\n1. Genera lsita de XMLS's a subir")
        print("2. Fusiona acuse + factura")
        print("3. Copia a la carpeta de tycsa")
        print("4. Exit")
        choice = input("Enter your choice: ")

        if choice == '1':
            generate_upload_list()
        elif choice == '2':
            merge_altas_teoricas(df_altasFacturadas)
        elif choice == '3':
            print("\n*******************************\nCopiando PDFs con factura y comprobante\n*******************************\n")
            source_merged_folder = r".\Facturas y acuses"  # Replace with the path to target_dir_xmls
            target_tycsa_folder = r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Transición TYCSA\IMSS\Facturas"
            copy_files(source_merged_folder, target_tycsa_folder)
            check_directory_sync(source_merged_folder, target_tycsa_folder)            
            print("\n*******************************\nMoviendo XMLSs\n*******************************\n")
            source_dir = r".\XMLs"  # Replace with the path to target_dir_xmls
            target_dir = r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Transición TYCSA\IMSS\XMLs"
            copy_files(source_dir, target_dir)
            check_directory_sync(source_dir, target_dir)            
        elif choice == '4':
            print("Exiting program.")
            break
        else:
            print("Opción no enlistada.")

if __name__ == "__main__":
    main_menu()