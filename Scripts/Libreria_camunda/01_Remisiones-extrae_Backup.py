import re
import os
import pandas as pd
import PyPDF2
import glob
import shutil
from pathlib import Path

print("\n*******************************\n Extractor de datos de remisiones \n*******************************\n")

"""
    def copy_pdfs_to_temp(source_directory, temp_directory):
        if not os.path.exists(temp_directory):
            os.makedirs(temp_directory)
        for pdf_file in glob.iglob(f'{source_directory}/**/*.pdf', recursive=True):
            shutil.copy(pdf_file, temp_directory)

"""    

"""
    def copy_pdfs_to_temp(source_directory, temp_directory):
        # Create the temp directory if it does not exist
        if not os.path.exists(temp_directory):
            os.makedirs(temp_directory)

        # Pattern to match 'O.S.' in any combination of uppercase and lowercase letters
        pattern = re.compile(r'o\.s\.', re.IGNORECASE)
        
        # Use pathlib to handle case-insensitive file search
        source_path = Path(source_directory)
        
        # Iterate over all files in the source directory and its subdirectories
        for pdf_file in source_path.rglob('*'):
            # Check if the file is a PDF, ignoring case, and does not include 'O.S.' in any case
            if pdf_file.suffix.lower() == '.pdf' and not pattern.search(pdf_file.name):
                # If conditions are met, construct the new destination path without doubling '.pdf' extension
                destination_file_name = pdf_file.stem + '.pdf'  # Use 'stem' to drop original suffix if it's '.pdf'
                destination_file = Path(temp_directory) / destination_file_name
                # Copy the file to the temp directory with the new name
                shutil.copy(pdf_file, destination_file)
"""

def extract_info_from_pdf(file_path):
    pdf_file_obj = open(file_path, 'rb')
    print(pdf_file_obj)
    pdf_reader = PyPDF2.PdfReader(pdf_file_obj)
    
    num_pages = len(pdf_reader.pages)
    text = ""

    for page in range(num_pages):
        page_obj = pdf_reader.pages[page]
        text += page_obj.extract_text()

    pdf_file_obj.close()

    # Extract the text between "ENTREGARALTO ANCHO PROFUNDIDAD" and "ORDEN DE REMISIÓN"
    Tabla_lotes = re.search(r'ENTREGARALTO ANCHO PROFUNDIDAD(.*?)ORDEN DE REMISIÓN', text, re.DOTALL)
    num_orden_remision = re.search(r'NÚMERO DE ORDEN DE SUMINISTRO:\n(.*?)\n', text, re.DOTALL)
    contrato = re.search(r'(LA-E115-2022-MED-INSABI-\d+-\d+/\d+)', text)
    fecha_expedicion = re.search(r'Fecha expedición de la orden: (\d+/\d+/\d+)', text)
    fecha_entrega = re.search(r'Fecha de entrega: (\d+/\d+/\d+ \d+:\d+)', text)
    pattern = re.search(r'(\d{3}\.\d{3}\.\d{4}\.\d{2})', text)
    #lote = re.search(r'ENTREGARALTO ANCHO PROFUNDIDAD\n(.*?)(?=\n\n)', text, re.DOTALL)
    #cantidad= re.search(r'ENTREGARALTO ANCHO PROFUNDIDAD((?:.*\n)*?)\n\n', text)

    return [file_path, 
            num_orden_remision.group(1) if num_orden_remision else "",
            contrato.group(1) if contrato else "",
            fecha_expedicion.group(1) if fecha_expedicion else "",
            fecha_entrega.group(1) if fecha_entrega else "",
            pattern.group(1) if pattern else "",
            #lote.group(1).strip() if lote else "",
            #cantidad.group(1) if cantidad else"",
            Tabla_lotes.group(1).strip() if Tabla_lotes else ""]

def delete_pdfs_from_temp(temp_directory):
    for pdf_file in glob.glob(f'{temp_directory}/*.pdf'):
        os.remove(pdf_file)

def clean_pdf_extensions(source_directory):
    # List all files in the directory
    files = os.listdir(source_directory)
    
    for file_name in files:
        # Check if the file ends with any variation of .pdf extension
        if file_name.lower().endswith('.pdf'):
            # Remove all occurrences of .pdf and add it back once
            base_name = file_name.lower().replace('.pdf', '')
            # Construct the new file name with a single .pdf extension
            new_file_name = base_name + '.pdf'
            # Build the full path for old and new file names
            print(f"Before Renaming: {file_name}, After Renaming: {new_file_name}")

            """
            old_file_path = os.path.join(source_directory, file_name)
            new_file_path = os.path.join(source_directory, new_file_name)
            # Rename the file only if the new name is different
            if old_file_path != new_file_path:
                os.rename(old_file_path, new_file_path)
                print(f'Renamed: {old_file_path} -> {new_file_path}')
            """
            
def main():
    #temp_directory = ".\\INSABI_Remisiones 2024\\Temporales"
    #source_directory = "C:\\Users\\armjorge\\Dropbox\\REMISIONES\\02 IMSS BIENESTAR"
    source_directory2 = ".\\INSABI_Remisiones 2024\\Descargadas"

    # Update directory variable to temp directory
    directory = source_directory2
    clean_pdf_extensions(directory)

    files = [f for f in os.listdir(directory) if f.endswith('.pdf')]
    data = []
    for file in files:
        extracted_data = extract_info_from_pdf(os.path.join(directory, file))
        data.append(extracted_data)
        
        print(f"File: {file}")
        print("Tabla_lotes:", extracted_data[-1])
        print("------")

    df = pd.DataFrame(data, columns=["Archivo", "O Suministro y Remision", "Contrato", "Fecha emisión", "FEEN", "Clave", "Tabla_lotes"])
    
    # Initialize new columns
    df['Oremision'] = None
    df['Osuministro'] = None
    
    # Log file for skipped rows
    log_file = './INSABI_Remisiones 2024/skipped_rows.log'
    skipped_rows = []

    for index, row in df.iterrows():
        try:
            # Attempt to split and assign
            split_values = row['O Suministro y Remision'].split(" ")
            if len(split_values) == 2:
                df.at[index, 'Oremision'] = split_values[0]
                df.at[index, 'Osuministro'] = split_values[1]
            else:
                raise ValueError("Incorrect split length")
        except Exception as e:
            # Log skipped rows
            skipped_rows.append((index, row['O Suministro y Remision'], str(e)))

    # Write skipped rows to the log file
    with open(log_file, 'w') as log:
        for row in skipped_rows:
            log.write(f"Row Index: {row[0]}, Value: {row[1]}, Error: {row[2]}\n")

    print(f"Skipped rows logged to {log_file}")

    # Reorder columns
    
    # Find duplicates based on 'Osuministro' column
    duplicates = df[df.duplicated('Osuministro', keep=False)]

    # Sort the DataFrame based on 'Osuministro' to group duplicates together
    duplicates_sorted = duplicates.sort_values(by='Osuministro')

    # Group by 'Osuministro' and filter out groups with all identical 'Oremision'
    def filter_different_oremision(group):
        if group['Oremision'].nunique() > 1:
            return True
        return False
    # Apply filter
    filtered_groups = duplicates_sorted.groupby('Osuministro').filter(filter_different_oremision)

    # If you want to print the DataFrame in a more terminal-friendly format, use:
    print("\n*******************************\nHay más de una remisión para las siguientes Órdenes: ")
    print(filtered_groups[['Oremision', 'Osuministro']].to_string(index=False))
    print("\n*******************************\n")
    # Esta parte del código quita las Osuministro duplicadas. 
    for osuministro in filtered_groups['Osuministro'].unique():
        # Get all rows for this 'Osuministro'
        rows = df[df['Osuministro'] == osuministro]
        
        # Find the max 'Oremision' value among these rows
        max_oremision = rows['Oremision'].max()
        
        # Drop rows from 'df' where 'Osuministro' matches but 'Oremision' is not the max
        df = df.drop(rows[rows['Oremision'] < max_oremision].index)
        
    print("\n*******************************\n Para las órdenes duplicadas con diferente remisión, nos quedamos la remisión mayor del grupo.  \n*******************************\n")
    #df.to_csv('extraccionRemisionesINSABI.csv', index=False)
    directory_path = './INSABI_Remisiones 2024'

    # Check if the directory exists, if not, create it
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)

    # Define the full path for the CSV file
    file_path = os.path.join(directory_path, 'extraccionRemisionesINSABI.csv')
    print(f"\n*******************************\n Archivo CSV salvado a {file_path}  \n*******************************\n")
    # Save the DataFrame to CSV without the index
    df.to_csv(file_path, index=False)
    
    # Delete PDFs from temporary directory after processing
    #delete_pdfs_from_temp(temp_directory)

if __name__ == "__main__":
    main()
