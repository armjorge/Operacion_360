import re
import os
import pandas as pd
import PyPDF2
import glob
import shutil
from pathlib import Path

print("\n*******************************\n Extractor de datos de remisiones ***2.0 **** \n*******************************\n")


def inputfiles(source_folder):
    print("Starting file listing process...")

    # Ensure the folder exists
    if not os.path.exists(source_folder):
        raise FileNotFoundError(f"The folder {source_folder} does not exist.")
    
    # List all files in the directory and filter only .pdf files (case-insensitive)
    pdf_files = [
        filename for filename in os.listdir(source_folder)
        if filename.lower().endswith('.pdf')
    ]

    print(f"Found {len(pdf_files)} PDF files.")
    return pdf_files


def extract_info_from_pdf(files, source):
    print("Starting text extraction process...")

    # Define the workload CSV path
    csv_workload_path = os.path.join(source, "extracted.csv")

    # Load or create the workload DataFrame
    if os.path.exists(csv_workload_path):
        csv_workload = pd.read_csv(csv_workload_path, usecols=['filename', 'text'])
        print(f"Loaded existing workload with {len(csv_workload)} files.")
    else:
        csv_workload = pd.DataFrame(columns=['filename', 'text'])
        print("No previous workload found. Creating a new one.")

    # Identify processed filenames
    processed_filenames = set(csv_workload['filename'].str.strip())  # Ensure no trailing spaces
    print(f"Processed filenames: {len(processed_filenames)} files found.")
    print(f"Sample processed filenames: {list(processed_filenames)[:5]}")  # Log a sample for inspection

    # Debugging: Log input files
    input_files_base = [os.path.basename(file).strip() for file in files]  # Normalize input filenames
    print(f"Input filenames: {len(input_files_base)} files found.")
    print(f"Sample input filenames: {input_files_base[:5]}")

    # Filter out files already processed
    unprocessed_files = [file for file in files if os.path.basename(file).strip() not in processed_filenames]
    print(f"Unprocessed filenames: {len(unprocessed_files)} files found.")
    print(f"Sample unprocessed filenames: {unprocessed_files[:5]}")

    if not unprocessed_files:
        print("All files have already been processed. Nothing to do.")
        return csv_workload

    # Ask user for action
    proceed = input(f"Found {len(processed_filenames)} previously processed files. "
                    "Do we skip them or read them again? (Type 'skip' or 'read'): ").strip().lower()

    if proceed == 'read':
        print("Re-reading all files.")
        files_to_process = files  # Re-process all files
        csv_workload = pd.DataFrame(columns=['filename', 'text'])  # Clear the existing workload
    elif proceed == 'skip':
        print("Skipping already processed files.")
        files_to_process = unprocessed_files  # Only process unprocessed files
    else:
        print("Invalid input. Defaulting to 'skip'.")
        files_to_process = unprocessed_files

    # Process the selected files
    for file in files_to_process:
        file_path = os.path.join(source, file)
        print(f"Processing file: {file_path}")

        try:
            with open(file_path, 'rb') as pdf_file_obj:
                pdf_reader = PyPDF2.PdfReader(pdf_file_obj)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text() or ""  # Avoid issues with NoneType
                
                # Add the extracted data to the workload DataFrame
                new_row = {'filename': os.path.basename(file).strip(), 'text': text}
                csv_workload = pd.concat([csv_workload, pd.DataFrame([new_row])], ignore_index=True)
        except Exception as e:
            print(f"Error processing file {file}: {e}")

    # Save the updated workload DataFrame
    csv_workload.to_csv(csv_workload_path, index=False)
    print(f"Updated workload saved to {csv_workload_path} with {len(csv_workload)} files.")

    return csv_workload



def process_extracted_text(df_input, source_folder):
    print("Iniciando el procesamiento del texto extraído")
    # Create the DataFrame with the required columns
    df_analyzed = pd.DataFrame(columns=[
        'filename', 'Oremision', 'Osuministro', 'Contrato', 
        'Fecha emisión', 'FEEN', 'Clave', 'Tabla_lotes'
    ])

    # Process each row in the input DataFrame
    for index, row in df_input.iterrows():
        text = row['text']  # Extract the text content
        filename = row['filename']  # Extract the filename

        # Apply the regex patterns to extract information
        num_orden_remision = re.search(r'NÚMERO DE ORDEN DE SUMINISTRO:\n(.*?)\n', text, re.DOTALL)
        contrato = re.search(r'(LA-E115-2022-MED-INSABI-\d+-\d+/\d+)', text)
        fecha_expedicion = re.search(r'Fecha expedición de la orden: (\d+/\d+/\d+)', text)
        fecha_entrega = re.search(r'Fecha de entrega: (\d+/\d+/\d+ \d+:\d+)', text)
        pattern = re.search(r'(\d{3}\.\d{3}\.\d{4}\.\d{2})', text)
        tabla_lotes = re.search(r'ENTREGARALTO ANCHO PROFUNDIDAD(.*?)ORDEN DE REMISIÓN', text, re.DOTALL)

        # Split O Suministro and O Remisión if found
        if num_orden_remision:
            num_parts = num_orden_remision.group(1).strip().split(" ", 1)
            oremision = num_parts[0].strip() if len(num_parts) > 0 else ""
            osuministro = num_parts[1].strip() if len(num_parts) > 1 else ""
        else:
            oremision = ""
            osuministro = ""

        # Add data to the analyzed DataFrame
        df_analyzed = pd.concat([df_analyzed, pd.DataFrame([{
            'filename': filename,
            'Oremision': oremision,
            'Osuministro': osuministro,
            'Contrato': contrato.group(1) if contrato else "",
            'Fecha emisión': fecha_expedicion.group(1) if fecha_expedicion else "",
            'FEEN': fecha_entrega.group(1) if fecha_entrega else "",
            'Clave': pattern.group(1) if pattern else "",
            'Tabla_lotes': tabla_lotes.group(1).strip() if tabla_lotes else ""
        }])], ignore_index=True)

    # Check for duplicate rows based on 'Osuministro' and 'Oremision'
    duplicates = df_analyzed.duplicated(subset=['Osuministro', 'Oremision'], keep=False)

    # Separate unique rows and those with duplicates
    unique_rows = df_analyzed[~duplicates]
    duplicate_rows = df_analyzed[duplicates]

    # For duplicate rows, keep only one representative per group (same Osuministro and Oremision)
    if not duplicate_rows.empty:
        representative_duplicates = duplicate_rows.drop_duplicates(subset=['Osuministro', 'Oremision'], keep='first')
        df_analyzed = pd.concat([unique_rows, representative_duplicates], ignore_index=True)
        print(f"Filtered duplicates: {len(duplicate_rows) - len(representative_duplicates)} duplicate rows removed.")
    else:
        print("No duplicates found.")

    return df_analyzed




def optimization_workload(df_input, source):
    # Check for duplicated rows based on 'Oremision' and 'Osuministro'
    duplicates = df_input[df_input.duplicated(subset=['Oremision', 'Osuministro'], keep=False)]
    duplicate_count = len(duplicates)

    print(f"We found {duplicate_count} duplicated files. Do we delete them? (Type 'yes' or 'no'):")
    user_input = input().strip().lower()

    if user_input == 'yes' and duplicate_count > 0:
        # Identify duplicate groups
        for osuministro, oremision in duplicates[['Osuministro', 'Oremision']].drop_duplicates().values:
            # Get all duplicate rows for the current group
            duplicate_files = df_input[
                (df_input['Osuministro'] == osuministro) &
                (df_input['Oremision'] == oremision)
            ]

            # Keep only the first occurrence, delete the rest
            to_delete = duplicate_files.iloc[1:]['filename']  # All but the first
            for filename in to_delete:
                file_path = os.path.join(source, filename)
                if os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                        print(f"Deleted file: {file_path}")
                    except Exception as e:
                        print(f"Error deleting file {file_path}: {e}")

        # Update the DataFrame to remove duplicates, keeping only the first occurrence
        df_input = df_input.drop_duplicates(subset=['Oremision', 'Osuministro'], keep='first')
        print(f"Deleted {duplicate_count} duplicated rows and their corresponding files.")
    else:
        print("No duplicates were deleted.")

    return df_input

def workingdf(df_notduplicated):
    # Check if the required columns exist
    required_columns = ['Oremision', 'Osuministro']
    if not all(col in df_notduplicated.columns for col in required_columns):
        print("Error: One or more required columns are missing in the DataFrame.")
        return None

    # Process the DataFrame
    df_processed = (
        df_notduplicated
        .sort_values(by=['Osuministro', 'Oremision'], ascending=[True, False])  # Sort by Osuministro, then by Oremision descending
        .drop_duplicates(subset=['Osuministro'], keep='first')  # Keep the row with the greatest Oremision per Osuministro
    )
    
    # Define the output path
    output_path = ".\\INSABI_Remisiones 2024\\extraccionRemisionesINSABI.csv"
    df_processed.to_csv(output_path, index=False)
    
    # Print success message
    print(f"DataFrame processed and saved successfully to {output_path}.")
    
    return df_processed


def main():
    source = ".\\INSABI_Remisiones 2024\\Descargadas"
    files = inputfiles(source)  # List of normalized PDF files

    # Extract text and get the master DataFrame
    df_master = extract_info_from_pdf(files, source)
    df_extraction_brut = process_extracted_text(df_master, source)
    
    df_notduplicated = optimization_workload(df_extraction_brut, source)
    #print(df_notduplicated.head)
    df_working = workingdf(df_notduplicated)
    print("\n*****\nworking dataframe was generated, keep moving\n *******")
    output_csv_path = '.\\Camunda\\INSABI_Remisiones 2024\\extraccionRemisionesINSABI_brut.csv'

    # Ensure the directory exists and save the CSV
    output_path = ".\\INSABI_Remisiones 2024\\extraccionRemisionesINSABI_brut.csv"
    df_notduplicated.to_csv(output_path, index=False)
    print(f"Saved deduplicated data to {output_csv_path}")

    
if __name__ == "__main__":
    main()
    
