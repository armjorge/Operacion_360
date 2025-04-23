import os
import pandas as pd
import PyPDF2

def read_filenames_from_excel(folder_path):
    """
    Read filenames from an Excel file within the given folder.

    :param folder_path: Path to the folder containing the Excel file.
    :return: List of filenames starting with 'NC', with the folder portada file added.
    """
    # Define the Excel file path
    excel_path = os.path.join(folder_path, f"{os.path.basename(folder_path)}_audit.xlsx")
    
    # Read the Excel file
    df = pd.read_excel(excel_path)

    # Filter rows where 'Filename' starts with 'NC'
    df = df[df['Filename'].str.startswith('NC', na=False)]

    # Add the portada file at the beginning
    print("Agregando la protada a la lista")
    portada_file = f"{os.path.basename(folder_path)}_portada"
    print(f"Agregando la protada a la lista {portada_file}")
    filenames = [portada_file] + df['Filename'].tolist()

    return filenames

def generate_file_paths(folder_path, filenames):
    """Generate full paths for each file and suffix, check existence."""
    file_paths = []
    missing_files = []
    #suffixes=['', '_SAT', '_TXT']
        # Loop until a valid choice is provided
    valid_choices = {'1', '2', '3', '4'}
    while True:
        print("Choose which suffixes to include:")
        print("1: ['']")
        print("2: ['_SAT']")
        print("3: ['_TXT']")
        print("4: ['', '_SAT', '_TXT'] (default)")
    
        choice = input("Enter your choice (1, 2, 3, or 4): ").strip()
        if choice in valid_choices:
            break
        else:
            print("Invalid choice. Please try again.\n")
    
    # Set suffixes based on the user's choice
    if choice == '1':
        suffixes = ['']
    elif choice == '2':
        suffixes = ['_SAT']
    elif choice == '3':
        suffixes = ['_TXT']
    elif choice == '4':
        suffixes = ['', '_SAT', '_TXT']

    for filename in filenames:
        for suffix in suffixes:
            full_path = os.path.join(folder_path, f"{filename}{suffix}.pdf")
            if os.path.exists(full_path):
                file_paths.append(full_path)
            else:
                if suffix == '':  # Report only if the primary file is missing
                    missing_files.append(full_path)
    return file_paths, missing_files

def merge_pdfs(file_list, output_filename):
    """Merge multiple PDFs into one."""
    merger = PyPDF2.PdfMerger()
    
    for file in file_list:
        with open(file, 'rb') as pdf:
            merger.append(pdf)

    with open(output_filename, 'wb') as output_pdf:
        merger.write(output_pdf)

def PDF_para_imprimir(folder_path): 
    filenames = read_filenames_from_excel(folder_path)
    file_paths, missing_files = generate_file_paths(folder_path, filenames)
    
    if missing_files:
        print("Missing files:")
        for file in missing_files:
            print(file)
    
    if file_paths:
        output_filename = os.path.join(folder_path, os.path.basename(folder_path) + "_merged.pdf")
        merge_pdfs(file_paths, output_filename)
        print(f"Merged PDFs into {output_filename}")
    else:
        print("No files to merge.")
