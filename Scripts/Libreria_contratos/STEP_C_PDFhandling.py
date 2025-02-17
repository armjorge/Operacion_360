
from folders_files_open import open_folder, create_directory_if_not_exists, open_pdf

import os
import pyperclip
import PyPDF2
import re

def STEP_C_PDF_HANDLING(temp_path, valid_dict):
    """
    Handles PDF loading, renaming, and opening based on extracted data.
    
    Parameters:
        temp_path (str): The path where the temporary PDF is stored.
        valid_dict (dict): Dictionary containing metadata fields for naming.

    Returns:
        None
    """
    print("\n📄 Con los datos extraídos, es momento de pasar a cargar el PDF\n")

    # Ensure the temp directory exists
    create_directory_if_not_exists(temp_path)
    open_folder(temp_path)  # Opens the temp directory for user review

    while True:
        # List PDF files in the directory
        pdf_files = [f for f in os.listdir(temp_path) if f.endswith('.pdf')]

        if len(pdf_files) == 1:
            break  # Continue processing if exactly one PDF is found

        elif len(pdf_files) > 1:
            print("⚠️ Se detectaron múltiples archivos PDF en la carpeta temporal:")
            for file in pdf_files:
                print(f"   - {file}")

            input("🗑️ Elimina los archivos extra y deja solo uno. Luego presiona ENTER para continuar...")
            continue  # Restart the loop after cleaning

        elif not pdf_files:
            print("❌ No se encontró ningún archivo PDF en la carpeta temporal.")
            input("📂 Agrega un archivo PDF en la carpeta y presiona ENTER para continuar...")
            continue  # Restart the loop

    # Build the filename using dictionary values
    day, month, year = valid_dict['Primer registro'].split('/')

    pdf_name = f"{year} {month} {day} {valid_dict['Materia']}.pdf"

    # Locate the first (and now only) PDF
    temp_pdf_path = os.path.join(temp_path, pdf_files[0])
    pdf_path = os.path.join(temp_path, pdf_name)

    # Rename the file
    try:
        os.rename(temp_pdf_path, pdf_path)
        print(f"✅ Archivo renombrado a: {pdf_path}")
    except Exception as e:
        print(f"❌ Error al renombrar el archivo: {e}")
        return

    # Open the renamed PDF
    open_pdf(pdf_path)
    return pdf_path

def extract_text_between_braces(text):
    """ Extracts and cleans text enclosed in curly braces `{}` from a given string. """
    matches = re.findall(r'\{.*?\}', text, re.DOTALL)
    return " ".join(matches) if matches else ""

def read_pdf(file_name):
    """
    Reads a PDF and extracts all lines that contain `{` or `}`.
    
    Parameters:
        file_name (str): Path to the PDF file.

    Returns:
        str: Cleaned text containing curly braces `{}`.
    """
    with open(file_name, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        
        extracted_lines = []  # Store matching lines

        # Loop through each page and extract text
        for page in pdf_reader.pages:
            page_text = page.extract_text()
            if page_text:
                # Find all lines that contain '{' or '}'
                filtered_lines = [line.strip() for line in page_text.split('\n') if '{' in line or '}' in line]
                extracted_lines.extend(filtered_lines)  # Add them to the list

        # Join all extracted lines into a single string
        cleaned_text = "\n".join(extracted_lines)

        return cleaned_text if cleaned_text else None  # Return None if no valid lines found

def STEP_C_read_labeled_pdf(pdf_list, valid_dict):
    """
    Reads labeled PDF files, validates the presence of the copied dictionary inside the document, 
    and ensures the label has been correctly added.

    Parameters:
        df_procesados (str): Path to store processed PDFs.
        pdf_list (list): List of PDF file paths.
        valid_dict (dict): The dictionary label to be inserted.

    Returns:
        None
    """
    pdf_dict  = {}
    for pdf_file in pdf_list:
        print("\n📄 **PASO C: COPIAR Y PEGAR ETIQUETA EN PDF**\n")
        filename = os.path.basename(pdf_file)  # ✅ Extracts "filename.pdf"
        # Convert the dictionary to a string and copy it to clipboard
        dict_text = str(valid_dict)
        print("🔹 **Etiqueta generada:**")
        print(dict_text)
        pyperclip.copy(dict_text)
        print("✅ **Etiqueta copiada al portapapeles.**")

        while True:
            input(f"\n✏️ **Por favor, pega la etiqueta en el archivo {pdf_file} y presiona ENTER cuando termines...**")

            # Extract text from PDF
            extracted_text = read_pdf(pdf_file)

            if extracted_text:
                # Extract dictionary-like text from the PDF
                extracted_text_cleaned = extract_text_between_braces(extracted_text)

                # Compare extracted text with copied dictionary text
                if dict_text in extracted_text_cleaned:
                    print("✅ **Etiqueta encontrada en el archivo PDF.**")                    
                    pdf_dict[filename] = dict_text
                    sucess_message = "Continúa, el PDF está rotulado y listo para moverse"
                    print(sucess_message)                    
                    break
                else:
                    print("⚠️ **Etiqueta NO encontrada en el archivo.** Intenta nuevamente.")

            else:
                print("⚠️ **No se encontró texto válido en el PDF.** Intenta nuevamente.")
    

    
    return pdf_dict