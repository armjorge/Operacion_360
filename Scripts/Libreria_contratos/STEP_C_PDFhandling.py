
from folders_files_open import open_folder, create_directory_if_not_exists, sanitize_filename, trim_to_limit, open_pdf

import os
import pyperclip
import PyPDF2
import re
import shutil
import pandas as pd



def STEP_C_PDF_HANDLING(temp_path, valid_dict, carpeta_contratos):
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
    if os.path.exists(temp_path):
        shutil.rmtree(temp_path)

    create_directory_if_not_exists(temp_path)
    open_folder(temp_path)  # Opens the temp directory for user review
    input("\n🔄 Por favor presiona Enter cuando hayas movido el PDF a la carpeta temporal que se abrió.\n")

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

    # Build the filename using dictionary values ##
    # Sanitize and build the filename
    contrato_sanitized = sanitize_filename(valid_dict.get('Contrato', ''))
    estatus = sanitize_filename(valid_dict.get('Estatus', ''))
    pdf_name = f"{contrato_sanitized}_{estatus}.pdf"
    pdf_name = trim_to_limit(pdf_name)

    # Locate the first (and only) PDF in the temp_path
    pdf_files = [f for f in os.listdir(temp_path) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print("❌ No PDF files found in the directory.")
        return

    temp_pdf_path = os.path.join(temp_path, pdf_files[0])
    pdf_path = os.path.join(temp_path, pdf_name)

    # Rename the file
    try:
        os.rename(temp_pdf_path, pdf_path)
        print(f"✅ Archivo renombrado a: \t{os.path.basename(pdf_path)}")
    except Exception as e:
        print(f"❌ Error al renombrar el archivo: {e}")
        return
    # After renaming the file in STEP_C_PDF_HANDLING
    print(f"🐞 Archivo PDF procesado: {pdf_path}")
    print(f"🐞 Tipo de pdf_path: {type(pdf_path)}")
    pdf_path = [pdf_path]
    STEP_C_write_label_to_PDF(pdf_path, temp_path, carpeta_contratos, valid_dict)

    return pdf_path

def normalize_text(text):
    """Normalize text by removing extra spaces, line breaks, and standardizing quotes (without lowercasing)."""
    if not text:
        return ""
    
    return (
        text.replace('\n', ' ')  # Remove line breaks
        .replace('\r', ' ')
        .replace("‘", "'").replace("’", "'")  # Normalize single quotes
        .replace("“", '"').replace("”", '"')  # Normalize double quotes
        .replace("  ", " ")  # Remove double spaces
        .strip()  # Trim leading and trailing spaces
    )

def STEP_C_write_label_to_PDF(pdf_path, temp_path, carpeta_contratos, valid_dict):
    """
    Reads labeled PDF files, validates the presence of the copied dictionary inside the document,
    and ensures the label has been correctly added.
    """
    if not pdf_path or not pdf_path[0]:
        print("❌ No se proporcionó un archivo PDF válido.")
        return

    source_path = pdf_path[0]
    filename = os.path.basename(source_path)
    dict_text = str(valid_dict).replace('\n', ' ').strip()

    print("\n📄 **PASO C: COPIAR Y PEGAR ETIQUETA EN PDF**\n")
    print(f"Procesando: {filename}")
    print(f"🔹 **Etiqueta generada (normalizada):**\n{normalize_text(dict_text)}")

    # Copy label to clipboard
    pyperclip.copy(dict_text)
    print("✅ **Etiqueta copiada al portapapeles.**")

    # Open the PDF for editing
    open_pdf(source_path)

    while True:
        input(f"\n✏️ **Por favor, agrega una página en blanco, pega la etiqueta y presiona ENTER cuando termines...**")

        # Extract text from the last page of the updated PDF
        extracted_text = read_last_page_pdf(source_path)
        normalized_extracted_text = normalize_text(extracted_text)

        # Debug: Print the normalized extracted text
        print(f"\n🐞 **Texto extraído (normalizado):**\n{normalized_extracted_text}\n")

        if extracted_text:
            # Compare normalized texts without changing capitalization
            if normalize_text(dict_text) in normalized_extracted_text:
                print("✅ **Etiqueta detectada en el archivo PDF.**")
                print("📦 Moviendo archivo a la biblioteca de contratos...")

                # Move the file to the destination directory
                try:
                    destination_file_path = os.path.join(carpeta_contratos, filename)
                    shutil.move(source_path, destination_file_path)
                    print(f"✅ Archivo movido a: {destination_file_path}")

                    # Confirm the file exists in the destination directory
                    if os.path.exists(destination_file_path):
                        print("✅ Confirmación: El archivo está en la biblioteca.")

                        # Remove temp directory
                        if os.path.exists(temp_path):
                            shutil.rmtree(temp_path)
                            print("✅ Directorio temporal eliminado.")
                    
                    print("📄 Archivo correctamente rotulado y almacenado.")
                    break
                except Exception as e:
                    print(f"❌ Error al mover el archivo: {e}")
                    break
            else:
                print("⚠️ **Etiqueta NO encontrada en la última página. Intenta nuevamente.**")
        else:
            print("⚠️ **No se encontró texto válido en la última página.** Intenta nuevamente.")


def read_last_page_pdf(file_path):
    """
    Reads only the last page of a PDF and extracts lines containing `{` or `}`.

    Parameters:
        file_path (str): Path to the PDF file.

    Returns:
        str: Cleaned text containing curly braces `{}`.
    """
    try:
        print(f"🐞 Intentando abrir el archivo PDF en la ruta: {file_path}")
        if not os.path.isfile(file_path):
            print(f"❌ Error: {file_path} no es un archivo válido.")
            return None

        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            num_pages = len(pdf_reader.pages)

            if num_pages == 0:
                print("⚠️ El PDF no contiene páginas.")
                return None

            print(f"🐞 El PDF tiene {num_pages} páginas. Leyendo solo la última...")

            # Read only the last page
            last_page = pdf_reader.pages[-1]
            page_text = last_page.extract_text()

            if page_text:
                # Flatten line breaks and clean up spaces
                cleaned_text = " ".join(page_text.splitlines()).strip()
                return cleaned_text
            else:
                print("⚠️ No se extrajo texto de la última página.")
                return None

    except Exception as e:
        print(f"❌ Error al leer el PDF: {e}")
        return None
    
def extract_text_between_braces(text):
    """Extracts and cleans text enclosed in curly braces `{}` from a given string."""
    # Remove line breaks and clean extra spaces
    cleaned_text = text.replace('\n', ' ').replace('\r', ' ').strip()
    matches = re.findall(r'\{.*?\}', cleaned_text, re.DOTALL)
    return " ".join(matches) if matches else ""

def extract_dict_from_text(text):
    """Extracts dictionary-like content from the given text."""
    try:
        match = re.search(r'\{.*\}', text)
        if match:
            extracted_dict = eval(match.group())  # Convert string to dictionary (ensure source is trusted)
            return extracted_dict if isinstance(extracted_dict, dict) else None
        return None
    except Exception as e:
        print(f"❌ Error al extraer el diccionario: {e}")
        return None

def STEP_C_read_PDF_from_source(carpeta_contratos):
    """Reads PDFs from the specified folder, extracts dictionaries, and populates a DataFrame."""
    pdf_files = [f for f in os.listdir(carpeta_contratos) if f.endswith('.pdf')]
    extracted_data = []

    if not pdf_files:
        print("⚠️ No se encontraron archivos PDF en la carpeta de contratos.")
        return None

    for pdf_file in pdf_files:
        pdf_path = os.path.join(carpeta_contratos, pdf_file)
        print(f"\n📖 Procesando: {pdf_file}")

        # Read the last page and extract text
        extracted_text = read_last_page_pdf(pdf_path)

        if extracted_text:
            extracted_dict = extract_dict_from_text(extracted_text)

            if extracted_dict:
                extracted_data.append(extracted_dict)
                print(f"✅ Datos extraídos de {pdf_file}: {extracted_dict}")
            else:
                print(f"⚠️ No se pudo extraer un diccionario válido de {pdf_file}.")
        else:
            print(f"❌ No se pudo leer la última página de {pdf_file}.")

    # Create DataFrame if data was extracted
    if extracted_data:
        # Get unique keys across all dictionaries
        all_keys = {key for d in extracted_data for key in d.keys()}

        # Create DataFrame
        df_extracted = pd.DataFrame([{key: d.get(key, None) for key in all_keys} for d in extracted_data])

        # Display DataFrame to user
        
        #tools.display_dataframe_to_user("Extracted PDF Data", df_extracted)
        print(df_extracted.head())
    else:
        print("❌ No se extrajo ningún dato de los archivos PDF.")

