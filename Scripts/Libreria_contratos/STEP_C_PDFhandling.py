
from folders_files_open import open_folder, create_directory_if_not_exists, sanitize_filename, trim_to_limit, open_pdf

import os
import pyperclip
import PyPDF2
import re
import shutil
import pandas as pd



def STEP_C_PDF_HANDLING(valid_dict, carpeta_contratos):
    """
    Handles PDF loading, renaming, and opening based on extracted data.
    
    Parameters:
        carpeta_contratos (str): s.path.join(folder_root, "Implementación", "Contratos", f"{procedimiento}")
        valid_dict (dict): Dictionary containing metadata fields for naming.
    Returns:
        None
    """
    print("\n📄 Con los datos extraídos, es momento de pasar a cargar el PDF\n")

    temp_path = os.path.join(carpeta_contratos, 'Temp')
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
            pdf_file = pdf_files[0]
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

    # Pull el filename del diccionario que pasamos. 
    pdf_name = valid_dict['Nombre del archivo']
    temp_pdf_path = os.path.join(temp_path, pdf_file)
    pdf_path = os.path.join(temp_path, pdf_name) # Así se debe renombrar. 

    # Rename the file
    try:
        os.rename(temp_pdf_path, pdf_path)
        print(f"✅ Archivo {os.path.basename(temp_pdf_path)} -> renombrado a: {os.path.basename(pdf_path)}")
    except Exception as e:
        print(f"❌ Error al renombrar el archivo: {e}")
        return
    # After renaming the file in STEP_C_PDF_HANDLING
    #print(f"🐞 Archivo PDF procesado: {pdf_path}")
    #print(f"🐞 Tipo de pdf_path: {type(pdf_path)}")
    #pdf_path = [pdf_path]
    STEP_C_write_label_to_PDF(pdf_path, carpeta_contratos, valid_dict, just_read=False)

def STEP_C_write_label_to_PDF(pdf_path, carpeta_contratos, valid_dict, just_read=False):
    """
    Reads labeled PDF files, validates the presence of the copied dictionary inside the document,
    and ensures the label has been correctly added.
    """
    if not pdf_path or not pdf_path[0]:
        print("❌ No se proporcionó un archivo PDF válido.")
        return
    filename = valid_dict['Nombre del archivo']
    dict_text = str(valid_dict).replace('\n', ' ').strip()
    dict_text = normalize_text(dict_text)
    print("\n📄 **PASO C: COPIAR Y PEGAR ETIQUETA EN PDF**\n")
    print(f"Procesando: {filename}")
    print(f"🔹 **Etiqueta generada (normalizada):**\n{dict_text}")

    # Copy label to clipboard
    pyperclip.copy(dict_text)
    print("✅ **Etiqueta copiada al portapapeles.**")

    # Open the PDF for editing
    
    print("✅ **Abriendo el PDF.**")
    input(f"\n✏️ **Cuando estés listo presiona enter, se abrirá acrobat, agrega una página en blanco al final, pega la etiqueta y cierra el acrobat.**")
    open_pdf(pdf_path)

    while True:
        # Extract text from the last page of the updated PDF
        extracted_text = read_last_page_pdf(pdf_path)
        normalized_extracted_text = normalize_text(extracted_text)
        # Debug: Print the normalized extracted text
        print(f"\n🐞 **Texto extraído (normalizado):**\n{normalized_extracted_text}\n")

        if extracted_text:
            # Compare normalized texts without changing capitalization
            if just_read is True:
                return extracted_text            
            if dict_text in normalized_extracted_text:
                print("✅ **Etiqueta detectada en el archivo PDF.**")
                print("📦 Moviendo archivo a la biblioteca de contratos...")

                # Move the file to the destination directory
                try:
                    shutil.move(pdf_path, carpeta_contratos)
                    print(f"✅ Archivo movido a: {carpeta_contratos}")

                    # Confirm the file exists in the destination directory
                    if os.path.exists(carpeta_contratos):
                        print("✅ Confirmación: El archivo está en la biblioteca.")

                        # Remove temp directory
                        temp_path = os.path.dirname(pdf_path)  
                        if os.path.exists(temp_path):
                            shutil.rmtree(temp_path)
                            print(f"✅ Directorio temporal {os.path.basename(temp_path)} eliminado.")
                    
                    print("\n✅ Archivo correctamente rotulado y almacenado.\n")
                    print(f"{'*'*5}\n✅ Captura del contrato {valid_dict['Contrato']} - Institución {valid_dict['Institución']} finalizada.\n{'*'*5}")
                    break
                except Exception as e:
                    print(f"❌ Error al mover el archivo: {e}, ciérralo y vuelve a intentar")
                    input("Presiona ENTER para intentar moverlo de nuevo")
                    continue
            else:
                print("⚠️ **El texto de la última página no coincide. Intenta nuevamente.**")
                input("Presiona ENTER para leer la última página de nuevo")
                continue                
        else:
            print("⚠️ **No se encontró texto válido en la última página.** Intenta nuevamente.")
            input("Presiona ENTER para leer la última página de nuevo")
            continue 

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
