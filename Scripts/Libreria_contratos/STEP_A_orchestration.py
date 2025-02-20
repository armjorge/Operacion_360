import os
import sys
import shutil
from folders_files_open import create_directory_if_not_exists, load_dataframe
from STEP_B_Dict import STEP_B_get_string_populated
from STEP_C_PDFhandling import STEP_C_PDF_HANDLING
import ast


def STEP_A_orchestration(folder_path, df_desagregada_procedimiento, institucion_column, selected_procedimiento, carpeta_contratos):
    while True:
        tipo = input("Define el tipo P) Primigenio o M) Modificatorio: ")
        tipo = tipo.lower()
        if tipo == 'p':
            print("Tipo seleccionado: Primigenio")
            tipo = 'Primigenio'
            break
        elif tipo == 'm':
            print("Tipo seleccionado: Modificatorio")
            tipo = 'Modificatorio'
            break
        else:
            print("Entrada no válida, por favor elija P o M.")
    print(tipo)
    working_folder = folder_path
    print("\tIniciando el orquestador para poblar el diccionario")
    PDF_library = os.path.join(working_folder, "PDF_Library")
    create_directory_if_not_exists(PDF_library)
    print("\n\tPASO A: GENERAR DE DICCIONARIO PARA NUEVO CONTRATO\n")
    human_dict, valid_dict = STEP_B_get_string_populated(df_desagregada_procedimiento, tipo, institucion_column, selected_procedimiento, folder_path)
    print('Diccionario Capturado\n', human_dict, "\nDiccionario Procesado\n", valid_dict)
    print("\n\tPASO C: MANEJAR EL PDF\n") 
    # Convert the SKU string to a list of dictionaries
    try:
        sku_list = ast.literal_eval(valid_dict['SKU'])
    except (SyntaxError, ValueError) as e:
        print(f"❌ Error al convertir el campo 'SKU' en una lista de diccionarios: {e}")
        sku_list = []

    # Calculate the total
    total = sum(item['Precio'] * item['Máximo'] for item in sku_list)

    # Print the total
    print(f"💰 Total del contrato: {total}")
    temp_directory = os.path.join(working_folder, 'Temp')
    contrato_pdf_temporal = [STEP_C_PDF_HANDLING(temp_directory, valid_dict, carpeta_contratos)]

    #result_dict = STEP_C_read_labeled_pdf(pdf_path_list, valid_dict)
    #print(f"Inserto agregado: {result_dict}")
    #df_feed = STEP_C_feed_DF(df_referencia_interna, result_dict)
