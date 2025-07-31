
from STEP_B_Dict import STEP_B_get_string_populated
from STEP_C_PDFhandling import STEP_C_PDF_HANDLING
import pandas as pd
import os
import sys

def STEP_A_orchestration(df_desagregada_procedimiento, selected_procedimiento, carpeta_contratos):
    # Bucle para definir el tipo de instrumento (Primigenio o Modificatorio)
    print("\tPaso A: Definir si vamos a capturar un primigenio o modificatorio")
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
            continue
    
    print(f"Se generó la variable 'tipo' con el valor elegido: {tipo}")
    
    print(f"\n{'*' * 5}\nIniciando el orquestador para poblar el diccionario\n{'*' * 5}\n")
    print("\nPASO B: GENERAR DE DICCIONARIO PARA NUEVO CONTRATO\n")

    # Enviamos a generar un diccionario. 

    institucion_column = 'Institución'
    print(f"Conforme a la funció load_pickles(desagregadas_folder), esperamos que la desagregada incluya la columna: {institucion_column}")
    computer_dict, df_contratos_convenios = STEP_B_get_string_populated(df_desagregada_procedimiento, tipo, institucion_column, selected_procedimiento, carpeta_contratos)
    #print("\nDiccionario Procesado:\n", computer_dict)
    
    step_C = "Paso C: Pegar el diccionario generado en el PDF"
    step_c_highlight= f"{'*' * len(step_C)}"
    print(f"{step_c_highlight}\n{step_C}\n{step_c_highlight}")

    # Pegar el diccionario en el PDF y archivar el PDF.
    while True: 
        user_input = input("Tenemos diccionario creado ¿procedemos a pegarlo al PDF? si o no").strip().lower()
        if user_input == "si":        
            STEP_C_PDF_HANDLING(computer_dict, carpeta_contratos)
            break
        else:
            print("❌ Consigue el PDF y vuelve para pegarle el diccionario.")
            continue
    
    # Actualizar el dataframe df_contratos_convenios

    mask = pd.Series(True, index=df_contratos_convenios.index)
    for col, val in computer_dict.items():
        # Si alguna columna no existiera, esto levantaría KeyError.
        # Pero asumimos que todas las llaves de computer_dict están en df_contratos_convenios.columns.
        mask &= (df_contratos_convenios[col] == val)

    already_exists = mask.any()

    if not already_exists:
        # (4) Crear un DataFrame de una sola fila a partir de computer_dict
        new_row = pd.DataFrame([computer_dict])

        # (5) Concatenar esa fila con el DataFrame original
        df_contratos_convenios = pd.concat(
            [df_contratos_convenios, new_row],
            ignore_index=True
        )

        # (6) Sobrescribir el pickle
        contratos_convenios = os.path.join(carpeta_contratos, f"{selected_procedimiento}.pickle")
        df_contratos_convenios.to_pickle(contratos_convenios)
        print("Se agregó la fila y se guardó en el pickle.")
    else:
        print("Esa fila ya existe en df_contratos_convenios; no se hace nada.")
    #result_dict = read_last_page_pdf(pdf_path_list, valid_dict)
    #print(f"Inserto agregado: {result_dict}")
    #df_feed = STEP_C_feed_DF(df_referencia_interna, result_dict)

if __name__ == "__main__":
    script_folder_root = os.getcwd()
    print('\nLibrería script \n', script_folder_root)
    libreria_folder = os.path.abspath(os.path.join(script_folder_root, "..")) 
    print('\nLibrerías de scripts \n', libreria_folder)
    # Ensure the script folder is added to sys.path
    if libreria_folder not in sys.path:
        sys.path.append(libreria_folder)
    contratos_library_scripts = os.path.join(libreria_folder, "Libreria_contratos")
    if contratos_library_scripts not in sys.path:
        sys.path.append(contratos_library_scripts)
    comunes_library_scripts = os.path.join(libreria_folder, "Libreria_comunes")
    if comunes_library_scripts not in sys.path:
        sys.path.append(comunes_library_scripts)
    working_path = os.path.abspath(os.path.join(script_folder_root, "..", "..")) 
    print("\nPath de trabajo\n", working_path)
    # Cargar librerías internas
    from folders_files_open import create_directory_if_not_exists
    #from STEP_C_PDFhandling import STEP_C_read_PDF_from_source
    from dataframes_generation import create_dataframe #(extension, dataframe_name, columns, output_folder)
    # Generador de carpetas
    from administracion_de_contratos import administracion_de_contratos
    from print_columns import print_columns
    #Cargar los folders requeridos. 
    working_folder = desagregadas_folder = os.path.join(working_path, "Implementación")
    desagregadas_folder = os.path.join(working_path, "Implementación", "Desagregadas")
    create_directory_if_not_exists(working_folder)
    create_directory_if_not_exists(desagregadas_folder)  
    carpeta_contratos = os.path.join(working_path, "Implementación", "Contratos")
    from desagregada_to_pickle import load_pickles
    df_proccedure, procedimiento = load_pickles(desagregadas_folder)
    df_proccedure.info()
    print(df_proccedure.sample(4))    
    from desagregada_to_pickle import base_contratos_convenios_pickle
    Columnas_captura = ['Institución', 'Procedimiento', 'Contrato', 'Fecha Inicio', 'Fecha Fin', 'Productos y precio', 'Total', 'Nombre del archivo', 'Estatus', 'Convenio modificatorio', 'Objeto del convenio', 'Tipo']
    Folder_procedimiento = os.path.join(working_path, "Implementación", "Contratos", f"{procedimiento}")
    create_directory_if_not_exists(Folder_procedimiento)
    base_contratos_convenios = base_contratos_convenios_pickle(Folder_procedimiento, procedimiento, Columnas_captura) 
    print(f"\nprocedimiento:{procedimiento}")
    while True:
        print("¿Qué paso quieres realizar?")
        print("A) Capturar procedimiento")
        print("B) Exportar base de datos")
        print("C) Eliminar registros de la tabla")
        print("D) Generar el archivo de copias físicas")
        print("E) Actualizar copia no formalizada")
        print("salir")
        
        choice = input("Selecciona una opción (A/B/C): ").strip().upper()

        if choice == "A":
            from STEP_A_orchestration import STEP_A_orchestration
            STEP_A_orchestration(df_proccedure, procedimiento, Folder_procedimiento)
        elif choice == "B":
            contratos_base_pickle = os.path.join(Folder_procedimiento, f"{procedimiento}.pickle")
            # Definir la ruta base del pickle
            contratos_base_pickle = os.path.join(Folder_procedimiento, f"{procedimiento}.pickle")

            def export_to_excel(pickle_path):
                # 1. Cargar el DataFrame desde el pickle
                df = pd.read_pickle(pickle_path)
                
                # 2. Construir la ruta de salida .xlsx
                carpeta = os.path.dirname(pickle_path)
                nombre = os.path.splitext(os.path.basename(pickle_path))[0] + ".xlsx"
                excel_path = os.path.join(carpeta, nombre)
                
                # 3. Guardar a Excel
                df.to_excel(excel_path, index=False)
                
                print(f"✅ Exportado a Excel: {excel_path}")
            export_to_excel(contratos_base_pickle)            
        elif choice == "C":
            # Eliminar registros y archivos pickle
            from desagregada_to_pickle import handle_pickle
            contratos_base_pickle = os.path.join(Folder_procedimiento, f"{procedimiento}.pickle")
            handle_pickle(contratos_base_pickle, Folder_procedimiento) 
        elif choice == "D":
            # Genera la carpeta con contratos
            from STEP_D_Hard_Copy_Handling import STEP_D_Hard_Copy_Handling
            STEP_D_Hard_Copy_Handling(carpeta_contratos)
        elif choice == "E":
            # Genera la carpeta con contratos
            from STEP_E_copies_udpate import STEP_E_copies_udpate
            STEP_E_copies_udpate(Folder_procedimiento, procedimiento)            
        elif choice == "salir":
            print("Saliendo del programa.")
            break
        else:
            print("Selecciona una opción válida (A, B o C).")
            # Exporta la base a excel en caso de necesitarl