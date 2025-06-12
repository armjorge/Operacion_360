from STEP_B_Dict import STEP_B_get_string_populated
from STEP_C_PDFhandling import STEP_C_PDF_HANDLING
import pandas as pd
import os

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