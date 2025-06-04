from STEP_B_Dict import STEP_B_get_string_populated
from STEP_C_PDFhandling import STEP_C_PDF_HANDLING

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
    valid_dict = STEP_B_get_string_populated(df_desagregada_procedimiento, tipo, institucion_column, selected_procedimiento, carpeta_contratos)
    print("\nDiccionario Procesado:\n", valid_dict)
    
    # Pegar el diccionario en el PDF y archivar el PDF.
    while True: 
        user_input = input("Tenemos diccionario creado ¿procedemos a pegarlo al PDF? si o no").strip().lower()
        if user_input == "si":        
            STEP_C_PDF_HANDLING(valid_dict, carpeta_contratos)
            break
        else:
            print("❌ Consigue el PDF y vuelve para pegarle el diccionario.")
            continue

    #result_dict = read_last_page_pdf(pdf_path_list, valid_dict)
    #print(f"Inserto agregado: {result_dict}")
    #df_feed = STEP_C_feed_DF(df_referencia_interna, result_dict)