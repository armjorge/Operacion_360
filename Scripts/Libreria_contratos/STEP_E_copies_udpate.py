import pandas as pd
import os 
import sys
import shutil

def STEP_E_copies_udpate(Folder_procedimiento, procedimiento): 
    print("Logramos llamar una función desde un menú")
    #print(Folder_procedimiento) # Directorio del procedimiento
    #print(procedimiento)# Nombre del procedimiento
    contratos_base_pickle = os.path.join(Folder_procedimiento, f"{procedimiento}.pickle")
    df = pd.read_pickle(contratos_base_pickle)
    print(df.columns)
    Column_formalizado = {
        'Estatus': [
            {'label': 'NO formalizado', 'path': './NO formalizado'},
            {'label': 'Formalizado', 'path': './'}
        ]
    }
    for key, items in Column_formalizado.items():
        for value in items:
            label = value['label']
            path = value['path']
            df_filtered = df[df[key] == label]
            print(f"\nFilas con {key} = '{label}':")
            print(df_filtered.head())

    # Preguntar al usuario hasta que seleccione un índice válido
    # Mostrar opciones de columnas con índice
    column_keys = list(Column_formalizado.keys())
    for idx, key in enumerate(column_keys):
        print(f"{idx}) {key}")

    # Preguntar al usuario hasta que seleccione un índice válido
    while True:
        try:
            index_selected = int(input("Selecciona la columna a filtrar (por número): "))
            if 0 <= index_selected < len(column_keys):
                break
            else:
                print("Índice fuera de rango. Intenta nuevamente.")
        except ValueError:
            print("Entrada inválida. Por favor, ingresa un número.")

    chosen_key = column_keys[index_selected]
    chosen_filter = Column_formalizado[chosen_key]

    print(f"\nHas seleccionado la columna: {chosen_key}")
    print("Diccionarios disponibles como filtros:")
    #print(f"\nDiccionario de filtro\n{chosen_filter}")
    #for idx, item in enumerate(chosen_filter):
    #    print(f"{idx}) {item}")
    print("Está pensado para los contratos con el estado 'NO formalizado', de los que recibimos posteriormente la versión Formalizado")
    # Mostrar opciones con índice
    for idx, item in enumerate(chosen_filter):
        print(f"{idx}) {item}")

    # Selección del campo fuente
    while True:
        try:
            chosen_source = int(input("¿Cuál es el campo en la base y el archivo que quieres actualizar? (índice): "))
            if 0 <= chosen_source < len(chosen_filter):
                break
            else:
                print("Índice fuera de rango. Intenta nuevamente.")
        except ValueError:
            print("Entrada inválida. Por favor, ingresa un número.")

    estado_inicial = chosen_filter[chosen_source]

    # Selección del campo destino
    while True:
        try:
            chosen_target = int(input("¿Cuál es el campo de destino? (índice): "))
            if 0 <= chosen_target < len(chosen_filter):
                break
            else:
                print("Índice fuera de rango. Intenta nuevamente.")
        except ValueError:
            print("Entrada inválida. Por favor, ingresa un número.")

    estado_final = chosen_filter[chosen_target]

    # Mostrar selección final
    print("\nSelección realizada:")
    print("Estado inicial:", estado_inicial)
    
    df_filtrado = df[df[chosen_key] == estado_inicial['label']]
    print(f"\nFilas con {chosen_key} = '{estado_inicial['label']}':")
    print(df_filtrado.head())
    # Mostrar las columnas deseadas con sus índices
    print("\nFilas filtradas:")
    subset_cols = ['Contrato', 'Convenio modificatorio', 'Nombre del archivo', 'Estatus']

    if not all(col in df_filtrado.columns for col in subset_cols):
        raise ValueError("Una o más columnas no existen en el DataFrame.")

    for idx, row in df_filtrado[subset_cols].iterrows():
        print(f"{idx}) {row.to_dict()}")

    # Preguntar por índice de fila
    while True:
        try:
            chosen_row = int(input("\nÍndice del DataFrame a actualizar: "))
            if chosen_row in df_filtrado.index:
                break
            else:
                print("Índice fuera del rango del DataFrame filtrado. Intenta nuevamente.")
        except ValueError:
            print("Entrada inválida. Por favor, ingresa un número válido.")

    # Extraer la fila seleccionada
    df_row = df_filtrado.loc[[chosen_row]]

    print("\nFila seleccionada:")
    print(df_row)    
    print("Estado final:  ", estado_final)
    archivo_a_mover = df_row['Nombre del archivo'].values[0]
    path_archiv_a_mover = os.path.join(Folder_procedimiento, f"{archivo_a_mover}")
    print(path_archiv_a_mover)
    valid_dict = df_row.fillna('').iloc[0].to_dict()
    valid_dict['Estatus'] = estado_final['label']

    # Reemplazar en el nombre del archivo el texto del estado inicial por el del estado final
    if 'Nombre del archivo' in valid_dict and isinstance(valid_dict['Nombre del archivo'], str):
        valid_dict['Nombre del archivo'] = valid_dict['Nombre del archivo'].replace(
            estado_inicial['label'],
            estado_final['label']
        )
    #dict_file = STEP_C_write_label_to_PDF(path_archiv_a_mover, Folder_procedimiento, valid_dict, just_read=True)
    # Pegar el diccionario en el PDF y archivar el PDF.
    while True: 
        user_input = input("Tenemos diccionario creado ¿procedemos a pegarlo al PDF? si o no").strip().lower()
        if user_input == "si":        
            STEP_C_PDF_HANDLING(valid_dict, Folder_procedimiento)
            break
        else:
            print("❌ Consigue el PDF y vuelve para pegarle el diccionario.")
            continue
    
    target_path = os.path.join(Folder_procedimiento, estado_inicial['label'])
    create_directory_if_not_exists(target_path)

    # Ruta final con nombre de archivo
    archivo_destino = os.path.join(target_path, archivo_a_mover)

    try:
        shutil.move(path_archiv_a_mover, archivo_destino)
        print(f"✅ Archivo movido de:\n  {path_archiv_a_mover}\na:\n  {archivo_destino}")
    except Exception as e:
        print(f"❌ Error al mover el archivo: {e}")
    
    # Re-escribir el dataframe 
    df.at[chosen_row, 'Nombre del archivo'] = valid_dict['Nombre del archivo']
    df.at[chosen_row, 'Estatus'] = valid_dict['Estatus']

    print("\n✅ DataFrame actualizado en la fila seleccionada:\n")
    print(df.head())
    df.to_pickle(contratos_base_pickle)
    print(f"💾 DataFrame guardado en: {contratos_base_pickle}")


if __name__ == "__main__":
    script_folder_root = os.getcwd()
    print('\nLibrería script \n', script_folder_root)
    libreria_folder = os.path.abspath(os.path.join(script_folder_root, "..")) 
    print('\nLibrerías de scripts \n', libreria_folder)
    # Ensure the script folder is added to sys.path
    if libreria_folder not in sys.path:
        sys.path.append(libreria_folder)
    working_path = os.path.abspath(os.path.join(script_folder_root, "..", "..")) 
    print("\nPath de trabajo\n", working_path)
    procedimiento = 'LA-012M7B997-E115-2022'
    Folder_procedimiento = os.path.join(working_path, "Implementación", "Contratos", f"{procedimiento}")
    print('\nFolder procedimiento\n',Folder_procedimiento)
    from STEP_C_PDFhandling import STEP_C_PDF_HANDLING, STEP_C_write_label_to_PDF
    from folders_files_open import open_folder, create_directory_if_not_exists
    STEP_E_copies_udpate(Folder_procedimiento, procedimiento)