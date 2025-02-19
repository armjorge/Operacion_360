import os
import sys
import pandas as pd
script_directory = os.path.dirname(os.path.abspath(__file__))
function_library = os.path.abspath(os.path.join(script_directory, 'Libreria_contratos'))
sys.path.append(function_library) 
from STEP_A_orchestration import STEP_A_orchestration
from folders_files_open import load_dataframe

def administracion_de_contratos(data_warehouse, working_folder, carpeta_contratos):
    console_path = os.path.join(data_warehouse, 'desagregada.xlsx')
    sheet = 'Sheet1'
    columns = ['Institución', 'Procedimiento', 'Clave', 'Descripción', 'Precio', 'Total']
    df_desagregada = load_dataframe(console_path, sheet, columns)
    print(df_desagregada.columns, '/n')
    print(df_desagregada.head())
    
    procedimiento_column = 'Procedimiento'
    institucion_column = 'Institución'
    if procedimiento_column in columns and institucion_column in columns:
        print(f"Columna de procedimiento {procedimiento_column} e institución {institucion_column} solicitadas y presentes al cargar el dataframe")
    unique_procedimientos = df_desagregada[procedimiento_column].dropna().unique()

    if len(unique_procedimientos) == 0:
        print(f"No hay procedimientos cargados en el archivo {os.path.basename(console_path)} hoja {sheet}.")
        return  # Exit the code
    # Display unique procedimientos
    for i, procedimiento_unique in enumerate(unique_procedimientos):
        print(f"{i + 1}) {procedimiento_unique}")

    # Ask user to select a procedure
    try:
        selected_index = int(input("Seleccione el número del procedimiento: ")) - 1
        if 0 <= selected_index < len(unique_procedimientos):
            selected_procedimiento = unique_procedimientos[selected_index]
            print(f"Procedimiento seleccionado: {selected_procedimiento}")
        else:
            print("Número no válido.")
    except ValueError:
        print("Entrada no válida. Debe ingresar un número.")
    filter_query = f'{procedimiento_column} == "{selected_procedimiento}"'
    df_desagregada = df_desagregada.query(filter_query)
    print(df_desagregada.head())
    # Seleccionar el 
    print("¿1) Capturar un contrato o 2) extraer base contratos existentes?")
    while True:
        step_0 = input("Seleccione una opción (1 o 2): ")
        if step_0 == "1":
            STEP_A_orchestration(working_folder, df_desagregada, institucion_column, selected_procedimiento, carpeta_contratos)
        elif step_0 == "2":
            # Suponiendo que aquí debería haber algún código para extraer la base de contratos
            print("Base de contratos extraída correctamente.")
        else:
            print("Opción no válida. Por favor, seleccione 1 o 2")
            print("¿1) Capturar un contrato o 2) extraer base contratos existentes?")

  
# Call the main function
#if __name__ == "__main__":
#    main()