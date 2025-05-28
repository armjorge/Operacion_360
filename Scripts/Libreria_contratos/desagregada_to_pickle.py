import os
import re
import glob
import pandas as pd

def clean_filename(name: str) -> str:
    """
    Sustituye en el nombre cualquier carácter no permitido en
    nombres de archivo de Windows por un guion bajo.
    """
    # Conjunto de caracteres inválidos: <>:"/\|?*
    return re.sub(r'[<>:"/\\|?*]', '_', name).strip()

def save_to_pickle(df_grouped, desagregadas_folder):
    """
    Para cada valor único en df_grouped['Procedimiento']:
      1. Limpia el nombre para usarlo como filename.
      2. Filtra el DataFrame por ese procedimiento.
      3. Guarda el subset en desagregadas_folder/{clean_name}.pickle
      4. Imprime información de progreso.
    """
    # Asegurarse de que la carpeta existe
    os.makedirs(desagregadas_folder, exist_ok=True)

    procedimientos = df_grouped['Procedimiento'].unique()
    print(f"Se encontraron {len(procedimientos)} procedimientos únicos.")
    for proc in procedimientos:
        clean_name = clean_filename(proc)
        filename = f"{clean_name}.pickle"
        filepath = os.path.join(desagregadas_folder, filename)

        subset = df_grouped[df_grouped['Procedimiento'] == proc]
        subset.to_pickle(filepath)

        print(f"• Procedimiento: '{proc}' ➔ '{filename}' ({len(subset)} registros)")

    print("Todos los archivos han sido guardados exitosamente.")

def load_pickles(desagregadas_folder):
    # 1. Obtiene rutas completas y nombres de archivo
    pickle_paths = glob.glob(os.path.join(desagregadas_folder, "*.pickle"))
    basenames = [os.path.basename(p) for p in pickle_paths]

    # 2. Muestra al usuario las opciones
    print(f"Encontrados {len(basenames)} archivos .pickle:")
    for name in basenames:
        print(" -", name)

    # 3. Pide la selección hasta que sea válida
    while True:
        selected = input("Por favor pega el nombre del archivo que quieres cargar: ").strip()
        if selected in basenames:
            full_path = os.path.join(desagregadas_folder, selected)
            # 4. Carga el DataFrame
            df_procedure = pd.read_pickle(full_path)
            print(f"✔ Archivo '{os.path.splitext(selected)[0]}' cargado correctamente ({len(df_procedure)} filas).")

            return df_procedure, os.path.splitext(selected)[0]
        else:
            print("✖ Opción inválida, inténtalo otra vez.")



def base_procedimiento_pickle(Folder_procedimiento, procedimiento, columnas):
    """
    Asegura que exista un pickle para el procedimiento dado:
      - Si NO existe: crea uno vacío con las columnas indicadas.
      - Si existe: lo carga y devuelve.

    Args:
      Folder_procedimiento (str): Ruta a la carpeta donde estarán los pickles.
      procedimiento (str): Nombre base del pickle (sin extensión).
      columnas (list of str): Lista de nombres de columnas para inicializar, si toca crear.

    Returns:
      pd.DataFrame: El DataFrame cargado o recién creado.
    """
    # 1. Asegurar existencia de carpeta
    os.makedirs(Folder_procedimiento, exist_ok=True)

    # 2. Construir ruta del pickle
    filepath = os.path.join(Folder_procedimiento, f"{procedimiento}.pickle")

    # 3. Crear si no existe
    if not os.path.exists(filepath):
        # DataFrame vacío con las columnas
        df_empty = pd.DataFrame(columns=columnas)
        df_empty.to_pickle(filepath)
        print(f"✔ No se encontró '{procedimiento}.pickle'. Se creó un DataFrame vacío con columnas {columnas}.")
        return df_empty

    # 4. Si existe, cargarlo
    df_procedimientos = pd.read_pickle(filepath)
    print(f"✔ Cargado pickle existente: '{procedimiento}.pickle' ({len(df_procedimientos)} filas).")
    return df_procedimientos
