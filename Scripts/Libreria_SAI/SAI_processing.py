import os
import re
import glob
import pandas as pd

def load_dataframes(files, date_regex, date_parse_format, pivot_columns):
    """
    Carga una lista de archivos Excel, extrae la fecha del nombre de archivo, 
    la añade como columna 'file_date', ordena y elimina duplicados por pivot_columns
    conservando siempre la fila más reciente (mayor file_date).

    Parameters
    ----------
    files : list of str
        Rutas a los .xlsx que quieres procesar.
    date_regex : str
        Expresión regular con un grupo capturante que extrae la fecha del filename.
        Ejemplo: r'(\d{4} \d{2} \d{2})' para fechas en formato 'YYYY MM DD'.
    date_parse_format : str
        Formato de datetime para pd.to_datetime, p.ej. '%Y %m %d'.
    pivot_columns : list of str
        Columnas que definen duplicados. Se conservará la fila con file_date más reciente.

    Returns
    -------
    pd.DataFrame
        DataFrame único resultante.
    """
    df_list = []
    for filepath in files:
        fname = os.path.basename(filepath)
        m = re.search(date_regex, fname)
        if not m:
            raise ValueError(f"No se pudo extraer fecha de '{fname}' usando regex '{date_regex}'")
        date_str = m.group(1)
        # Parsear fecha
        file_date = pd.to_datetime(date_str, format=date_parse_format)
        print(f"{fname} → file_date: {file_date.date()}")

        # Cargar Excel
        df = pd.read_excel(filepath)
        # Añadir columna file_date
        df['file_date'] = file_date
        # Asegurar tipo datetime
        df['file_date'] = pd.to_datetime(df['file_date'])
        df_list.append(df)

    # Concatenar todos
    full_df = pd.concat(df_list, ignore_index=True)
    # Ordenar por fecha ascendente
    full_df = full_df.sort_values('file_date')

    # Eliminar duplicados según pivot_columns, quedándose con la fila más reciente (última)
    full_df = full_df.drop_duplicates(subset=pivot_columns, keep='last')

    return full_df


def merge_SAI_files(SAI_folder, alta_pivot_columns, orden_pivot_columns,
                    date_regex, date_parse_format, Output_filename):
    """
    Busca subcarpetas que terminen en 'Final' dentro de SAI_folder, 
    carga por separado los archivos de 'Altas' y 'Ordenes', aplica load_dataframes
    y devuelve ambos DataFrames.

    Parameters
    ----------
    SAI_folder : str
        Carpeta raíz donde buscar subcarpetas '*Final'.
    alta_pivot_columns : list of str
        Columnas para deduplicar en los archivos de Altas.
    orden_pivot_columns : list of str
        Columnas para deduplicar en los archivos de Ordenes.
    date_regex : str, opcional
        Regex para extraer fecha (por defecto busca 'YYYY MM DD').
    date_parse_format : str, opcional
        Formato de fecha para pd.to_datetime (por defecto '%Y %m %d').

    Returns
    -------
    tuple(pd.DataFrame, pd.DataFrame)
        (df_altas, df_ordenes)
    """
    # 1) Encontrar todas las carpetas que acaben en 'Final'
    finals = [
        os.path.join(SAI_folder, d)
        for d in os.listdir(SAI_folder)
        if os.path.isdir(os.path.join(SAI_folder, d)) and d.endswith("Final")
    ]
    if not finals:
        raise FileNotFoundError(f"No se hallaron carpetas terminadas en 'Final' en {SAI_folder}")

    # 2) Recorrer cada carpeta 'Final' y recolectar todos los archivos
    all_files_altas = []
    all_files_ordenes = []
    for folder in finals:
        patrones_altas = glob.glob(os.path.join(folder, "*Altas.xlsx"))
        patrones_ordenes = glob.glob(os.path.join(folder, "*Ordenes.xlsx"))
        all_files_altas.extend(patrones_altas)
        all_files_ordenes.extend(patrones_ordenes)

    # 3) Cargar y procesar
    df_altas   = load_dataframes(all_files_altas,   date_regex, date_parse_format, alta_pivot_columns)
    df_ordenes = load_dataframes(all_files_ordenes, date_regex, date_parse_format, orden_pivot_columns)
    
    # 4) Guardar el archivo
    output_path = os.path.join(SAI_folder, Output_filename)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_altas.to_excel(writer, sheet_name='df_altas', index=False)
        df_ordenes.to_excel(writer, sheet_name='df_ordenes', index=False)
    print(f"Archivo guardado en: {os.path.basename(output_path)}")
    
    #return df_altas, df_ordenes


