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



def base_contratos_convenios_pickle(Folder_procedimiento, procedimiento, columnas):
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
        df_contratos_convenios = pd.DataFrame(columns=columnas)
        df_contratos_convenios.to_pickle(filepath)
        print(f"✔ No se encontró '{procedimiento}.pickle'. Se creó un DataFrame vacío con columnas {columnas}.")
        return df_contratos_convenios

    # 4. Si existe, cargarlo
    df_contratos_convenios = pd.read_pickle(filepath)

    print(f"✔ Cargado pickle existente: '{procedimiento}.pickle' ({len(df_contratos_convenios)} filas).")
    return df_contratos_convenios


def handle_pickle(pickle_path, Folder_procedimiento):
    """
    1. Carga el DataFrame desde pickle_path.
    2. Muestra los primeros 20 contratos.
    3. Pregunta al usuario si desea eliminar algún registro.
    4. Si responde "si", muestra todos los índices y contratos, pide el índice a eliminar,
       borra ese renglón del DataFrame y elimina el archivo PDF correspondiente en Folder_procedimiento.
    5. Guarda el pickle actualizado.
    """
    # 1. Cargar el DataFrame
    df = pd.read_pickle(pickle_path)

    # 2. Mostrar los primeros 20 valores de 'Contrato'
    len_print = 20
    print(df['Contrato'].head(len_print))
    print(f"Se imprimen los primeros {len_print} registros de {len(df['Contrato'])} contratos.")

    # 3. Preguntar si desea eliminar
    user_choice = input("¿Quieres eliminar algún registro? Ten en cuenta que también se eliminará el PDF (si/no): ").strip().lower()
    if user_choice != 'si':
        print("✅ No se realizará ninguna eliminación.")
        return

    # 4. Mostrar todos los índices con su 'Contrato'
    print("\nListado completo de índices y contratos:")
    for idx, row in df.iterrows():
        print(f"  {idx} → {row['Contrato']}")

    # 5. Pedir al usuario el índice a eliminar
    sel = input("\nÍndice del contrato a eliminar: ").strip()
    try:
        idx_to_del = int(sel)
    except ValueError:
        print("❌ Índice inválido. Debes ingresar un número entero.")
        return

    if idx_to_del not in df.index:
        print("❌ El índice ingresado no existe en el DataFrame.")
        return

    # 6. Obtener el nombre del archivo que se debe eliminar
    row = df.loc[idx_to_del]
    pdf_filename = row['Nombre del archivo']
    pdf_path = os.path.join(Folder_procedimiento, pdf_filename)

    # 7. Eliminar el archivo PDF si existe
    if os.path.exists(pdf_path):
        try:
            os.remove(pdf_path)
            print(f"🗑️ Se eliminó el archivo PDF: {pdf_filename}")
        except Exception as e:
            print(f"❌ Error al eliminar el archivo PDF: {e}")
            return
    else:
        print(f"⚠️ No se encontró el archivo PDF en: {pdf_path}")

    # 8. Eliminar la fila del DataFrame y guardar el pickle actualizado
    df = df.drop(idx_to_del)
    df.to_pickle(pickle_path)
    print("✅ Registro eliminado del DataFrame y pickle actualizado.")
