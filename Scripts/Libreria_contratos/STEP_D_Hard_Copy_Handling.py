import os 
import pandas as pd

def STEP_D_Hard_Copy_Handling(carpeta_contratos):
    """
    Busca en cada subdirectorio de carpeta_contratos un archivo .pickle
    cuyo nombre coincida con el del subdirectorio. Luego imprime:
      <nombre_de_carpeta> -> <ruta_al_pickle>
    """
    # I ELIJE EL FOLDER A PROCESAR
    # 1) Construir el mapping de subcarpetas a archivos .pickle
    mapping = {}
    for entry in os.listdir(carpeta_contratos):
        subdir = os.path.join(carpeta_contratos, entry)
        if not os.path.isdir(subdir):
            continue
        candidate = os.path.join(subdir, f"{entry}.pickle")
        if os.path.isfile(candidate):
            mapping[subdir] = candidate

    if not mapping:
        print("⚠️  No se encontraron archivos .pickle en ninguna subcarpeta.")
        return

    # 2) Mostrar lista numerada
    items = list(mapping.items())  # [(ruta_carpeta, ruta_pickle), ...]
    for idx, (carpeta, archivo) in enumerate(items):
        nombre = os.path.basename(carpeta)
        print(f"{idx}: {nombre} -> {os.path.basename(archivo)}")

    # 3) Pedir selección al usuario, validando que sea un entero en rango
    while True:
        choice = input(f"¿Cuál de los índices anteriores quieres cargar? [0–{len(items)-1}]: ")
        if choice.isdigit():
            idx = int(choice)
            if 0 <= idx < len(items):
                break
        print("Entrada inválida. Por favor, introduce un número válido.")
    # II CARGA EL PICKLE DEL FOLDER ELEGIDO
    # 4) Cargar y mostrar el DataFrame
    _, selected_pickle = items[idx]
    df_pickle = pd.read_pickle(selected_pickle)
    pickle_name = os.path.basename(selected_pickle)
    procedimiento, extension = os.path.splitext(pickle_name)  
    df_carpeta = df_pickle[
        ['Contrato', 'Convenio modificatorio', 'Estatus', 'Objeto del convenio', 'Tipo']
    ].copy()

    print("Procedimiento", procedimiento)
    df_carpeta['Contrato'] = (
        df_carpeta['Contrato']
        + df_carpeta['Convenio modificatorio']
            .fillna('')
            .apply(lambda x: f"-{x}" if x != '' else '')
    )

    df_carpeta = df_carpeta.drop(columns=['Convenio modificatorio'])
    # III DUPLICADOS: Formalizados y no formalizados. 
    # 1) Encuentra todos los índices que hay que eliminar
    to_drop = []

    for contrato, group in df_carpeta.groupby('Contrato'):
        if len(group) > 1:
            # ¿Hay al menos un "Formalizado"?
            formalizados = group[group['Estatus'] == 'Formalizado']
            if not formalizados.empty:
                # Elimino todos los NO formalizado
                non_form = group[group['Estatus'] != 'Formalizado']
                for idx in non_form.index:
                    print(f"Dropping row index {idx} for Contrato {contrato} (Estatus='NO formalizado')")
                    to_drop.append(idx)
            else:
                # No hay ningún Formalizado, dejo sólo el primero
                for idx in group.index[1:]:
                    print(f"No 'Formalizado' found for Contrato {contrato}. Dropping duplicate at index {idx}")
                    to_drop.append(idx)

    # 2) Drop de una sola vez
    df_carpeta.drop(index=to_drop, inplace=True)
    # 3) Resetear índice (opcional, pero suele ayudar)
    df_carpeta.reset_index(drop=True, inplace=True)
    # 4) Finalmente, elimino la columna Estatus
    df_carpeta.drop(columns=['Estatus'], inplace=True)
    # IIII 
    # 1) Construir ruta al Excel
    user_info_path = os.path.join(
        carpeta_contratos,
        procedimiento,
        f"{procedimiento}_hard_copy.xlsx"
    )

    if os.path.isfile(user_info_path):
        # 2) Cargar el Excel existente
        df_user_info = pd.read_excel(user_info_path)

        # 3) Validar columnas requeridas
        required = ['Contrato', 'Objeto del convenio', 'Tipo', 'Carpeta física']
        missing = [c for c in required if c not in df_user_info.columns]
        if missing:
            print(f"⚠️ El archivo `{user_info_path}` no tiene columnas: {missing}.")
            print("   Por favor corrige el archivo y vuelve a ejecutar.")
            return  # salimos para que el usuario lo corrija

        print(f"✔ Usando plantilla existente: {user_info_path}")

        # 4) Asegurar la columna en df_carpeta antes de merge
        df_carpeta['Carpeta física'] = ''

        # 5) Traer sólo los valores de 'Carpeta física' coincidentes por 'Contrato'
        df_carpeta = df_carpeta.merge(
            df_user_info[['Contrato', 'Carpeta física']],
            on='Contrato',
            how='left',
            suffixes=('', '_user')
        )
        # 6) Rellenar y limpiar sufijo
        df_carpeta['Carpeta física'] = df_carpeta['Carpeta física_user'].fillna('')
        df_carpeta.drop(columns=['Carpeta física_user'], inplace=True)

        # 7) Guardar de vuelta sobre el mismo archivo
        df_carpeta.to_excel(user_info_path, index=False)
        print(f"📄 Actualizado: {user_info_path}")

    else:
        # 8) No existe: crearlo con columna vacía y pedir al usuario que la llene
        df_carpeta['Carpeta física'] = ''
        df_carpeta.to_excel(user_info_path, index=False)
        print(f"⚠️ He creado: {user_info_path}")
        print("   Por favor llena la columna 'Carpeta física' con la ruta donde guardaste la copia.")

if __name__ == "__main__":
    # Esto solo se ejecuta cuando corres el script, no cuando lo importas desde otro módulo.
    carpeta_contratos = r"C:\Users\arman\Dropbox\3. Armando Cuaxospa\Adjudicaciones\Licitaciones 2025\E115 360\Implementación\Contratos"
    STEP_D_Hard_Copy_Handling(carpeta_contratos)
