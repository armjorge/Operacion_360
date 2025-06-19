import os
import pandas as pd
import numpy as np


def base_pagos_atencion_proveedores(xlsx_atencion_proveedores):
    mensaje = "Iniciando la extracción de pagos reportados por la ventanilla"
    print(f"{'*'* len(mensaje)}\n{mensaje}\n{'*' * len(mensaje)}\n")
    # Buscar todos los archivos .xlsx en la carpeta
    files = [
        os.path.join(xlsx_atencion_proveedores, f)
        for f in os.listdir(xlsx_atencion_proveedores)
        if f.endswith('.xlsx')
    ]

    dfs = []
    for file in files:
        try:
            df = pd.read_excel(file)
            if 'FOLIO FISCAL' in df.columns:
                dfs.append(df[['FOLIO FISCAL']])
            else:
                print(f"[EXCLUIDO] No se encontró 'FOLIO FISCAL' en: {os.path.basename(file)}")
        except Exception as e:
            print(f"[ERROR] Al leer el archivo: {os.path.basename(file)}. Detalle: {e}")

    # Concatenar los dataframes válidos
    if dfs:
        df_base_pagos = pd.concat(dfs, ignore_index=True)
        df_base_pagos = df_base_pagos.dropna(subset=['FOLIO FISCAL'])
    else:
        df_base_pagos = pd.DataFrame(columns=['FOLIO FISCAL'])

    # Guardar el archivo de salida
    output_path = os.path.join(xlsx_atencion_proveedores, '..', 'Pagos_reportados.xlsx')
    df_base_pagos.to_excel(output_path, index=False)
    print(f"Archivo generado: {output_path}")


def C1_returning_values(dataframe, origin_column, match_column, return_value):
    #Helper function to find and return a value from a dataframe based on a match.
    return dataframe.loc[dataframe[match_column] == origin_column, return_value].values[0] if not dataframe.loc[dataframe[match_column] == origin_column].empty else None


def generacion_xlsx_atencion_proveedores(SAGI_path,sanciones_path, sheet_sanciones, camunda_path, atencion_proveedores_path, columnas_layout): 
    mensaje = "Generando el archivo layout de atención a proveedores"

    print(f"{'*'* len(mensaje)}\n{mensaje}\n{'*'* len(mensaje)}")
    # Cargar archivos de entrada
    df_SAGI = pd.read_excel(SAGI_path)
    df_sanciones = pd.read_excel(sanciones_path, sheet_name=sheet_sanciones)
    df_camunda = pd.read_excel(camunda_path)
    # Crear dataframe vacío con el layout solicitado
    df_atencion_proveedores = pd.DataFrame(columns=columnas_layout)

    df_atencion_proveedores['Contrato'] = df_SAGI['Número de contrato']
    df_atencion_proveedores['Proveedor']= df_SAGI['Proveedor']
    df_atencion_proveedores['Factura'] = df_atencion_proveedores['Factura'] = 'P-' + df_SAGI['Número de factura'].astype(str)
    df_atencion_proveedores['Folio Fiscal'] = df_SAGI['Folio fiscal']
    df_atencion_proveedores['Orden suministro'] = df_SAGI['Orden de suministro']
    df_atencion_proveedores['Importe'] = df_SAGI['Total']
    df_atencion_proveedores['Importe Sanción'] = df_atencion_proveedores['Orden suministro'].apply(lambda x: C1_returning_values(df_sanciones, x, 'ORDEN DE SUMINISTRO', 'PENA'))
    df_atencion_proveedores['Importe Deductiva'] = 0
    df_atencion_proveedores['Cédula'] = df_atencion_proveedores['Orden suministro'].apply(lambda x: C1_returning_values(df_sanciones, x, 'ORDEN DE SUMINISTRO', 'OFICIO'))
    df_atencion_proveedores['Fte Fmto'] = np.select(
        [
            df_atencion_proveedores['Contrato'] == 'LA-E115-2022-MED-INSABI-122-2023/2024',
            df_atencion_proveedores['Contrato'].isin([
                'LA-E115-2022-MED-INSABI-034-2023/2024',
                'LA-E115-2022-MED-INSABI-188-2023/2024'
            ])
        ],
        ['FONSABI', '32%'],
        default='No identificado'
    )
    df_atencion_proveedores['Año Contrato'] = 2023 
    df_atencion_proveedores['Fecha ingreso factura'] = ""
    df_atencion_proveedores['Estatus'] = df_SAGI['Estado de la factura']

    print(df_atencion_proveedores.head(10))
    output_path =  os.path.join(atencion_proveedores_path, 'Atencion_proveedores_layout.xlsx')
    df_atencion_proveedores.to_excel(output_path, index=False)



if __name__ == "__main__":
    # Esto solo se ejecuta cuando corres el script, no cuando lo importas desde otro módulo.
    folder_root = r"C:\Users\arman\Dropbox\3. Armando Cuaxospa\Adjudicaciones\Licitaciones 2025\E115 360"
    SAGI_path = os.path.join(folder_root, "Implementación", 'SAGI', '05 29 2024 ESTATUS_SAGI.xlsx') # cambiar dinámicamente
    atencion_proveedores_path = os.path.join(folder_root, "Implementación", "IMSS-Bienestar Atención Proveedores")
    xlsx_atencion_proveedores = os.path.join(folder_root, "Implementación", "IMSS-Bienestar Atención Proveedores", "Atencion proveedores raw")
    base_pagos_atencion_proveedores(xlsx_atencion_proveedores)
    sanciones_path = os.path.join(folder_root, "Implementación", 'Sanciones IMSSB', 'Penas-Oficios-Ordenes.xlsx')
    sheet_sanciones = 'Desglose'
    camunda_path = os.path.join(folder_root, "Implementación", 'CAMUNDA', 'Camunda v2025.xlsx')
    columnas_layout = ['Contrato', 'Proveedor', 'Factura', 'Folio Fiscal', 'Orden suministro', 'Importe', 'Importe Sanción', 'Importe Deductiva', 'Cédula', 'Fte Fmto', 'Año Contrato', 'Fecha ingreso factura', 'Estatus']
    generacion_xlsx_atencion_proveedores(SAGI_path,sanciones_path, sheet_sanciones, camunda_path, atencion_proveedores_path, columnas_layout)