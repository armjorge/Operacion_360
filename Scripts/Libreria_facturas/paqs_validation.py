import os
import re
import glob
import pandas as pd
import numpy as np

def correccion_types(df_entregas_o_altas, df_facturas, info_types):
    if info_types == 'IMSS': 
        print(f"Iniciamos la corrección de tipos para el {info_types}")
        print(f"Número de filas del dataframe facturas al iniciar {df_facturas.index.size}")
        print(f"Número de filas del dataframe altas al iniciar {df_entregas_o_altas.index.size}")
        # El siguiente paso es debido a la existencia de valores infinitos en la columna Referencia
        df_facturas['Referencia'] = (
            df_facturas['Referencia']
            .replace([np.inf, -np.inf], np.nan)  # Inf  → NaN
            .fillna(0)                           # NaN  → 0
            .astype('int64')                     # ahora ya solo hay valores válidos para int64
        )        

        df_facturas['Referencia'] = df_facturas['Referencia'].astype('int64')
        df_entregas_o_altas['noOrden'] = df_entregas_o_altas['noOrden'].astype('int64')
       
        # (II.2) Duplicados. 
        duplicados_facturas = df_facturas.duplicated().sum()
        print(f"El dataframe facturas tiene {duplicados_facturas} filas duplicadas, vamos a removerlas\n")
        df_facturas = df_facturas.drop_duplicates()
        # (II.3) Ausentes
        print("Removemos del dataframe facturas aquellas filas con Alta y Orden vacíos\n")
        mask = ((df_facturas['Referencia'].isna() | (df_facturas['Referencia'] == 0)) & df_facturas['Alta'].isna()
)
        print(f"Totales de las filas con Referencia y Alta ausentes = {mask.index.size}")
        df_facturas = df_facturas.loc[~mask].reset_index(drop=True)

        return df_entregas_o_altas, df_facturas

    else: 
        print("no considerado aún")   

def validacion_de_paqs(dict_path_sheet, dic_columnas, paq_folder, df_entregas_o_altas, info_types):
    # (I) Carga
    columnas_objetivo = ["Folio", "Referencia", "Alta", "Total", "UUID"]
    df_facturas = pd.DataFrame(columns=columnas_objetivo)

    # (2) Iteramos simultáneamente sobre ambos diccionarios. 
    #     Se asume que dict_path_sheet y dic_columnas ya están “alineados” en el mismo orden de inserción.
    for (ruta_excel, nombre_hoja), lista_cols in zip(dict_path_sheet.items(), dic_columnas.values()):
        # ruta_excel: p.ej. r"C:\Users\arman\Dropbox\FACT 2023\Generacion facturas IMSS VFinal.xlsx"
        # nombre_hoja:    p.ej. "Reporte Paq"
        # lista_cols:     p.ej. ["Folio", "Referencia", "Alta", "Total", "UUID"]

        # (3) Leemos únicamente las columnas indicadas en lista_cols
        df_temp = pd.read_excel(
            ruta_excel,
            sheet_name=nombre_hoja,
            usecols=lista_cols
        )

        # (4) Concatenamos al DataFrame global
        df_facturas = pd.concat([df_facturas, df_temp], ignore_index=True)
    # (II) Limpia

    # (II.1) Corrección de tipos, remover duplicados y lógica de referencias.
    df_entregas_o_altas, df_facturas = correccion_types(df_entregas_o_altas, df_facturas, info_types)
 
    print("Información del dataframe altas: \n")
    print(df_entregas_o_altas.info())
    print("Información del dataframe de facturas: \n")
    print(df_facturas.info())
    excel_facturas= os.path.join(paq_folder, f"{info_types}.xlsx")
    df_facturas.to_excel(excel_facturas, index=False)
    print("\nExcel generado de facturas generado exitosamente\n")

if __name__ == "__main__":
    # Esto solo se ejecuta cuando corres el script, no cuando lo importas desde otro módulo.
    #Test variables: 
    dict_path_sheet = {'C:\\Users\\arman\\Dropbox\\FACT 2023\\Generacion facturas IMSS VFinal.xlsx': 'Reporte Paq', 'C:\\Users\\arman\\Dropbox\\FACT 2024\\Generacion facturas IMSS 2024.xlsx': 'Reporte Paq', 'C:\\Users\\arman\\Dropbox\\FACT 2025\\Copy of Generacion facturas IMSS 2024.xlsx': 'Reporte Paq'} 
    dic_columnas = {'IMSS_2023': ['Folio', 'Referencia', 'Alta', 'Total', 'UUID'], 'IMSS_2024': ['Folio', 'Referencia', 'Alta', 'Total', 'UUID'], 'IMSS_2025': ['Folio', 'Referencia', 'Alta', 'Total', 'UUID']} 
    paq_folder = r"C:\Users\arman\Dropbox\3. Armando Cuaxospa\Adjudicaciones\Licitaciones 2025\E115 360\Implementación\Facturas\IMSS"
    altas_path = r"C:\Users\arman\Dropbox\3. Armando Cuaxospa\Adjudicaciones\Licitaciones 2025\E115 360\Implementación\SAI\Ordenes_altas.xlsx"
    df_entregas_o_altas = pd.read_excel(altas_path, sheet_name='df_altas')
    info_types = 'IMSS'
    validacion_de_paqs(dict_path_sheet, dic_columnas, paq_folder, df_entregas_o_altas, info_types)
