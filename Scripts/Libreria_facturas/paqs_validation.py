import os
import re
import glob
import pandas as pd
import numpy as np
from openpyxl import load_workbook


def validacion_de_paqs(dict_path_sheet, dic_columnas, paq_folder, altas_path, altas_sheet, info_types,xlsx_database ):
    # (I) Carga
    df_entregas_o_altas = pd.read_excel(altas_path, sheet_name=altas_sheet)
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
   
    # III.I Cargar IMSS y ligar.     
    df_altas_df_facturas = {
        'noAlta': 'Alta',
        'noOrden': 'Referencia'
    }

    df_entregas_o_altas['Factura'] = multi_column_lookup(
        df_to_fill=df_entregas_o_altas,
        df_to_consult=df_facturas,
        match_columns=df_altas_df_facturas,
        return_column='Folio',
        default_value='sin match'
    )
    # III.II Cargar IMSS y ligar.     
    df_altas_df_facturas = {
        'Alta': 'noAlta',
        'Referencia': 'noOrden'
    }

    df_facturas['Alta_ligada'] = multi_column_lookup(
        df_to_fill=df_facturas,
        df_to_consult=df_entregas_o_altas,
        match_columns=df_altas_df_facturas,
        return_column='noAlta',
        default_value='alta no localizada'
    )
    

    df_facturas.to_excel(excel_facturas, index=False)

    #IV Sobreescribir UUID y totales 
    print("Vamos a poblar el UUID de la base de facturación con info extraída de los XML's")
    if os.path.exists(xlsx_database):
        columna_UUID ='UUID'
        df_database = pd.read_excel(xlsx_database)
        df_database = (
            df_database
            .drop_duplicates(subset='UUID', keep='first')
            .reset_index(drop=True)
        )
        df_UUIDS = {'Folio': 'Folio'}
        df_facturas['UUID'] = multi_column_lookup(
            df_to_fill=df_facturas,
            df_to_consult=df_database,
            match_columns=df_UUIDS,
            return_column=columna_UUID,
            default_value=f'{columna_UUID} no localizado'
        )
    
    if os.path.exists(xlsx_database):
        columna_retorno ='Importe'
        columna_poblar = 'Total'
        print(f"Vamos a poblar l columna {columna_poblar} con de la columna {columna_retorno} base de facturación con info extraída de los XML's")
        df_database = pd.read_excel(xlsx_database)
        df_database = (
            df_database
            .drop_duplicates(subset='UUID', keep='first')
            .reset_index(drop=True)
        )
        columns_totales_match = {'Folio': 'Folio'}
        df_facturas[columna_poblar] = multi_column_lookup(
            df_to_fill=df_facturas,
            df_to_consult=df_database,
            match_columns=columns_totales_match,
            return_column=columna_retorno,
            default_value=f'{columna_UUID} no localizado'
        )
        
    # Cargar el archivo conservando las hojas
    with pd.ExcelWriter(altas_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_entregas_o_altas.to_excel(writer, sheet_name=altas_sheet, index=False)

    print("\nExcel generado de facturas generado exitosamente\n")



def multi_column_lookup(df_to_fill, df_to_consult, match_columns: dict, return_column, default_value='sin match'):
    """
    Realiza búsqueda con múltiples columnas y retorna valor o advertencias
    Args:
        df_to_fill (pd.DataFrame): DataFrame que queremos llenar.
        df_to_consult (pd.DataFrame): DataFrame que consultamos.
        match_columns (dict): {col_df_to_fill: col_df_to_consult} pares de columnas para hacer match.
        return_column (str): Columna en df_to_consult con el valor a traer.
        default_value (any): Valor si no hay coincidencia.
    Returns:
        pd.Series: Valores para poblar la columna deseada.
    """

    if not isinstance(df_to_consult, pd.DataFrame):
        raise TypeError(f"El argumento 'df_to_consult' debe ser un DataFrame, se recibió: {type(df_to_consult)}")

    result_list = []
    message_falta_liga = 'Renglones del dataframe a llenar sin filtro en el dataframe a consultar'
    ligas_duplicadas = 0
    ligas_vacías = 0
    for _, row in df_to_fill.iterrows():
        # Construir filtro booleano dinámico
        mask = pd.Series([True] * len(df_to_consult))

        for source_col, consult_col in match_columns.items():
            mask &= df_to_consult[consult_col] == row[source_col]

        filtered = df_to_consult[mask]

        if filtered.empty:
            result_list.append(default_value)
            print("\tEncontramos renglones sin poder ligar")
            ligas_vacías += 1
        elif len(filtered) > 1:
            resultados_duplicados = ", ".join(filtered[return_column].astype(str).tolist())
            result_list.append(f'Peligro: 2 resultados {resultados_duplicados}')
            ligas_duplicadas += 1
        else:
            result_list.append(filtered.iloc[0][return_column])
    success_message = "✅ Se ligaron el 100% de los renglones y no hubo duplicados."
    if ligas_duplicadas == 0 and ligas_vacías == 0:
        print(f"{'*'*len(success_message)}\n{success_message}\n{'*'*len(success_message)}")
    elif ligas_duplicadas == 0 and ligas_vacías > 0:
        print("⚠️ Hay renglones que no se pudieron llenar con el dataframe consultado.")
    elif ligas_duplicadas > 0 and ligas_vacías == 0:
        print("⚠️ Hay renglones para los que se encontraron más de un resultado en el dataframe de consulta.")
    else:
        print("⚠️ Hay renglones vacíos y renglones con duplicados.")

    return pd.Series(result_list, index=df_to_fill.index)


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
        mask = ((df_facturas['Referencia'].isna() | (df_facturas['Referencia'] == 0)) & df_facturas['Alta'].isna())
        print(f"Totales de las filas con Referencia y Alta ausentes = {mask.index.size}")
        df_facturas = df_facturas.loc[~mask].reset_index(drop=True)
        
        df_facturas = (
            df_facturas[~df_facturas['Total'].isin([0, '', ' '])]
            .dropna(subset=['Total'])
            .reset_index(drop=True)
        )



        return df_entregas_o_altas, df_facturas

    else: 
        print("no considerado aún")   


if __name__ == "__main__":
    # Esto solo se ejecuta cuando corres el script, no cuando lo importas desde otro módulo.
    #Test variables: 
    folder_root = r"C:\Users\arman\Dropbox\3. Armando Cuaxospa\Adjudicaciones\Licitaciones 2025\E115 360"
    dict_path_sheet = {'C:\\Users\\arman\\Dropbox\\FACT 2023\\Generacion facturas IMSS VFinal.xlsx': 'Reporte Paq', 'C:\\Users\\arman\\Dropbox\\FACT 2024\\Generacion facturas IMSS 2024.xlsx': 'Reporte Paq', 'C:\\Users\\arman\\Dropbox\\FACT 2025\\Copy of Generacion facturas IMSS 2024.xlsx': 'Reporte Paq'} 
    dic_columnas = {'IMSS_2023': ['Folio', 'Referencia', 'Alta', 'Total', 'UUID'], 'IMSS_2024': ['Folio', 'Referencia', 'Alta', 'Total', 'UUID'], 'IMSS_2025': ['Folio', 'Referencia', 'Alta', 'Total', 'UUID']} 
    paq_folder = r"C:\Users\arman\Dropbox\3. Armando Cuaxospa\Adjudicaciones\Licitaciones 2025\E115 360\Implementación\Facturas\IMSS"
    altas_path = r"C:\Users\arman\Dropbox\3. Armando Cuaxospa\Adjudicaciones\Licitaciones 2025\E115 360\Implementación\SAI\Ordenes_altas.xlsx"
    info_types = 'IMSS'
    altas_sheet = 'df_altas'
    xlsx_database = os.path.join(folder_root, "Implementación", "Facturas", 'xmls_extraidos.xlsx')

    validacion_de_paqs(dict_path_sheet, dic_columnas, paq_folder, altas_path, altas_sheet, info_types, xlsx_database)
