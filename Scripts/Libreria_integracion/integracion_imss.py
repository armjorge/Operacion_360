import sys 
import os
import pandas as pd

def integracion_IMSS(IMSS_PREI, IMSS_altas, IMSS_Facturas, xlsx_database, multi_column_lookup):
    # Carga de archivos Excel
    altas_sheet = 'df_altas'
    df_IMSS_facturas = pd.read_excel(IMSS_Facturas)
    df_PREI = pd.read_excel(IMSS_PREI)
    df_IMSS_altas = pd.read_excel(IMSS_altas, sheet_name=altas_sheet)
    df_XMLS = pd.read_excel(xlsx_database)
    df_XMLS = (df_XMLS.drop_duplicates(subset='UUID', keep='first').reset_index(drop=True))

    # PREI con Folio-Serie
    print("Poblando PREI con el Folio-Serie")
    columna_retorno ='Folio'
    columna_poblar = 'Factura'
    columns_to_match = {'Folio Fiscal': 'UUID'}

    print(f"Vamos a generar y poblar el df_PREI columna {columna_poblar} con el df_XMLS la columna {columna_retorno} de consulta filtrando para {columns_to_match}")

    df_PREI[columna_poblar] = multi_column_lookup(
        df_to_fill=df_PREI,
        df_to_consult=df_XMLS,
        match_columns=columns_to_match,
        return_column=columna_retorno,
        default_value=f'{columna_retorno} no localizado'
    )
    df_PREI.to_excel(IMSS_PREI, index=False)
    # Altas con Estatus PREI 
    print("Poblando PREI con el Folio-Serie")
    columna_retorno ='Estado C.R.'
    columna_poblar = 'Estado C.R.'
    columns_to_match = {'Factura': 'Factura'}

    print(f"Vamos a generar y poblar el df_altas columna {columna_poblar} con el df_PREI la columna {columna_retorno} de consulta filtrando para {columns_to_match}")

    df_IMSS_altas[columna_poblar] = multi_column_lookup(
        df_to_fill=df_IMSS_altas,
        df_to_consult=df_PREI,
        match_columns=columns_to_match,
        return_column=columna_retorno,
        default_value=f'{columna_retorno} no localizado'
    )
    with pd.ExcelWriter(IMSS_altas, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_IMSS_altas.to_excel(writer, sheet_name=altas_sheet, index=False)
    
if __name__ == "__main__":
    folder_root = r"C:\Users\arman\Dropbox\3. Armando Cuaxospa\Adjudicaciones\Licitaciones 2025\E115 360"
    IMSS_PREI    = os.path.join(folder_root, "Implementación", "PREI",'PREI.xlsx')
    IMSS_altas   = os.path.join(folder_root, "Implementación", "SAI", 'Ordenes_altas.xlsx')
    IMSS_Facturas= os.path.join(folder_root, "Implementación", "Facturas", 'IMSS', 'IMSS.xlsx')
    xlsx_database = os.path.join(folder_root, "Implementación", "Facturas", 'xmls_extraidos.xlsx')

    # 1) Añade al path la carpeta donde está df_multi_match.py
    libs_dir = os.path.join(folder_root, "Scripts", "Libreria_comunes")
    sys.path.insert(0, libs_dir)

    # 2) Ahora importa la función directamente
    from df_multi_match import multi_column_lookup

    # 3) Llama a tu función pasándola como parámetro
    integracion_IMSS(IMSS_PREI, IMSS_altas, IMSS_Facturas, xlsx_database, multi_column_lookup)

