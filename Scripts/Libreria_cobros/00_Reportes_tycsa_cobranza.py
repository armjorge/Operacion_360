import os
import pandas as pd
import numpy as np
import re

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1KN4XwXQlZ5jhyErxdpA_R6kNk2tKwwFeDnCDuHCr0fo/edit#gid=2033397596')

def get_latest_file(directory):
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    paths = [os.path.join(directory, f) for f in files]
    latest_file = max(paths, key=os.path.getmtime)
    return latest_file

def main():
    directory = ".\\Reportes Rogelio"
    latest_file = get_latest_file(directory)
    
    user_input = input(f"¿Usamos este archivo? es el más reciente\n{os.path.basename(latest_file)}? \n(yes/no): ").strip().lower()
    
    if user_input == 'yes':
        excel_to_clean = latest_file
    else:
        excel_to_clean = input("Please input the filename to process (with extension): ").strip()
        excel_to_clean = os.path.join(directory, excel_to_clean)
    
    headers_user = int(input("Númer de filas antes de que el encabezado empiece (e.g., 4): ").strip())
    sheet_user = input("Please enter the name of the sheet to extract: ").strip()
    
    df_toclean = pd.read_excel(excel_to_clean, sheet_name=sheet_user, skiprows=headers_user)
    df_toclean.columns = [re.sub(r'[^a-zA-Z0-9_]+', '_', col) for col in df_toclean.columns]
    
    print("DataFrame head:\n", df_toclean.head())
    
    continue_process = input("Do you want to continue? (yes/no): ").strip().lower()
        
    if continue_process != 'yes':
        print("Process terminated.")
        return
    # Empezamos con el código para generar una fecha de ingreso basada en la lógica
        # Si ya existe fecha de pago y no existen fechas de ingreso, quítale 20 días
        # Si la fecha de pago pago es menor a la fecha de ingreso, bórrala y toma el pago menos 20 días
        # Si la fecha de pago es mayor a la fecha de ingreso, no hagas nada. 
        # La fecha de ingreso es la que más temprana de las 3 columnas de Fecha (sin incluir fecha de pago). 
        
    df_final = pd.DataFrame(columns=['Factura', 'Importe', 'Fecha de ingreso', 'Alta', 'Orden', 'Fecha de pago'])

    # Print headers starting with 'Fecha'
    date_columns = [col for col in df_toclean.columns if col.startswith('Fecha')]
    print(f"Columnas de fecha: {date_columns}")

    # Ask for user input for the date columns
    first_date_col = input("Columna entrega a gestor: ").strip()
    second_date_col = input("Columna ingreso a instituto: ").strip()
    third_date_col = input("Fecha CR: ").strip()
    payment_date_col = input("Fecha pago: ").strip()

    # Rename the original columns
    df_toclean.rename(columns={
        first_date_col: '1st_date',
        second_date_col: '2nd_date',
        third_date_col: '3e_date',
        payment_date_col: 'payment_date'
    }, inplace=True)

    # Clean the date columns
    for col in ['1st_date', '2nd_date', '3e_date', 'payment_date']:
        df_toclean[col] = pd.to_datetime(df_toclean[col], errors='coerce')

    # Calculate 'Fecha Ingreso a cobro'
    df_toclean['Fecha Ingreso a cobro'] = df_toclean[['1st_date', '2nd_date', '3e_date']].min(axis=1)
    
    # Adjust 'Fecha Ingreso a cobro' based on 'payment_date'
    payment_valid = df_toclean['payment_date'].notna()
    cobro_invalid = df_toclean['Fecha Ingreso a cobro'].isna()

    df_toclean.loc[payment_valid & cobro_invalid, 'Fecha Ingreso a cobro'] = df_toclean['payment_date'] - pd.Timedelta(days=20)
    condition = df_toclean['Fecha Ingreso a cobro'] >= df_toclean['payment_date']
    df_toclean.loc[condition, 'Fecha Ingreso a cobro'] = df_toclean['payment_date'] - pd.Timedelta(days=20)
    
    print("Fechas corregidas\n")
    print("DataFrame head:\n", df_toclean.head())

    # Ahora a mapear el resto de las columnas
    print("Columnas del dataframe:\n", df_toclean.columns.tolist())

    # Ask for user input for the other columns
    factura_col = input("Columna Factura: ").strip()
    importe_col = input("Columna Importe: ").strip()
    alta_col = input("Columna Alta: ").strip()
    orden_col = input("Columna Orden: ").strip()

    # Rename the original columns
    df_toclean.rename(columns={
        factura_col: 'Factura',
        importe_col: 'Importe',
        alta_col: 'Alta',
        orden_col: 'Orden'
    }, inplace=True)

    # Map to the standardized dataframe
    df_final = pd.DataFrame(columns=['Factura', 'Importe', 'Fecha de ingreso', 'Alta', 'Orden', 'Fecha de pago'])
    df_final['Factura'] = df_toclean['Factura']
    df_final['Importe'] = df_toclean['Importe']
    df_final['Fecha de ingreso'] = df_toclean['Fecha Ingreso a cobro']
    df_final['Alta'] = df_toclean['Alta']
    df_final['Orden'] = df_toclean['Orden']
    df_final['Fecha de pago'] = df_toclean['payment_date']

    print("Final DataFrame head:\n", df_final.head())
    duplicates = df_final[df_final.duplicated('Factura', keep=False)]
    
    # Aquí empieza la lógica para eliminar duplicadas del dataframe limpio, procesado. 
    removed_rows = pd.DataFrame(columns=df_final.columns)
    
    if not duplicates.empty:
        grouped = duplicates.groupby('Factura')
        for name, group in grouped:
            if group['Importe'].nunique() == 1:
                valid_payment_date = group['Fecha de pago'].notna()
                if valid_payment_date.any():
                    to_keep = group[valid_payment_date].iloc[0]
                else:
                    to_keep = group.loc[group['Fecha de ingreso'].idxmin()]
                to_remove = group.drop(to_keep.name)
                removed_rows = pd.concat([removed_rows, to_remove])
                df_final = df_final.drop(to_remove.index)

    print("Removed rows summary:\n", removed_rows)
    print("Updated DataFrame head:\n", df_final.head())
    # Save the final dataframe to excel
    output_path = os.path.join(directory, f"{os.path.splitext(os.path.basename(excel_to_clean))[0]}_estandarized.xlsx")
    df_final.to_excel(output_path, index=False)
    print(f"Standardized DataFrame saved to {output_path}")

    
    if df_final is not None:
        print("\n*******************************\nReemplazando información en google sheet\n*******************************\n")
        gsheet_REPORTEtycsaCOBROS = spreadsheet.worksheet('Reporte TYCSA')
        gsheet_REPORTEtycsaCOBROS.clear()
        set_with_dataframe(gsheet_REPORTEtycsaCOBROS, df_final)
        print("\n*******************************\nReemplazadas\n*******************************\n")
    else:
        print("\n*******************************\nNo hay datos para reemplazar en el google sheet\n*******************************\n")
if __name__ == "__main__":
    main()