import re
import os
import pandas as pd
from PyPDF2 import PdfReader
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe



def extract_info_from_pdf(file_path):
    pdf_file_obj = open(file_path, 'rb')
    pdf_reader = PdfReader(pdf_file_obj)
    
    contrato_info = []
    orden_info = []

    for page in pdf_reader.pages:
        text = page.extract_text()

        # Search for contrato and orden patterns
        contrato_match = re.search(r'Contrato Procedimiento Fianza Partida presupuestal\n(.*?) ', text)
        orden_match = re.search(r'NÚMERO DE ORDEN DE SUMINISTRO:\n(.*?)\n', text)

        contrato = contrato_match.group(1) if contrato_match else ""
        orden = orden_match.group(1) if orden_match else ""

        contrato_info.append(contrato)
        orden_info.append(orden)

    pdf_file_obj.close()
    return contrato_info, orden_info

def clean_contrato_value(contrato):
    # Remove the pattern "-***" at the end of the string, if present
    #return re.sub(r'-\w{3}$', '', contrato)
    contrato = re.sub(r'-\w{3}$', '', contrato)
    contrato = re.sub(r'-HRAEPY$', '', contrato)
    contrato = re.sub(r'-INER$', '', contrato)
    return contrato
    
def main():
    #directory = r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Dataframes\Camunda\INSABI_Órdenes 2024"  # Replace with your directory path
    root = os.path.dirname(os.path.abspath(__file__))
    # Build file_path
    directory = os.path.join(root, 'INSABI_Órdenes 2024')

    files = [f for f in os.listdir(directory) if f.endswith('.pdf')]

    all_contrato = []
    all_orden = []

    for file in files:
        print(directory, '\n')
        contrato, orden = extract_info_from_pdf(os.path.join(directory, file))
        cleaned_contrato = [clean_contrato_value(c) for c in contrato]
        all_contrato.extend(cleaned_contrato)
        all_orden.extend(orden)

    df = pd.DataFrame({
        'Contrato': all_contrato,
        'Orden': all_orden
    })

    #df.to_excel('INSABI_ordenes-contratos.xlsx', index=False)
    #file_path = r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Dataframes\Camunda\INSABI_ordenes-contratos-sellos-remisiones.xlsx"
    file_path = os.path.join(root, "INSABI_ordenes-contratos-sellos-remisiones.xlsx")

    # Use ExcelWriter in append mode to open the workbook
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        # Access the workbook via the writer
        workbook = writer.book
        
        # If 'Contratos' sheet exists, remove it
        if 'Contratos' in workbook.sheetnames:
            del workbook['Contratos']
        
        # Write the DataFrame to a new 'Contratos' sheet
        df.to_excel(writer, sheet_name='Contratos', index=False)
    print(f"\n*******************************\n Hoja 'Contratos' actualizada {file_path} \n*******************************\n")
    
    scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
    # Add your service account file
    creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)  # Ensure the correct path
    # Authorize the client sheet
    client = gspread.authorize(creds)
    # Get the instance of the Spreadsheet
    spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1KN4XwXQlZ5jhyErxdpA_R6kNk2tKwwFeDnCDuHCr0fo/edit#gid=2033397596')
    # Access the specific worksheet 'PagosKM'
    gsheet_contratosINSABI = spreadsheet.worksheet('Contratos_INSABI')
    # Clear existing data
    gsheet_contratosINSABI.clear()
    # Upload the DataFrame to the worksheet
    #df_2023 = pd.read_excel(r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Dataframes\Camunda\INSABI_Órdenes 2024\INSABI_ordenes-contratos2023.xlsx")
    file_2023 = os.path.join(root, "INSABI_Órdenes 2024", "INSABI_ordenes-contratos2023.xlsx")
    df_2023 = pd.read_excel(file_2023) 
    # Concatenate the two DataFrames vertically
    df_2023_2024 = pd.concat([df, df_2023], ignore_index=True)
    set_with_dataframe(gsheet_contratosINSABI, df_2023_2024)  # Ensure 'merged_df' is your DataFrame variable
    print(f"\n*******************************\n {file_path} Contratos actualizados en el google sheet {gsheet_contratosINSABI}\n*******************************\n")
    
    # Load the 'Órdenes' sheet, columns A:O from the Excel file into a DataFrame
    #df_ordenes = pd.read_excel(r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Dataframes\Camunda\INSABI_ordenes-contratos-sellos-remisiones.xlsx", sheet_name='Órdenes', usecols='A:O')
    file_ordenes = os.path.join(root, "INSABI_ordenes-contratos-sellos-remisiones.xlsx")
    df_ordenes = pd.read_excel(file_ordenes, sheet_name='Órdenes', usecols='A:O')

    # Remove from df_ordenes those rows where 'NÚMERO DE ORDEN DE SUMINISTRO' exists in df['Orden']
    df_missingContracts = df_ordenes[~df_ordenes['NÚMERO DE ORDEN DE SUMINISTRO'].isin(df['Orden'])]
    # Print the head of the resulting DataFrame
    #print(df_missingContracts)
    # Concatenate the two columns into a new Series
    concatenated = "Estatus Camunda: " + df_missingContracts['ESTATUS'] + " Fecha: "+  df_missingContracts['FECHA EXPEDICIÓN DE LA ORDEN'].astype(str)
    # Get unique values
    unique_values = concatenated.unique()
    unique_values=unique_values.tolist()

    print("\n*******************************\nDescarga las siguientes órdenes, en el estatus y fecha que se indican: ")
    for i, value in enumerate(unique_values):
        print(f"{value}")
    print("*******************************\n")
    # Display the unique values
    
    
     
if __name__ == "__main__":
    main()