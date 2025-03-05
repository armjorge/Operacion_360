import os
import re
import pandas as pd
import PyPDF2
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

# Output file
output_file = './Resumen.xlsx'

# Folder containing validated e5 files
e5_folder = './e5 validados'

def e5_total_extract(file_list):
    extracted_totals = []

    for file_e5 in file_list:
        print(f"Extracting total from {file_e5}")
        with open(file_e5, 'rb') as file:
            # Create a PDF reader object
            pdf_reader = PyPDF2.PdfReader(file)
            
            # Loop through each page and extract text
            for page in pdf_reader.pages:
                text = page.extract_text()
                print(f"Original extracted text from {file_e5}:\n{text}")
                # Normalize text by removing unnecessary spaces
                normalized_text = re.sub(r'\s+', '', text)
                print(f"Normalized text from {file_e5}:\n{normalized_text}")
                # Search for the total amount to be paid in normalized text
                match = re.search(r'TOTALAPAGAR\$([0-9,]+)', normalized_text)
                if match:
                    total_amount = match.group(1).replace(',', '')
                    print(f"Found total amount: {total_amount} in {file_e5}")
                    extracted_totals.append({'filename': os.path.basename(file_e5), 'e5_total': int(total_amount)})
                else:
                    print(f"No match found in {file_e5}")

    # Convert the list of dictionaries to a DataFrame
    df_extracted_totals = pd.DataFrame(extracted_totals)
    #df_extracted_totals.to_csv('extracted_totals.csv', index=False)
    return df_extracted_totals
    
def mirror_folder_excel(df_summary, e5_folder):
    files_in_folder = [f for f in os.listdir(e5_folder) if f.endswith('.pdf')]
    new_files = [{'filename': file} for file in files_in_folder if not any(df_summary['filename'].str.contains(file))]
    
    if new_files:
        df_new_files = pd.DataFrame(new_files)
        print(f"We found {len(df_new_files)} new files, do we proceed to add them?")
        proceed = input("Enter 'Yes' to proceed or 'No' to skip: ")
        if proceed.lower() == 'yes':
            df_summary = pd.concat([df_summary, df_new_files], ignore_index=True)
        else:
            print("Operation skipped, check your source files.")
    
    return df_summary

def removing_rows(df_updated, e5_folder):
    files_in_folder = [f for f in os.listdir(e5_folder) if f.endswith('.pdf')]
    missing_files_indices = []

    for index, row in df_updated.iterrows():
        if row['filename'] not in files_in_folder:
            missing_files_indices.append(index)

    if missing_files_indices:
        print(f"The next rows are in the excel but no longer in the file: {', '.join(map(str, missing_files_indices))}")
        proceed = input("Do we remove them? Enter 'Yes' to proceed or 'No' to skip: ")
        if proceed.lower() == 'yes':
            df_updated = df_updated.drop(missing_files_indices)
        else:
            print("Operation skipped, check your source files.")

    return df_updated

def carga_desde_el_GOOGLESHEET(sheets_in_management):
    # Constants for Google Sheets API
    spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1KN4XwXQlZ5jhyErxdpA_R6kNk2tKwwFeDnCDuHCr0fo/edit#gid=2033397596'
    credentials_file = 'key.json'
    
    # Setup Google Sheets API
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
    client = gspread.authorize(creds)
    
    # Open the spreadsheet by URL
    spreadsheet = client.open_by_url(spreadsheet_url)
    
    # Dictionary to store the resulting dataframes
    dfs = {}
    
    # Function to load data from a worksheet into a DataFrame
    def worksheet_to_df(spreadsheet, worksheet_name, columns_range):
        worksheet = spreadsheet.worksheet(worksheet_name)
        data = worksheet.get(columns_range)
        return pd.DataFrame(data[1:], columns=data[0])

    # Iterate over the sheets_in_management dictionary and create dataframes
    for df_name, (sheet_name, columns_range) in sheets_in_management.items():
        try:
            df = worksheet_to_df(spreadsheet, sheet_name, columns_range)
            dfs[df_name] = df
        except Exception as e:
            print(f"An error occurred while loading {sheet_name}: {e}")
            dfs[df_name] = pd.DataFrame()  # Return an empty DataFrame in case of an error
    
    return dfs

def complet_sale_info(df_fromfiles, df_sales):
    # Populate 'Sales_order' column based on the pattern
    df_fromfiles['Sales_order'] = df_fromfiles['filename'].apply(
        lambda x: re.search(r'(SO2[^\s]*)', x).group(1) if re.search(r'(SO2[^\s]*)', x) else 'Not detected'
    )
    
    # List of columns to copy if match
    columns_to_copy_if_match = ['lugarEntrega', 'Orden', 'Importe']
    
    # Initialize columns in df_fromfiles with NaN values
    for col in columns_to_copy_if_match:
        df_fromfiles[col] = pd.NA

    # For each row in df_fromfiles, copy corresponding values from df_sales if Sales_order matches
    for index, row in df_fromfiles.iterrows():
        sales_order = row['Sales_order']
        if sales_order != 'Not detected':
            match_row = df_sales[df_sales['SalesOrder Number'] == sales_order]
            if not match_row.empty:
                for col in columns_to_copy_if_match:
                    df_fromfiles.at[index, col] = match_row.iloc[0][col]
    df_final_with_zoho_data = df_fromfiles
    return df_final_with_zoho_data

def main():    
    hojas_de_la_consola_necesarias = {
        'df_Zoho': ('ZOHO', 'A:W')#,'df_anotherSheet': ('Another_Sheet_Name', 'A:B') 
        }
    dataframes = carga_desde_el_GOOGLESHEET(hojas_de_la_consola_necesarias)    
    df_Zoho = dataframes['df_Zoho']
    
    if os.path.exists(output_file):
        df_inputed = pd.read_excel(output_file, sheet_name='e5_management')
    else:
        print("Excel not found")

    # Update the summary DataFrame
    df_files_in_folder = mirror_folder_excel(df_inputed, e5_folder)
    df_final = removing_rows(df_files_in_folder, e5_folder)  
    
    file_list = df_final[df_final['e5_total'].isna()]['filename'].tolist()
    print(f"The following are missing {file_list}")
    if file_list:
        df_totales = e5_total_extract([os.path.join(e5_folder, file) for file in file_list])

        if not df_totales.empty:
            # Merge df_totales into df_final based on the 'filename' column
            df_final = pd.merge(df_final, df_totales, on='filename', how='left', suffixes=('', '_new'))

            # Update 'e5_total' with the new values from df_totales
            df_final['e5_total'] = df_final['e5_total'].fillna(df_final['e5_total_new'])
            df_final.drop(columns=['e5_total_new'], inplace=True)
    # a partir de aquí se completan con la info del zoho 
    
    df_final_zoho = complet_sale_info(df_final, df_Zoho)
    df_final_zoho['Importe'] = pd.to_numeric(df_final_zoho['Importe'], errors='coerce')
    df_final_zoho['Porcentaje_sancion'] = (df_final_zoho['e5_total'] / df_final_zoho['Importe']).round(2).apply(lambda x: f"{x:.2%}")

    # Save the final DataFrame back to the Excel file
    with pd.ExcelWriter(output_file, mode='w', engine='openpyxl') as writer:
        df_final_zoho.to_excel(writer, sheet_name='e5_management', index=False)
    
   
    
    
if __name__ == "__main__":
    main()