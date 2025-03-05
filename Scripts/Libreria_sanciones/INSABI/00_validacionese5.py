import os
import PyPDF2
import re
import pandas as pd
import math

def read_pdf(file_name, text_to_remove_file):
    # Load the text inside the txt file
    with open(text_to_remove_file, 'r', encoding='utf-8') as f:
        text_to_remove = f.readlines()
        # Strip any leading/trailing whitespace characters from each line
        text_to_remove = [line.strip() for line in text_to_remove]

    # Create a PDF reader object
    with open(file_name, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        
        # Get the total number of pages in the PDF
        num_pages = len(pdf_reader.pages)
        
        # Open the output file in write mode
        with open("extracted.txt", "w", encoding="utf-8") as output_file:
            # Loop through each page and extract text
            for page_num in range(num_pages):
                page = pdf_reader.pages[page_num]
                page_text = page.extract_text()
                
                # Remove dollar signs
                page_text = page_text.replace('$', '')
                page_text = re.sub(r'[ \t]+$', '', page_text, flags=re.MULTILINE)

                page_text = page_text.replace('DENOMINACIÓN O RAZÓN SOCIAL', '')
                page_text = page_text.replace('CLAVE DEPENDENCIA', '')
                page_text = page_text.replace('REGISTRO FEDERAL DE CONTRIBUYENTES', '')
                
                for line in text_to_remove:
                    # Use regular expression to match whole lines exactly
                    pattern = re.compile(r'^' + re.escape(line) + r'$', re.MULTILINE)
                    page_text = pattern.sub('', page_text)
                
                # Remove empty lines
                page_text = '\n'.join([line for line in page_text.split('\n') if line.strip() != ''])
                #page_text = page_text.replace(',', '')
                    
                # Write the modified text to the output file
                output_file.write(f"Page {page_num + 1}:\n{page_text}\n{'-'*40}\n")

                # Print the modified text to the console
                print(f"Page {page_num + 1}:\n{page_text}\n{'-'*40}")
                
                print(f"Page {page_num + 1}:\n{page_text}\n{'-'*40}")
    return page_text

def auditory_dataframe(cleantext):
    # Create an empty DataFrame with specified headers
    df_master = pd.DataFrame(columns=['RFC', 'RazonSocial', 'Total', 'Ejercicio fiscal', 'Referencia', 'Cadena_dependencia'])
    
    # Search and extract data based on patterns
    rfc_pattern = re.search(r"EPH161215NS9", cleantext)
    rfc = rfc_pattern.group(0) if rfc_pattern else ''

    razon_social_pattern = re.search(r"ESEOTRES PHARMA", cleantext)
    razon_social = razon_social_pattern.group(0) if razon_social_pattern else ''

    total_pattern = re.search(r"TOTAL A PAGAR (\d{1,3}(,\d{3})*)", cleantext)
    total = total_pattern.group(1) if total_pattern else ''

    ejercicio_pattern = re.search(r"EJERCICIO:(\d{4})", cleantext)
    ejercicio = ejercicio_pattern.group(1) if ejercicio_pattern else ''

    referencia_pattern = re.search(r"REFERENCIA(\d+)", cleantext)
    referencia = referencia_pattern.group(1) if referencia_pattern else ''

    dependencia_pattern = re.search(r"DEPENDENCIA(\d+)", cleantext)
    dependencia = dependencia_pattern.group(1) if dependencia_pattern else ''
    
    # Append extracted data to the DataFrame
    df_master.loc[0] = [rfc, razon_social, total, ejercicio, referencia, dependencia]
    df_master['Total'] = df_master['Total'].str.replace(',', '').astype(float).astype(int)
    print("\n*******************************\nDataframe extraído\n")
    print(df_master)
    return df_master

def process_files_in_directory(folder_name):
    # Construct the expected PDF file name
    expected_file_name = f"{folder_name}_e5.pdf"
    
    # Search for the file in the specified folder
    for root, dirs, files in os.walk(folder_name):
        for file in files:
            if file == expected_file_name:
                return os.path.join(root, file)
    
    print(f"Expected e5 file was not found in {folder_name}")
    return None

def comparison(df_extracted, folder_name):
    # Load client specifications from the Excel file
    df_client_specs = pd.read_excel('.\\df_datos_e5.xlsx')    
    df_client_specs['Referencia'] = df_client_specs['Referencia'].astype(str)
    df_tocompare = pd.DataFrame(columns=['Instituto', 'RFC', 'RazonSocial', 'Referencia', 'Cadena_dependencia', 'Total', 'Ejercicio fiscal'])
    
    print("\nSome questions to build the expected dataframe\n")    
    # Ask the user for input
    while True:
        try:
            ejercicio_fiscal = input("Year of the orders (Tax year): ")
            if not re.match(r'^\d{4}$', ejercicio_fiscal):
                raise ValueError("Invalid year format. Please enter a 4-digit year.")
            break
        except ValueError as e:
            print(e)

    while True:
        try:
            total = input("Total of the penalty: ")
            total = float(total)
            # Round up logic based on the specified rules
            total = math.ceil(total) if total % 1 >= 0.51 else math.floor(total)
            break
        except ValueError:
            print("Invalid number format. Please enter a valid number.")
    cadena_dependencia_year = input("Year of the penalties letter issue date: ")
    cadena_dependencia = cadena_dependencia_year + folder_name
    
    # Ask the user for the client name
    print(df_client_specs['Instituto'])
    client_name = input("\nClient name? ")

    # Search for the client name in df_client_specs['Instituto']
    client_row = df_client_specs[df_client_specs['Instituto'] == client_name]
    
    if not client_row.empty:
        # Grab the corresponding values and update df_tocompare
        df_tocompare.loc[0] = {
            'Instituto': client_name,
            'RFC': client_row['RFC'].values[0],
            'RazonSocial': client_row['RazonSocial'].values[0],
            'Referencia': str(client_row['Referencia'].values[0]),
            'Cadena_dependencia': cadena_dependencia,
            'Total': total,
            'Ejercicio fiscal': ejercicio_fiscal
        }
    else:
        print(f"Client name '{client_name}' not found in the client specifications.")
    
    # Ensure 'Total' is an integer
    df_tocompare['Total'] = df_tocompare['Total'].astype(int)
    
    # Print the expected dataframe
    print("\nExpected dataframe")
    #print(df_tocompare)
    print("it works")
    #print(df_extracted)
    df_audit = pd.DataFrame(columns=['Column', 'Dataframe extraído', 'Dataframe esperado'])

    # Populate the audit DataFrame
    for column in df_extracted.columns:
        df_audit.loc[len(df_audit)] = {
            'Column': column,
            'Dataframe extraído': df_extracted.iloc[0][column],
            'Dataframe esperado': df_tocompare.iloc[0][column]
        }

    print(df_audit)
    output_path = os.path.join(folder_name, f"{folder_name}_e5audit.xlsx")
    df_audit.to_excel(output_path, index=False)

    # Print confirmation message
    print(f"\n{folder_name}_e5audit.xlsx saved to {folder_name}")

def main():
    folder_name = input("Enter the folder name containing XML files: ")
    pdf_file_name = process_files_in_directory(folder_name)
    if pdf_file_name:
        text_to_analyze = read_pdf(pdf_file_name, "Texte5.txt")
    df_extracted = auditory_dataframe(text_to_analyze)
    comparison(df_extracted,folder_name)
    print("\n*******************************\nEND\n*******************************\n")
if __name__ == "__main__":
    main()
