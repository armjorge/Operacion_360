import os
import xml.etree.ElementTree as ET
import PyPDF2
import re
from openpyxl import Workbook
import pandas as pd

def extract_xml_data(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Extract MetodoPago
    metodo_pago = root.get('MetodoPago', 'N/A')

    # Extract UUID and Importe
    uuid = 'N/A'
    importe = 'N/A'
    for elem in root.iter():
        if 'CfdiRelacionado' in elem.tag:
            uuid = elem.get('UUID', 'N/A')
        if 'Concepto' in elem.tag:
            importe = elem.get('Importe', 'N/A')
        if 'TimbreFiscalDigital' in elem.tag:
            uuid_nc = elem.get('UUID', 'N/A')
    
    return metodo_pago, uuid, importe, uuid_nc

def extract_info_from_pdf_text(text):
    # Define the patterns to search for
    metodo_pago_pattern = r"Metodo de Pago: (.+)"
    cfdi_relacionados_pattern = r"(?<=01 - Nota de crédito de los documentos relacionados\n).+"
    subtotal_pattern = r"Subtotal: (.+)"

    # Search for patterns in the text
    metodo_pago_match = re.search(metodo_pago_pattern, text)
    cfdi_relacionados_match = re.search(cfdi_relacionados_pattern, text)
    subtotal_match = re.search(subtotal_pattern, text)

    # Extract the matched groups
    metodo_pago = metodo_pago_match.group(1) if metodo_pago_match else "Not found"
    cfdi_relacionados = cfdi_relacionados_match.group(0) if cfdi_relacionados_match else "Not found"  # Changed to group(0) to capture the entire matched string
    subtotal = subtotal_match.group(1) if subtotal_match else "Not found"

    return metodo_pago, cfdi_relacionados, subtotal

def read_pdf(file_name):
    with open(file_name, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"  # Concatenate text from all pages

    return extract_info_from_pdf_text(text)

#if __name__ == "__main__":
#    pdf_file_name = input("Please enter the name of the PDF file: ")
#    try:
#        metodo_pago, cfdi_relacionados, subtotal = read_pdf(pdf_file_name)
#        print(f"Metodo de Pago: {metodo_pago}")
#        print(f"CFDI Relacionados: {cfdi_relacionados}")
#        print(f"Subtotal: {subtotal}")
#   except Exception as e:
#      print(f"An error occurred: {e}")


def adding_total_asINSABI(folder_name):
    path_audit = os.path.join(folder_name, f"{folder_name}_audit.xlsx")
    path_relacion = os.path.join(folder_name, f"{folder_name}_relacion.xlsx")
    
    try:
        df_XML_PDF = pd.read_excel(path_audit)
        df_relacion = pd.read_excel(path_relacion)
    except Exception as e:
        print("Error reading files:", e)
        return

    # Check if 'FOLIO FISCAL FACTURA' exists
    if 'FOLIO FISCAL FACTURA' not in df_relacion.columns:
        print("Headers from df_relacion are:", df_relacion.columns)
        return
    
    # Improved logic to process 'FOLIO FISCAL FACTURA' column
    df_relacion['FOLIO FISCAL FACTURA'] = df_relacion['FOLIO FISCAL FACTURA'].apply(
        lambda x: re.sub(r'[\s\n]+[\d]+$','', x)
    )
    # Clean up unwanted characters without removing essential parts
    df_relacion['FOLIO FISCAL FACTURA'] = df_relacion['FOLIO FISCAL FACTURA'].str.replace("[-\s\n]", "", regex=True)
    df_XML_PDF['XML_CFDIrelacionado'] = df_XML_PDF['XML_CFDIrelacionado'].str.replace("-", "")

    # Debugging: Save intermediate DataFrame to CSV for inspection
    debug_path = os.path.join(folder_name, f"{folder_name}_df_relacion.csv")
    df_relacion.to_csv(debug_path, index=False)
    print(f"Debug CSV has been saved to {debug_path}")

    # Merging dataframes
    df_XML_PDF = df_XML_PDF.merge(df_relacion[['FOLIO FISCAL FACTURA', 'PENA', 'ORDEN DE SUMINISTRO']], left_on='XML_CFDIrelacionado', right_on='FOLIO FISCAL FACTURA', how='left')

    #df_XML_PDF['Pena_oficio'] = df_XML_PDF['PENA']
    df_XML_PDF['PDF_total'] = pd.to_numeric(df_XML_PDF['PDF_total'].str.replace(r'[$,]', '', regex=True), errors='coerce')
    
    df_XML_PDF['XML_Total'] = pd.to_numeric(df_XML_PDF['XML_Total'], errors='coerce')
    #df_XML_PDF['PDF_total'] = pd.to_numeric(df_XML_PDF['PDF_total'], errors='coerce')

    # Create 'Compara total' column
    df_XML_PDF['Compara total'] = df_XML_PDF.apply(
        lambda x: 'Coincide' if x['PENA'] == x['XML_Total'] == x['PDF_total'] else 'No coinciden', axis=1
    )
    # Save the updated dataframe
    df_XML_PDF.to_excel(path_audit, index=False)
    print("Updated file has been saved.")


def process_folder(folder_name):
    wb = Workbook()
    ws = wb.active
    ws.append(["Filename", "XML_metododepago", "XML_CFDIrelacionado", "XML_Total", "XML_UUID_NC", "PDF_metododepago", "PDF_CFDIrelacionado", "PDF_total", "Auditory"])

    data_dict = {}  # Dictionary to store data for each file base name

    for filename in os.listdir(folder_name):
        # Skip files with specific suffixes
        if filename.endswith('_SAT.pdf') or filename.endswith('_TXT.pdf') or filename.endswith('.csv'):
            continue

        file_base_name = os.path.splitext(filename)[0]  # Get the base name of the file without extension

        if file_base_name not in data_dict:
            # Initialize data for this file base name
            data_dict[file_base_name] = ["N/A"] * 8  # Eight 'N/A' for eight data fields

        if filename.endswith('.xml'):
            xml_data = extract_xml_data(os.path.join(folder_name, filename))
            data_dict[file_base_name][:4] = xml_data  # Update the first four elements for XML data
        elif filename.endswith('.pdf'):
            pdf_data = read_pdf(os.path.join(folder_name, filename))
            data_dict[file_base_name][4:7] = pdf_data  # Update the next three elements for PDF data

    # Write data to the worksheet and perform comparison
    for file_base_name, data in data_dict.items():
        xml_metododepago, xml_cfdi, xml_total, xml_uuid_nc = data[:4]
        pdf_metododepago, pdf_cfdi, pdf_total = data[4:7]

        # Remove dollar sign from PDF total for comparison
        pdf_total_clean = pdf_total.replace('$', '').replace(',', '').strip() if pdf_total != "N/A" else "N/A"

        # Compare the first word of PDF_metododepago with XML_metododepago
        pdf_metododepago_first_word = pdf_metododepago.split()[0] if pdf_metododepago != "N/A" else "N/A"
        auditory_result = "Coincide" if (xml_metododepago == pdf_metododepago_first_word and xml_total == pdf_total_clean and xml_cfdi == pdf_cfdi) else "Needs to be reviewed"

        ws.append([file_base_name, xml_metododepago, xml_cfdi, xml_total, xml_uuid_nc, pdf_metododepago, pdf_cfdi, pdf_total, auditory_result])
        #ws.append([file_base_name] + data + [auditory_result])

    path = os.path.join(folder_name, f"{folder_name}_audit.xlsx")
    # Make sure the directory exists, create if it doesn't
    os.makedirs(folder_name, exist_ok=True)
    # Save the workbook
    wb.save(path)


folder_name = input("Enter the folder name: ")
process_folder(folder_name)

while True:
    user_input = input("Do we try to add total? (yes/no): ")
    if user_input.lower() == 'no':
        break
    elif user_input.lower() == 'yes':
        adding_total_asINSABI(folder_name)
        break
    else:
        print("Please enter 'yes' or 'no'.")