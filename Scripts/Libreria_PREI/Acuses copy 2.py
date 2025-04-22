import PyPDF2
import openpyxl
import glob
import re
import os

print("\n*******************************\n Renombre de acuses de carga al PREI \n*******************************\n")
def read_pdf(file_path):
    with open(file_path, 'rb') as pdf_file_obj:
        pdf_reader = PyPDF2.PdfReader(pdf_file_obj)
        text = ''

        for page in range(len(pdf_reader.pages)):
            page_obj = pdf_reader.pages[page]
            text += page_obj.extract_text()

    return text

def extract_p_number(text):
    match = re.search(r'P\d{4}', text)
    if match:
        # Add a "-" after the "P" in the extracted value
        return match.group(0)[:1] + '-' + match.group(0)[1:]
    else:
        return 'Not found'

# Create a new workbook and select the active shee
wb = openpyxl.Workbook()
ws = wb.active

# Loop over all PDF files in the current directory
for file_path in glob.glob('./*.pdf'):
    # Read the PDF and extract the P#### number
    text = read_pdf(file_path)
    p_number = extract_p_number(text)

    # Append the file name and P#### number to the Excel sheet
    ws.append([file_path, p_number])

    # Rename the PDF file based on its extracted value, adding "_" before the extension
    new_file_path = os.path.join(os.path.dirname(file_path), p_number + '_.pdf')
    os.rename(file_path, new_file_path)

# Save the Excel file
wb.save('output.xlsx')
