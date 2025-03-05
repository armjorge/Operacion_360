import os
import re
import pandas as pd
from PyPDF2 import PdfReader

# Define the folder paths
letter_folder = './Oficios'

# Adjusted pattern for flexibility
pattern = r'{\s*(?P<Sale_order>[^,]*),\s*(?P<Invoice>[^,]+),\s*(?P<Penalty>[^}]+)\s*}'

# PDF extraction function
def pdf_extraction(input_folder, pattern):
    df_consolidation = pd.DataFrame(columns=['Sale_order', 'Invoice', 'Penalty', 'Filename'])
    
    # Loop through each PDF file in the folder
    for file_name in os.listdir(input_folder):
        if file_name.endswith('.pdf'):
            file_path = os.path.join(input_folder, file_name)
            
            # Read the PDF
            with open(file_path, 'rb') as file:
                pdf_reader = PdfReader(file)
                num_pages = len(pdf_reader.pages)
                has_match = False  # To track if any page has a valid pattern or error
                
                # Loop through each page and extract text
                for page_num in range(num_pages):
                    page = pdf_reader.pages[page_num]
                    page_text = page.extract_text()
                    
                    # Check for unbalanced brackets
                    if check_bracket_error(page_text):
                        print(f"{file_name} has an unclosed bracket pair")
                        df_consolidation = df_consolidation._append({
                            'Sale_order': 'Bracket error',
                            'Invoice': 'Bracket error',
                            'Penalty': 'Bracket error',
                            'Filename': file_name
                        }, ignore_index=True)
                        has_match = True  # Mark that this page has been processed with an error
                        break  # Skip further processing of this file

                    # Clean the data and append to the dataframe
                    df_consolidation, match_found = data_cleaning(page_text, file_name, df_consolidation, pattern)
                    if match_found:
                        has_match = True

                # If no pattern or error was found, mark it as 'Not labeled'
                if not has_match:
                    df_consolidation = df_consolidation._append({
                        'Sale_order': 'Not labeled',
                        'Invoice': 'Not labeled',
                        'Penalty': 'Not labeled',
                        'Filename': file_name
                    }, ignore_index=True)
    
    return df_consolidation

# Function to check for bracket errors
def check_bracket_error(page_text):
    lines = page_text.splitlines()
    for line in lines:
        stripped_line = line.strip()
        if stripped_line == '{' or stripped_line == '}':
            return True
    return False

# Data cleaning function
def data_cleaning(page_text, filename, df_extraction, pattern):
    match_found = False
    # Search for the pattern in the page text
    matches = re.findall(pattern, page_text)
    
    # If matches are found, append to the dataframe
    if matches:
        for match in matches:
            sale_order, invoice, penalty = match
            
            # Clean up the penalty by removing $, commas, and spaces
            penalty_cleaned = re.sub(r'[$, ]', '', penalty)
            
            try:
                penalty_value = float(penalty_cleaned)
            except ValueError:
                penalty_value = None  # Handle the case where the value can't be converted to a float
            
            # Append to the dataframe
            df_extraction = df_extraction._append({
                'Sale_order': sale_order.strip() if sale_order else None,  # Remove leading/trailing spaces, if any
                'Invoice': invoice.strip(),
                'Penalty': penalty_value,
                'Filename': filename
            }, ignore_index=True)
        match_found = True
    return df_extraction, match_found

# Main function to run the extraction and save to Excel
def main():
    # Extract data from letters
    df_from_letters = pdf_extraction(letter_folder, pattern)
    
    # Export the dataframe to Excel
    output_file = './extracción.xlsx'
    df_from_letters.to_excel(output_file, index=False)
    print(f'Data successfully extracted and saved to {output_file}')

# Run the main function
if __name__ == '__main__':
    main()
