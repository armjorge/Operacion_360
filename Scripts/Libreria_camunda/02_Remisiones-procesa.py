import pandas as pd
import re

# Read the CSV file
df = pd.read_csv('./INSABI_Remisiones 2024/extraccionRemisionesINSABI.csv')
df = df.dropna(subset=['Tabla_lotes'])

# Define a function to extract the desired data
def extract_data(cell_content):
    # Regular expression to find the pattern: characters + date + date + number
    pattern = re.compile(r'([A-Za-z0-9]+)\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+(\d{1,2}/\d{1,2}/\d{2,4}).*')
    matches = pattern.finditer(cell_content)
    
    batches = []
    expiry_dates = []
    manufacturing_dates = []
    pieces = []
    
    prev_end = 0
    for match in matches:
        batches.append(match.group(1))
        expiry_dates.append(match.group(2))
        manufacturing_dates.append(match.group(3))
        
        # Extract the piece value from the previous match end to the current match start
        piece_part = cell_content[prev_end:match.start()]
        piece_match = re.search(r'(\d+)(?=\D*$)', piece_part)
        if piece_match:  # Only append if there's a match
            pieces.append(piece_match.group(1))
        
        prev_end = match.end()
    
    # Extract the piece value from the last match end to the end of the string
    piece_part = cell_content[prev_end:]
    piece_match = re.search(r'(\d+)(?=\D*$)', piece_part)
    if piece_match:  # Only append if there's a match
        pieces.append(piece_match.group(1))
    
    # Join multiple entries with newline character   
    return (
        ", ".join(batches),
        ", ".join(expiry_dates),
        ", ".join(manufacturing_dates),
        ", ".join(pieces)  # This should now correctly exclude initial empty strings
    )
    """
    # Línea original, esto agrega un Char(10) en las celdas encontradas
    # Join multiple entries with newline character   
    return (
        "\n".join(batches),
        "\n".join(expiry_dates),
        "\n".join(manufacturing_dates),
        "\n".join(pieces)
    )
"""
# Apply the function and create new columns
df[['Batch', 'Expiry_date', 'Manufacturing_date', 'Pieces']] = pd.DataFrame(
    df['Tabla_lotes'].apply(extract_data).tolist(), 
    index=df.index
)

# Select columns from A to H and the new columns
df = df.iloc[:, :8].join(df[['Batch', 'Expiry_date', 'Manufacturing_date', 'Pieces']])

# Save the DataFrame to a new XLSX file
#df.to_excel('remisiones_insabi2024.xlsx', index=False, engine='openpyxl')

# Path to your existing workbook
file_path = 'INSABI_ordenes-contratos-sellos-remisiones.xlsx'


# Initialize the ExcelWriter object in append mode
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    # Access the OpenPyXL workbook object
    workbook = writer.book
    
    # Check if the 'Remisiones' sheet exists, and remove it
    if 'Remisiones' in workbook.sheetnames:
        del workbook['Remisiones']
    
    # Now, write your DataFrame to the 'Remisiones' sheet in the workbook
    # This operation will overwrite the 'Remisiones' sheet with your DataFrame
    df.to_excel(writer, sheet_name='Remisiones', index=False)
    
print(f"\n*******************************\n Resultado guardado a: {file_path} Remisiones \n*******************************\n")
