import pandas as pd
import os

# Define the path to your source Excel file and the new Excel file
input_excel = ".\\INSABI_ordenes-contratos-sellos-remisiones.xlsx"
output_excel = ".\\info consola INSABI.xlsx"

# List of sheet names you are interested in
sheets_to_copy = ['Contratos', 'Sellos', 'Remisiones', 'Órdenes']

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
    for sheet_name in sheets_to_copy:
        # Read each specified sheet from the source Excel file
        df = pd.read_excel(input_excel, sheet_name=sheet_name)
        # Write the DataFrame to the new Excel file
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"New Excel file '{output_excel}' has been created with the specified sheets.")
os.startfile('.')
print(f"New Excel file '{output_excel}' has been created with the specified sheets.")
