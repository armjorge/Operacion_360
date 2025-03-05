import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe
import pandas as pd
import warnings

# Suppress specific warning
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl.worksheet._reader', message="Conditional Formatting extension is not supported and will be removed")

# Path to the source Excel file
source_file_path = r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Dataframes\STATUS ESEOTRES.xlsx"
# Output file path
output_file_path = r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Dataframes\INSABI_tycsa.xlsx"


# Open the Excel file and read the specified sheet
df = pd.read_excel(source_file_path, sheet_name='Concentrado', header=2)

# Filter rows where Instituto is either 'INSABI' or 'IMSS Bienestar'
filtered_df = df[(df['ESTATUS'] == 'Entregado') & (df['INSTITUTO'] != 'IMSS')]

# Select the specified columns
columns_to_keep = [
    'PEDIDO CLIENTE',  # Column E
    'CANTIDAD SOLICITADA',  # Column J
    'FECHA DE ENTREGA AL CLIENTE',  # Column AB
    'ESTATUS',  # Column AE
    'MOTIVO',  # Column AF
    'OBSERVACIONES',  # Column AJ
    'FACTURA',
    'OS- ESEOTRES',
    'REMISIÓN DE ENTREGA REDI',
    'INSTITUTO',
    'CLAVE'
]

selected_data = filtered_df[columns_to_keep]
selected_data_2023 = pd.read_excel(r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Dataframes\LOGISTICA 2023.xlsx")

selected_data = pd.concat([selected_data, selected_data_2023], ignore_index=True)

# Save the filtered data to a new Excel file
selected_data.to_excel(output_file_path, index=False)

print("\n*******************************\nExtracción de fechas realizadas, ahora al GSHEET\n*******************************\n")

# Define the scope
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# Add your service account file
creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)  # Ensure the correct path

# Authorize the client sheet
client = gspread.authorize(creds)

# Get the instance of the Spreadsheet
spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1KN4XwXQlZ5jhyErxdpA_R6kNk2tKwwFeDnCDuHCr0fo/edit#gid=2033397596')

# Access the specific worksheet 'PagosKM'
gsheet_logistica = spreadsheet.worksheet('TYCSA_logistica')

# Clear existing data
gsheet_logistica.clear()

# Upload the DataFrame to the worksheet
set_with_dataframe(gsheet_logistica, selected_data)  # Ensure 'merged_df' is your DataFrame variable

print("\n*******************************\nLogística Actualizada\n*******************************\n")
