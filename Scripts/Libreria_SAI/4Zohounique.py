import pandas as pd
import os
import glob

# 0) Delete the file zoho2024.xlsx if it exists
if os.path.exists("zoho2024.xlsx"):
    os.remove("zoho2024.xlsx")

# 1) Grab the files with the prefix "Orden_de_venta" in the same folder of the script
files = glob.glob("Orden_de_venta*.xlsx")

# If no files found, print an error and exit
if not files:
    print("No files found with the prefix 'Orden_de_venta'. Exiting.")
    exit(1)

# 2) Check if all files have the same column headers
dfs = []
for file in files:
    df = pd.read_excel(file, engine='openpyxl')
    dfs.append(df)

# Get the columns from the first DataFrame to compare with others
first_file_columns = dfs[0].columns

for df in dfs[1:]:
    if not all(df.columns == first_file_columns):
        print("Error: Not all files have the same column headers. Exiting.")
        exit(1)

# 3) Append the data from all files into one
result = pd.concat(dfs, ignore_index=True)

# Print the first 5 rows of the concatenated data for inspection
print("Concatenated Data (First 5 Rows):")
print(result.head())

# Convert the date format in the fechaExpedicion column
if "fechaExpedicion" in result.columns:
    result['fechaExpedicion'] = pd.to_datetime(result['fechaExpedicion']).dt.strftime('%d/%m/%Y')

# 4) Filter out rows where 'Institución Homologada' is not one of the specified values
allowed_values = ["CCINSHAE", "OADPRS", "ISSSTE", "SPPS-SAP"]
result_filtered = result[result['Institución Homologada'].isin(allowed_values)]

# Print the first 5 rows of the filtered data for inspection
print("\nFiltered Data (First 5 Rows):")
print(result_filtered.head())

# Save the concatenated data to zoho2024.xlsx
result_filtered.to_excel("zoho2024.xlsx", index=False, engine='openpyxl')

# 5) Print "file created"
print("file created")

# 6) Grab the values from the active sheet and grab the values of the column with header "SalesOrder Number"
sales_order_numbers = result['SalesOrder Number'].unique()

# 7) Create a new sheet in the original source file with the name "SO24 no duplicadas"
# Assuming the first file in the list is the original source file
original_file = pd.ExcelWriter(files[0], engine='openpyxl', mode='a')
pd.DataFrame(sales_order_numbers, columns=["Orden Zoho"]).to_excel(original_file, sheet_name="SO24 no duplicadas", index=False)

# 8) Save the original source file and print the message "Original file ready to be invoiced"
original_file.save()
print("Original file ready to be invoiced")

# 9) End
