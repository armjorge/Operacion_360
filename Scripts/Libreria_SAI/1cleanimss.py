import pandas as pd
import re
import glob
import os

# Define a function to get the most recently modified file with a certain prefix
def get_latest_file(prefix):
    files = glob.glob(f"{prefix}*.xlsx")
    if not files:
        raise FileNotFoundError(f"No file with prefix {prefix} found")
    latest_file = max(files, key=os.path.getctime)
    return latest_file
    
# Load the excel file
df = pd.read_excel(get_latest_file('ordenes_export_'))

# Define the pattern for the format ###.###.####.##
pattern = r'^\d{3}\.\d{3}\.\d{4}\.\d{2}$'

# Apply the operations if the format does not match
df['cveArticulo'] = df['cveArticulo'].apply(lambda x: re.sub(' ', '.', x[:-3]) if not re.match(pattern, x) else x)

# Save the dataframe back to the excel file
df.to_excel(get_latest_file('ordenes_export_'), index=False)

# Load the excel file for Altas.xlsx
df_altas = pd.read_excel(get_latest_file('altas_export_'))

# Remove trailing whitespace from descUnidad column
df_altas['descUnidad'] = df_altas['descUnidad'].str.rstrip()

# Apply the operations if the format does not match
df_altas['clave'] = df_altas['clave'].apply(lambda x: re.sub('-', '.', x[:-3]) if not re.match(pattern, x) else x)

# Save the dataframe back to the excel file
df_altas.to_excel(get_latest_file('altas_export_'), index=False)
