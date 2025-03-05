from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
from io import StringIO  # Import StringIO
from datetime import datetime  # Import datetime


# Read the Excel file
df = pd.read_excel('input.xlsx')


today_date = datetime.today().date()
today_mm_dd = today_date.strftime("%m %d")  # Format as "mm dd"

# Ask the user for the download set
downloaded_set = input("\n Descargamos facturas \n1)2023-2024 o \n2)2024?\n")
if downloaded_set == "1":
    downloaded_set = "2023-2024"
elif downloaded_set == "2":
    downloaded_set = "2024"
else:
    print("Invalid input. Please enter 1 or 2.")
    exit()
print(f"\n*******************************\nDescarga el\n{downloaded_set}\n*******************************\n")
chrome_options = webdriver.ChromeOptions()

# Path to the Chrome executable
chrome_binary_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome-win64\\chrome.exe"
chrome_options.binary_location = chrome_binary_path

# Set the default download directory to the specified directory
prefs = {
    #"download.default_directory": os.path.abspath(directory),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,  # Ensures directory creation
    "plugins.always_open_pdf_externally": True,  # Automatically open PDFs externally to prevent Chrome's viewer interference
}
chrome_options.add_experimental_option("prefs", prefs)

# Additional Chrome options for better automation handling
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-popup-blocking")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920x1080")

# Create a new instance of the Chrome driver with the given options
driver = webdriver.Chrome(options=chrome_options)

# Define XPaths for the elements
elements_xpaths = {
    'usuario': "/html/body/div[1]/div/div[2]/div/div/form[1]/div[3]/label/div/div[1]/div/input",
    'pass': "/html/body/div[1]/div/div[2]/div/div/form[1]/div[4]/label/div/div[1]/div[1]/input",
    'access': "/html/body/div[1]/div/div[2]/div/div/form[1]/div[5]/button",
    'cuadro_ordenes': "/html/body/div[1]/div/div[2]/div/div[2]/div/div[1]/div/div[2]/label/div/div/div[1]/input",
    'validacion': "/html/body/div[1]/div/div[2]/div/div[2]/div/div[2]/table/tbody/tr/td[11]",
    'numero_orden': "/html/body/div[1]/div/div[2]/div/div[2]/div/div[2]/table/tbody/tr[1]/td[6]",
    'click_cancelacion': "/html/body/div[1]/div/div[2]/div/div[2]/div/div[2]/table/tbody/tr[1]/td[12]/div/button[4]",
    'continuar_button': "/html/body/div[3]/div[2]/div/div[2]/button[2]",
    'clear_button': "/html/body/div[1]/div/div[2]/div/div[2]/div/div[1]/div/div[2]/label/div/div/div[2]/button",
    'no_results': "/html/body/div[1]/div/div[2]/div/div[2]/div/div[3]/div/span",
}

# Output DataFrame
output_data = []

# Go to the given URL
driver.get('https://sistemas.insabi.gob.mx/contratos/login')
time.sleep(5)  # Wait for the page to load

# Log in
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, elements_xpaths['usuario']))).send_keys("EPH161215NS9")
time.sleep(1)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, elements_xpaths['pass']))).send_keys("Pharmas21")
time.sleep(1)
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['access']))).click()
time.sleep(1)
# Ask the user to login
input("Déja filtradas las facturas y aprieta enter...")

# Loop through each row in the excel file
for _, row in df.iterrows():
    cancelacion_value = row['Orden']
    #numero_orden = 'N/A'  # Default value in case it's not found
    #validacion = 'Not processed'  # Default value in case it's not found

    try:
        # Send the cancelacion value to the cuadro_ordenes input box
        print(f"extrayendo información para la orden {cancelacion_value}")
        cuadro_ordenes = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, elements_xpaths['cuadro_ordenes'])))
        cuadro_ordenes.send_keys(Keys.CONTROL + "a")
        cuadro_ordenes.send_keys(Keys.DELETE)
        cuadro_ordenes.send_keys(cancelacion_value)
        cuadro_ordenes.send_keys(Keys.ENTER)
        time.sleep(5)
        


        try:
            table_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div/div[2]/table')))
            table_html = table_element.get_attribute('outerHTML')
            table_df = pd.read_html(StringIO(table_html))[0]  # Wrap HTML in StringIO and convert to DataFrame
            # Append the table dataframe to the output data list
            output_data.append(table_df)

        except TimeoutException:
            print(f"Table not found for the value {cancelacion_value}")
            output_data.append(pd.DataFrame({'Numero Orden': [cancelacion_value], 'Validacion': ['Table not found']}))
            continue  # Skip to the next iteration

    except Exception as e:
        print(f"An error occurred for the value {cancelacion_value}. Error: {e}")
        output_data.append(pd.DataFrame({'Numero Orden': [cancelacion_value], 'Validacion': [f'Error: {e}']}))
        # Click on the clear button to reset the input for the next iteration
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['clear_button']))).click()
        continue  # Skip to the next iteration

    # Click on the clear button and wait for 2 seconds
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['clear_button']))).click()
    time.sleep(2)

# Concatenate all dataframes in the output_data list
final_output_df = pd.concat(output_data, ignore_index=True)

output_file_name = today_mm_dd + " " + downloaded_set + ".xlsx"

# Save the concatenated dataframe to Excel
final_output_df.to_excel(output_file_name, index=False)

# Close the browser
driver.quit()