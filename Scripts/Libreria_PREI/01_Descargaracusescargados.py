from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

import pandas as pd
import os
import csv
import time
import pyperclip

# Set up Chrome options
chrome_options = Options()
# Specify the default download directory
download_folder = r'.\Acuses'
prefs = {"download.default_directory": os.path.abspath(download_folder)}
chrome_options.add_experimental_option("prefs", prefs)

# Create a new instance of the Chrome driver with the specified options
driver = webdriver.Chrome(options=chrome_options)

# Read the Excel file
df = pd.read_excel('CRsinCR.xlsx')

# Define XPaths of all the required elements
elements_xpaths = {
    'close_button': "/html/body/div[2]/div[1]/a",
    'User': "/html/body/main/div[1]/div[2]/div[2]/form/div[2]/div[1]/div/input",
    'Password': "/html/body/main/div[1]/div[2]/div[2]/form/div[2]/div[2]/div[1]/input",
    'Login_button': "/html/body/main/div[1]/div[2]/div[2]/form/div[2]/div[3]/div/button[1]",
    'ingreso_folio': "/html/body/main/div[3]/div/div/form/div[2]/div[3]/span/div[1]/input",
    'buscar_folio': "/html/body/main/div[3]/div/div/form/div[2]/div[4]/div[1]/button",
    'descargar_acuse': "/html/body/main/div[3]/div/div/form/div[3]/div/button",
}

# The headers of your excel file
header = 'Folio'

try:
    # Go to the given URL
    driver.get('https://pispdigital.imss.gob.mx/piref/')
    time.sleep(5)

    # Login Process
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['close_button'])))
    driver.find_element(By.XPATH, elements_xpaths['close_button']).click()
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Login_button'])))
    driver.find_element(By.XPATH, elements_xpaths['User']).send_keys("0000150462")
    driver.find_element(By.XPATH, elements_xpaths['Password']).send_keys("EPH161215NS9")
    driver.find_element(By.XPATH, elements_xpaths['Login_button']).click()
	
	# Ask the user to proceed
    input("Llega a la sección Recupera acuse por folio fiscal y presiona enter...")

    # For each row in the dataframe, perform the upload action
    for index, row in df.iterrows():

        # Wait until the 'ingreso_folio' element is visible
        input_folio = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, elements_xpaths['ingreso_folio'])))
        time.sleep(1)
        # Clear the input field
        input_folio.clear()
        time.sleep(2)
        # Paste the folio from the dataframe
        input_folio.send_keys(str(row[header]))
        time.sleep(1)
        print(f"Value {row[header]} is pasted into the input field.")

        # Click the 'buscar_folio' button
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['buscar_folio']))).click()
        time.sleep(2)
        # Click the 'descargar_acuse' button after it becomes clickable
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['descargar_acuse']))).click()
        time.sleep(2)
        # Wait for the download to complete
        time.sleep(5)

        # Print success message
        print(f"The folio {str(row[header])} was processed, moving to the next one.")

except Exception as e:
    print("An error occurred: ", e)

finally:
    # Keep the browser window open until the user manually closes it
    input("Press Enter to quit...")

    # Close the browser
    driver.quit()
