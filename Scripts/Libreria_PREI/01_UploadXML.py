from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import csv
import time
import pyperclip
import pyautogui
import pandas as pd

# Set the directory
directory = r'.\Acuses'  # Using a relative path; consider using an absolute path in production
if not os.path.exists(directory):
    os.makedirs(directory)  # Create the directory if it doesn't exist

# Set up Chrome options
chrome_options = webdriver.ChromeOptions()

# Path to the Chrome executable
chrome_binary_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome-win64\\chrome.exe"
chrome_options.binary_location = chrome_binary_path

# Set the default download directory to the specified directory
prefs = {
    "download.default_directory": os.path.abspath(directory),
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

# Read the Excel file
df = pd.read_csv('.\\uploadXMLlist.csv')
# Read the files in .\\Acuses and create df_uploaded
files_in_acuses = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
df_uploaded = pd.DataFrame(files_in_acuses, columns=['Filename'])
# Replace '_.pdf' by '.xml' in filenames
df_uploaded['Filename'] = df_uploaded['Filename'].str.replace('_.pdf', '.xml', regex=False)
# Update df to skip those who match any of files found in df_uploaded
df = df[~df['Factura'].isin(df_uploaded['Filename'])]
df.to_csv('.\\uploadXMLlist_sinacuse.csv', index=False)


# Define XPaths of all the required elements
elements_xpaths = {
    'close_button': "/html/body/div[2]/div[1]/a",
    'User': "/html/body/main/div[1]/div[2]/div[2]/form/div[2]/div[1]/div/input",
    'Password': "/html/body/main/div[1]/div[2]/div[2]/form/div[2]/div[2]/div[1]/input",
    'Login_button': "/html/body/main/div[1]/div[2]/div[2]/form/div[2]/div[3]/div/button[1]",
    'seleccionar': "/html/body/main/div[3]/div/div/form/div/div[1]/div/div[1]/span",
    'descargar': "/html/body/main/div[3]/div/div/form/div/div[2]/button",
    'table': ["/html/body/main/div[3]/div/div/span/div[1]/div[2]/table/tbody/tr/td[1]",
              "/html/body/main/div[3]/div/div/span/div[1]/div[2]/table/tbody/tr/td[2]",
              "/html/body/main/div[3]/div/div/span/div[1]/div[2]/table/tbody/tr/td[3]",
              "/html/body/main/div[3]/div/div/span/div[1]/div[2]/table/tbody/tr/td[4]",
              "/html/body/main/div[3]/div/div/span/div[1]/div[2]/table/tbody/tr/td[5]",
              "/html/body/main/div[3]/div/div/span/div[1]/div[2]/table/tbody/tr/td[6]"],
    'buzon': "/html/body/main/nav/div/div[2]/form/div/ul/li[1]/a/span"
}

# The headers of your excel file
header = 'Factura'

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
    input("Sube un XML de prueba para definir la carpeta fuente")

    # Initialize the file counter
    file_counter = 0

    # For each row in the dataframe, perform the upload action
    for index, row in df.iterrows():
        time.sleep(2)
        # Wait until the 'seleccionar' element is clickable
        #WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['seleccionar']))).click()
        
        max_retries = 3  # Maximum number of retries to avoid an infinite loop
        retry = 0

        while retry < max_retries:
            try:
                # Wait until the 'seleccionar' element is clickable and click it
                WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['seleccionar']))).click()
                time.sleep(2)
                break  # If click is successful, break the loop
            
            except Exception as e:
                print(f"An error occurred while trying to click 'seleccionar': {str(e)}")
                print("Refreshing the page and retrying...")
                
                # Refresh the page
                driver.refresh()
                
                # Increment retry count
                retry += 1
                
                # Optionally, wait for a few seconds after refreshing the page to ensure the page has fully loaded
                time.sleep(5)
                
        else:
            print("Failed to click 'seleccionar' after maximum retries. Moving to next iteration.")
            # Optionally, skip the current iteration and move to the next file
            continue  
        # Copy to clipboard
        pyperclip.copy(str(row[header]))
        # Use pyautogui to paste the value and press Enter
        time.sleep(1)  # Give a short delay before pasting
        pyautogui.hotkey('ctrl', 'v')  # This will paste the value
        time.sleep(1)  # Give a short delay before pressing Enter
        pyautogui.press('enter')  # This will press the Enter key
        print(f"Value {row[header]} is copied to clipboard. Please paste it into the web form.")
        time.sleep(2)
        # Get table data for each cell and print it
        table_data = [WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, cell_xpath))).text for cell_xpath in elements_xpaths['table']]

        # Write the table data and filename to a CSV file
        with open('output.csv', 'a', newline='') as f:
            writer = csv.writer(f)
            writer.writerow([str(row[header])] + table_data)

        # Click the 'descargar' button
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['descargar']))).click()

        # Wait for the download to complete
        time.sleep(2)

        # Print success message
        print(f"The value {str(row[header])} was uploaded, the following will proceed in 8 seconds.")

        # Click 'buzon'
        max_retries = 3  # Maximum number of retries to avoid infinite loop
        retry = 0

        while retry < max_retries:
            try:
                # Click 'buzon'
                WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['buzon']))).click()
                
                # Wait for 8 seconds before moving to the next file
                time.sleep(8)
                break  # If click is successful, break the loop
            
            except Exception as e:
                print(f"An error occurred while trying to click 'buzon': {str(e)}")
                print("Refreshing the page and retrying...")
                
                # Refresh the page
                driver.refresh()
                
                # Increment retry count
                retry += 1
                
                # Optionally, wait for a few seconds after refreshing the page
                time.sleep(5)
                
        else:
            print("Failed to click 'buzon' after maximum retries. Moving to next iteration.")

        # Increment the file counter
        file_counter += 1

        # Check if 5 files have been processed
        if file_counter >= 5:
            print("Processed 5 files, refreshing the page...")
            driver.refresh()
            file_counter = 0  # Reset the counter
            time.sleep(5)  # Wait for the page to reload
            
except Exception as e:
    print("An error occurred: ", e)

finally:
    # Keep the browser window open until the user manually closes it
    input("Press Enter to quit...")

    # Close the browser
    driver.quit()
