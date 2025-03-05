from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import glob


# Function to delete specific files
def delete_specific_files(directory, pattern):
    # Get the list of files matching the pattern
    files_to_delete = glob.glob(os.path.join(directory, pattern))
    
    # Delete the files and print their names
    deleted_files = []
    for file_path in files_to_delete:
        if file_path != os.path.join(directory, 'ordenesSuministro.xlsx') and file_path != os.path.join(directory, 'ordenesSuministro outCamunda.xlsx'):
            os.remove(file_path)
            deleted_files.append(file_path)
    
    # Print the deleted files
    if deleted_files:
        print("Deleted files:")
        for file in deleted_files:
            print(file)
    else:
        print("No files to delete.")

chrome_options = webdriver.ChromeOptions()
# Directories for Excel and PDF files

excel_directory = r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\2024 PISP\UpdateGSHEET2024"
# Archivos previamente descargados por ejemplo "ordenesSuministro (02).xlsx"
pattern = "ordenesSuministro*.xlsx"

pdfs_directory = r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Dataframes\Camunda\INSABI_Órdenes 2024"
choice = input("Do we download 'OSuministro' Excel or PDFs? (type 'excel' or 'pdfs'): ").strip().lower()
if choice == "excel":
    current_directory = excel_directory
    delete_specific_files(current_directory, pattern)    
elif choice == "pdfs":
    current_directory = pdfs_directory
else:
    current_directory = None
    print("Invalid choice. Please type 'excel' or 'pdfs'.")
if current_directory:
    print("Current directory set to:", current_directory)

chrome_binary_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome-win64\\chrome.exe"  # Update this path
chrome_options.binary_location = chrome_binary_path
prefs = {
    "download.default_directory": current_directory,
    "download.prompt_for_download": False}
chrome_options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=chrome_options)

elements_xpaths = {
    'user': '/html/body/div[2]/div/div/form/input[1]',
    'password': '/html/body/div[2]/div/div/form/input[2]',
    'sign_in': '/html/body/div[2]/div/div/form/button'
}
driver.get("https://gestorinsumos.salud.gob.mx/camunda/app/tasklist/default/#/?searchQuery=%5B%5D&variableCaseHandling=all&filter=ff2e4ae5-e09a-11e9-ab7a-005056876693&sorting=%5B%7B%22sortBy%22:%22created%22,%22sortOrder%22:%22desc%22%7D%5D&task=a851299e-8503-11ec-b746-005056ace5cb&detailsTab=task-detail-form")
# Wait until the input fields are loaded
WebDriverWait(driver, 120).until(EC.visibility_of_element_located((By.XPATH, elements_xpaths['user'])))
driver.find_element(By.XPATH, elements_xpaths['user']).send_keys("ArmandoJimenez")
time.sleep(1)
driver.find_element(By.XPATH, elements_xpaths['password']).send_keys("N29f6Mwif")
driver.find_element(By.XPATH, elements_xpaths['sign_in']).click()
time.sleep(3)  # give some time for the page to load the new content
input("Filtra y presiona enter cuando estés listo...")
