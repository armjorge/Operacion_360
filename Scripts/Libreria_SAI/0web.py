from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import os
from datetime import datetime
import time
import os
import glob

# Set default directory
#os.chdir(os.path.dirname(os.path.abspath(__file__)))
#current_directory = os.getcwd()

#Elimina archivos previos. 
def delete_files_with_prefix(prefixes):
    for prefix in prefixes:
        files = glob.glob(f"{prefix}*.xlsx")
        for file in files:
            try:
                os.remove(file)
                print(f"Deleted {file}")
            except Exception as e:
                print(f"Error occurred while deleting file {file}: {str(e)}")

# Use the function
prefixes_to_delete = ["altas_export_", "ordenes_export_"]
delete_files_with_prefix(prefixes_to_delete)

# Chrome options
chrome_options = webdriver.ChromeOptions()
chrome_binary_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome-win64\\chrome.exe"  # Update this path
chrome_options.binary_location = chrome_binary_path
prefs = {
    "download.default_directory": current_directory,
    "download.prompt_for_download": False,
}
chrome_options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=chrome_options)

# New XPaths
elements_xpaths = {
    'user': "/html/body/main/div[2]/app-root/app-autenticacion/div[3]/form/div[1]/input",
    'password': "/html/body/main/div[2]/app-root/app-autenticacion/div[3]/form/div[2]/input",
    'Login_button': "/html/body/main/div[2]/app-root/app-autenticacion/div[4]/button",
    'Menu': "/html/body/main/div[2]/app-root/app-home/app-header/nav/div/div/ul[2]/li/a",
    'Menu_ordenes': "/html/body/main/div[2]/app-root/app-altas/app-header/nav/div/div/ul[2]/li/a",
    'Altas': "/html/body/main/div[2]/app-root/app-home/app-header/nav/div/div/ul[2]/li/ul/li[6]/a",
    'Altas_inicial': "/html/body/main/div[2]/app-root/app-altas/div[1]/form/div[6]/div[7]/div/input",
    'Altas_final': "/html/body/main/div[2]/app-root/app-altas/div[1]/form/div[6]/div[8]/div/input",
    'Altas_consultar': "/html/body/main/div[2]/app-root/app-altas/div/div[2]/div[2]/button[2]",
    'Altas_exportar': "/html/body/main/div[2]/app-root/app-altas/div[2]/div/button",
    'Ordenes': "/html/body/main/div[2]/app-root/app-altas/app-header/nav/div/div/ul[2]/li/ul/li[3]/a",
    'Ordenes_inicial': "/html/body/main/div[2]/app-root/app-consulta-ordenes/div[3]/form/div[4]/div[3]/div/input",
    'Ordenes_final': "/html/body/main/div[2]/app-root/app-consulta-ordenes/div[3]/form/div[4]/div[4]/div/input",
    'Ordenes_consultar': "/html/body/main/div[2]/app-root/app-consulta-ordenes/div[3]/div[2]/div[2]/button[2]",
    'Ordenes_exportar': "/html/body/main/div[2]/app-root/app-consulta-ordenes/div[4]/div/button",
}

try:
    # Open the URL
    driver.get('https://ppsai-abasto.imss.gob.mx/abasto-web/reporteAltas')
    
    # Login
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Login_button'])))
    driver.find_element(By.XPATH, elements_xpaths['user']).send_keys("EPH -161215-NS9")
    time.sleep(1)
    driver.find_element(By.XPATH, elements_xpaths['password']).send_keys("Epharma2021")
    time.sleep(1)
    driver.find_element(By.XPATH, elements_xpaths['Login_button']).click()

    # Pause
    time.sleep(5)

    # Altas
    action = ActionChains(driver)
    menu_element = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, elements_xpaths['Menu'])))
    action.move_to_element(menu_element).click().perform()
    time.sleep(1)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Altas']))).click()
    today_date = datetime.now().strftime('%d/%m/%Y')
    input_date_element = driver.find_element(By.XPATH, elements_xpaths['Altas_inicial'])
    time.sleep(1)
    input_date_element.send_keys(Keys.ESCAPE)
    time.sleep(1)
    input_date_element.send_keys("01/01/2024")
    time.sleep(1)
    input_date_element = driver.find_element(By.XPATH, elements_xpaths['Altas_final'])
    input_date_element.send_keys(Keys.ESCAPE)
    time.sleep(1)
    input_date_element.send_keys(today_date)
    time.sleep(1)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Altas_consultar']))).click()
    time.sleep(5)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Altas_exportar']))).click()
    try:
        WebDriverWait(driver,300).until_not(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Altas_exportar'])))
    except TimeoutException:
        print("Button remained clickable, possible download error or the button was never made non-clickable during download.")
    # Now, wait for the button to become clickable again (assuming download is complete)
    WebDriverWait(driver, 600).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Altas_exportar'])))
    time.sleep(5)
    
    # Ordenes
    action = ActionChains(driver)
    menu_element = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, elements_xpaths['Menu_ordenes'])))
    action.move_to_element(menu_element).click().perform()
    time.sleep(1)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Ordenes']))).click()
    input_date_element = driver.find_element(By.XPATH, elements_xpaths['Ordenes_inicial'])
    input_date_element.send_keys(Keys.ESCAPE)
    time.sleep(1)
    input_date_element.send_keys("01/01/2024")
    time.sleep(1)
    input_date_element = driver.find_element(By.XPATH, elements_xpaths['Ordenes_final'])
    input_date_element.send_keys(Keys.ESCAPE)
    time.sleep(1)
    input_date_element.send_keys(today_date)
    time.sleep(1)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Ordenes_consultar']))).click()
    time.sleep(1)
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Ordenes_exportar']))).click()
    # Wait for the button to become not clickable (assuming download is in progress)
    try:
        WebDriverWait(driver,300).until_not(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Ordenes_exportar'])))
    except TimeoutException:
        print("Button remained clickable, possible download error or the button was never made non-clickable during download.")
    # Now, wait for the button to become clickable again (assuming download is complete)
    WebDriverWait(driver, 600).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Ordenes_exportar'])))

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    input("Press Enter to close the browser...")
    driver.quit()