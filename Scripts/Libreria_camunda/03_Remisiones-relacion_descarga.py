from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

import pandas as pd
import math
import os
import time
import pygame

pygame.mixer.init()

# Load your MP3 file
pygame.mixer.music.load("alarm.mp3")


chrome_options = webdriver.ChromeOptions()
#current_directory = os.getcwd()
current_directory = r"C:\Users\armjorge\Dropbox\3. Armando Cuaxospa\Reportes GPT\Dataframes\Camunda\INSABI_Remisiones 2024\Descargadas"

chrome_binary_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome-win64\\chrome.exe"  # Update this path
chrome_options.binary_location = chrome_binary_path
prefs = {
    "download.default_directory": current_directory,
    "download.prompt_for_download": False}
chrome_options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=chrome_options)

# Initialize DataFrame
df_index_orden = pd.DataFrame(columns=['Index', 'Orden'])

elements_xpaths = {
    'user': '/html/body/div[2]/div/div/form/input[1]',
    'password': '/html/body/div[2]/div/div/form/input[2]',
    'sign_in': '/html/body/div[2]/div/div/form/button'
}

df_grabResults = pd.DataFrame(columns=['orden', 'remision'])

driver.get("https://gestorinsumos.salud.gob.mx/camunda/app/tasklist/default/#/?searchQuery=%5B%5D&variableCaseHandling=all&filter=ff2e4ae5-e09a-11e9-ab7a-005056876693&sorting=%5B%7B%22sortBy%22:%22created%22,%22sortOrder%22:%22desc%22%7D%5D&task=a851299e-8503-11ec-b746-005056ace5cb&detailsTab=task-detail-form")
# Wait until the input fields are loaded
WebDriverWait(driver, 120).until(EC.visibility_of_element_located((By.XPATH, elements_xpaths['user'])))
driver.find_element(By.XPATH, elements_xpaths['user']).send_keys("ArmandoJimenez")
time.sleep(1)
driver.find_element(By.XPATH, elements_xpaths['password']).send_keys("N29f6Mwif")
driver.find_element(By.XPATH, elements_xpaths['sign_in']).click()
time.sleep(3)  # give some time for the page to load the new content
input("Filtra y presiona enter cuando estés listo...")

# Calculate total results and pages
total_results_xpath = '/html/body/div[2]/div/div/section[1]/div/div/div[3]/div[3]/h4/a/span'
total_results = int(WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, total_results_xpath))).text)
results_per_page = 15
total_pages = math.ceil(total_results / results_per_page)

# Initialize data for DataFrame
data = []

# Populate data considering the total results and results per page
for page in range(1, total_pages + 1):
    results_this_page = min(results_per_page, total_results - (page - 1) * results_per_page)
    for i in range(1, results_this_page + 1):
        data.append({"Page": page, "i-element": i, "order": None, "remision": None})

# Create the DataFrame
df_grabResults = pd.DataFrame(data)
df_grabResults.to_csv('grabResults.csv', index=False)

remision_xpath = '/html/body/div[2]/div/div/section[3]/div/div/div[2]/section/div[2]/div/div/view/div/div/div[2]/div[2]/div/div/div/div/form/div[1]/div[5]/div/div/div/div/div[3]/div[2]/div/div/div/div[2]/div'
orden_xpath = '/html/body/div[2]/div/div/section[3]/div/div/div[2]/section/div[2]/div/div/view/div/div/div[2]/div[2]/div/div/div/div/form/div[1]/div[5]/div/div/div/div/div[3]/div[2]/div/div/div/div[3]/div'

def download_if_missing(order, remission):  
    print(f"Downloading missing file for Order: {order}, Remission: {remission}")
    try:
        download_button_xpath = '/html/body/div[2]/div/div/section[3]/div/div/div[2]/section/div[2]/div/div/view/div/div/div[2]/div[2]/div/div/div/div/form/div[1]/div[4]/div/button'
        download_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, download_button_xpath)))
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", download_button)
        time.sleep(1)  # Allowing time for any potential animations to complete
        download_button.click()
        # Wait for the file to download
        time.sleep(10)
    except (NoSuchElementException, TimeoutException) as e:
        print("Error while trying to download the file. Please download manually.")
        pygame.mixer.music.play()
        input("After manual download, press Enter to continue...")

def click_and_extract(i):
    """Clicks on the i-th item and extracts 'order' and 'remission', with retry logic."""
    attempt_count = 0
    while attempt_count <= 3:
        try:
            i_element_xpath = f'/html/body/div[2]/div/div/section[2]/div/div/div[3]/div/ol/li[{i}]/div'
            clickable = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, i_element_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", clickable)
            time.sleep(0.5)  # Adjust based on your page's behavior
            """ Esto sí funciona, lo detecta bien, el problema es que sigue intentando extraer datos aunque el header indique que la información es vacía. 
            # Vamos a agregar la opción de no descargas los "ver órdenes recibidas por almacén"#
            header_xpath = '/html/body/div[2]/div/div/section[3]/div/div/div[2]/section/header/div/div[1]/h2'
            header = driver.find_element(By.XPATH, header_xpath)
            
            if header.text == "Ver órdenes recibidas por almacén":
                print(f"Skipping item {i} as it contains a specific header indicating it's empty.")
                return None, None  # Skip this item and return None values
            #Hasta aquí llega la parte que le agregué. 
            """
            clickable = WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, i_element_xpath)))
            clickable.click()

            # Extract the information
            order = WebDriverWait(driver, 50).until(EC.visibility_of_element_located((By.XPATH, orden_xpath))).text
            remission = WebDriverWait(driver, 50).until(EC.visibility_of_element_located((By.XPATH, remision_xpath))).text
            #print("Extracted remission:", remission)
            return order, remission
        except Exception as e:
            pygame.mixer.music.play()
            print(f"Error on item {i}: {e}")
            attempt_count += 1
            if attempt_count > 3:
                input(f"Manual intervention required for page {page}, item {i}. Press Enter to continue...")
                attempt_count = 0  # Reset attempt count after intervention
            else:
                print(f"Retrying extraction for page {page}, item {i}...")
    time.sleep(1)
    return None, None

def first_element_click_and_extract():
    """Extracts 'order' and 'remission' of the first item."""
    return click_and_extract(1)
    
def update_dataframe(page, i, order, remission):
    """Updates the DataFrame with the extracted details."""
    df_index = df_grabResults[(df_grabResults['Page'] == page) & (df_grabResults['i-element'] == i)].index
    if not df_index.empty:
        df_grabResults.at[df_index[0], 'order'] = order
        df_grabResults.at[df_index[0], 'remision'] = remission

def navigate_to_next_page(page, is_first_page=True):
    # Attempt to navigate directly to the page if it's visible
    if is_first_page or page <= 5:  # Assuming the first 5 pages are always directly accessible
        try:
            # Find all the pagination links currently visible
            pagination_links = driver.find_elements_by_xpath("/html/body/div[2]/div/div/section[2]/div/div/ul/li/a")
            
            # Iterate over each link to find the one that matches the desired page number
            for link in pagination_links:
                if link.text == str(page):
                    link.click()
                    print(f"Directly navigated to page {page}.")
                    time.sleep(1)  # Adjust based on the actual page response time
                    return
        except Exception as e:
            print(f"Error trying to directly navigate to page {page}: {e}")
            
            # If there's an error in direct navigation, fallback to manual intervention or next button click
    
    # If not the first page and page number is not directly accessible, click the "Next" button
    if not is_first_page:
        try:
            next_page_button_xpath = "/html/body/div[2]/div/div/section[2]/div/div/ul/li[8]/a"
            next_page_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, next_page_button_xpath))
            )
            next_page_button.click()
            time.sleep(1)  # Adjust based on the actual page response time
        except Exception as e:
            pygame.mixer.music.play()
            print(f"Error navigating to the next page {page}: {e}")
            input("Manual intervention required. Please navigate to the correct page and press Enter to continue...")
            
def check_for_error_and_go_back():
    try:
        # Attempt to find the element by XPath
        error_element = driver.find_element(By.XPATH, '/html/body/pre')

        if error_element:
            print("Error page detected. Navigating back...")
            driver.back()  # Navigate back if the element is found
            time.sleep(2)  # Wait for 2 seconds after navigating back
            return True
    except NoSuchElementException:
        # If the element is not found, just return False
        return False


"""
# Versión que ya funcionaba
def should_skip_item(i):
    #Checks if the item contains a header that indicates it should be skipped
    # Updated XPath to target the <a> tag inside the <h4> under the specified div for the i-th item
    header_xpath = f'/html/body/div[2]/div/div/section[2]/div/div/div[3]/div/ol/li[{i}]/div/div[1]/h4/a'
    try:
        # Using WebDriver to find the element based on XPath
        header = driver.find_element(By.XPATH, header_xpath)
        if header.text.strip() == "Ver órdenes recibidas por almacén":
            print(f"Skipping item {i} due to specific header.")
            return True
    except NoSuchElementException:
        print(f"No specific header found for item {i}, proceeding with extraction.")
        return False
    return False
"""

# Versión 18/06/2024
def should_skip_item(i):
    """Checks if the item contains a header that indicates it should be skipped."""
    header_xpath = f'/html/body/div[2]/div/div/section[2]/div/div/div[3]/div/ol/li[{i}]/div/div[1]/h4/a'
    
    try:
        # Wait for up to 30 seconds for the element to be present
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, header_xpath))
        )
        
        # Find the element after it's confirmed to be present
        header = driver.find_element(By.XPATH, header_xpath)
        if header.text.strip() == "Ver órdenes recibidas por almacén":
            print(f"Skipping item {i} due to specific header.")
            return True
    except NoSuchElementException:
        print(f"No specific header found for item {i}, proceeding with extraction.")
        return False
    except TimeoutException:
        print(f"Timeout reached while waiting for header in item {i}.")
        return False
    except Exception as e:
        print(f"An error occurred while checking item {i}: {e}")
        return False

    return False

# Main loop to populate the DataFrame
unique_orders = set()  # Initialize a set to track unique 'order' values
df_path = "./INSABI_Remisiones 2024/extraccionRemisionesINSABI_brut.csv"
df_PDFRemisiones = pd.read_csv(df_path)
#print(df_PDFRemisiones.head())
df_PDFRemisiones['Oremision'] = df_PDFRemisiones['Oremision'].astype(str).str.replace('.0', '', regex=False)
# Strip any leading or trailing whitespaces
df_PDFRemisiones['Oremision'] = df_PDFRemisiones['Oremision'].str.strip()
df_PDFRemisiones['Oremision'].to_csv('Oremision_list.csv', index=False)
print(df_PDFRemisiones.dtypes)

#print(df_PDFRemisiones.head())
#print(df_PDFRemisiones['Oremision'].tolist()[:10])  # Print the first 10 remission numbers to verify

for page in range(1, total_pages + 1):
    # Attempt to navigate to the next page
    navigate_to_next_page(page, is_first_page=(page==1))
    if check_for_error_and_go_back():
        continue  # Skip the rest of the loop iteration if error handled    
    # Check the first item to confirm it's a new page
    attempt_count = 0
    while True:
        order, remission = click_and_extract(1)  # Extract data for the first item to check if it's a new page
        if order not in unique_orders and order is not None:
            unique_orders.add(order)  # Add the unique 'order' to the set
            break  # Break from the loop if the page is confirmed to be new
        else:
            print(f"Page {page} not loaded correctly, attempting to reload...")
            attempt_count += 1
            if attempt_count > 3:  # After multiple failed attempts
                pygame.mixer.music.play()
                input("Manual intervention required. Please navigate to the correct page and press Enter to continue...")
                attempt_count = 0  # Reset the attempt count after manual intervention

            # Re-attempt to navigate to the page
            navigate_to_next_page(page)
    
    # Calculate the expected number of items on the current page
    items_expected_on_current_page = min(results_per_page, total_results - ((page - 1) * results_per_page))

    # Proceed with data extraction for each item on the page, adjusted for the last page
    for i in range(1, items_expected_on_current_page + 1):
        if should_skip_item(i):
            continue  # Skip this item and go to the next one    
        order, remission = click_and_extract(i)
        if order and remission:
            update_dataframe(page, i, order, remission)
            time.sleep(2)
            if should_skip_item(i):
                continue  # Skip this item and go to the next one            
            if not df_PDFRemisiones['Oremision'].isin([remission]).any():
                download_if_missing(order, remission) #Esto lo identé y le agregué el if not de arriba, si no funciona quítaselo
            else:
                print(f"Remission: {remission} already listed. Skipping download.")
        else:
            print(f"Continuing after issue with page {page}, item {i}.")


    # Save the DataFrame after each page to incrementally save progress.
    df_grabResults.to_csv('grabResults.csv', index=False)