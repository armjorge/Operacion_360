from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
from io import StringIO  # Import StringIO
from datetime import datetime  # Import datetime

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
    'down_arrow': "/html/body/div[1]/div/div[2]/div/div[2]/div/div[3]/div[2]/label/div/div/div[2]/i",
    '50_results_p_page': "/html/body/div[3]/div[2]/div[7]/div[2]/div",
    'next_page': '//*[@id="q-app"]/div/div[2]/div/div[2]/div/div[3]/div[3]/button[3]',
    'results_table': '//*[@id="q-app"]/div/div[2]/div/div[2]/div/div[2]/table'
}

today_date = datetime.today().date()
today_mm_dd = today_date.strftime("%m %d")  # Format as "mm dd"


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
# Define XPaths for the elements

def export_results(dataset):
    output_data = []

    while True:
        try:
            table_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, elements_xpaths['results_table'])))
            table_html = table_element.get_attribute('outerHTML')
            table_df = pd.read_html(StringIO(table_html))[0]  # Wrap HTML in StringIO and convert to DataFrame
            output_data.append(table_df)
            print("Data extracted and added to output.")

            # Check if the next page button is present and clickable
            try:
                next_page_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, elements_xpaths['next_page'])))
                if next_page_button.is_enabled():
                    next_page_button.click()
                    time.sleep(5)  # Give some time for the page to refresh
                else:
                    print("Next page button is not clickable. End of pages.")
                    break
            except (TimeoutException, NoSuchElementException):
                user_input = input("Next page button not found or not clickable. Is there only one page of results? (yes/no): ").strip().lower()
                if user_input == "yes":
                    break
                else:
                    print("Retrying to locate the next page button...")

        except TimeoutException:
            print("Timeout while trying to extract data.")
            break

    # Concatenate all dataframes in the output_data list
    final_output_df = pd.concat(output_data, ignore_index=True)

    output_file_name = f"{today_mm_dd} {dataset}.xlsx"

    # Save the concatenated dataframe to Excel
    final_output_df.to_excel(output_file_name, index=False)

    print(f"Data for {dataset} saved to {output_file_name}")
# Ask the user for the download set until a valid input is provided
while True:
    downloaded_set = input("\nDescargamos facturas \n1)2023-2024 o \n2)2024?\n")
    if downloaded_set == "1":
        downloaded_set = "2023-2024"
        break
    elif downloaded_set == "2":
        downloaded_set = "2024"
        break
    else:
        print("Invalid input. Please enter 1 or 2.")
        continue
print(f"\n*******************************\nNavega a la sección elegida\n{downloaded_set}\nEl script va a seleccionar 50 resultados por página y seguir*******************************\n")

# Navigate to the selected dataset section
input(f"Please navigate to the {downloaded_set} section and press Enter to continue.")

# Click on down_arrow
down_arrow = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['down_arrow'])))
down_arrow.click()
time.sleep(2)  # Give some time for the dropdown to expand

# Select 50 results per page
results_per_page = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['50_results_p_page'])))
results_per_page.click()
time.sleep(5)  # Give some time for the results to refresh

# Export the results for the selected dataset
export_results(downloaded_set)

print("Extraction process completed.")
driver.quit()