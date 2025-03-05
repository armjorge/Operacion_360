from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import time
from io import StringIO
from datetime import datetime
import os

def SAGI_download(driver, username, password, download_directory):
    """
    Automates login into INSABI Contracts and exports table data to Excel.
    
    Parameters:
      driver: Selenium WebDriver instance.
      username: Login username.
      password: Login password.
    
    The function navigates to the login page, logs in with the provided credentials,
    and then prompts the user to navigate to a dataset section. Once confirmed,
    it selects 50 results per page and exports the table data across multiple pages.
    """
    # Define all XPath selectors for the automation
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

    # Get today's date and formatted string for output filename
    today_date = datetime.today().date()
    today_mm_dd = today_date.strftime("%m %d")  # e.g. "03 04"

    # Navigate to the login page
    driver.get('https://sistemas.insabi.gob.mx/contratos/login')
    time.sleep(5)  # Allow the page to load
    
    # Log in using the provided credentials
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, elements_xpaths['usuario']))
    ).send_keys(username)
    time.sleep(1)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, elements_xpaths['pass']))
    ).send_keys(password)
    time.sleep(1)
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, elements_xpaths['access']))
    ).click()
    time.sleep(1)
    
    # Define inner function to extract and export table data
    def export_results(dataset):
        output_data = []
        while True:
            try:
                table_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, elements_xpaths['results_table']))
                )
                table_html = table_element.get_attribute('outerHTML')
                table_df = pd.read_html(StringIO(table_html))[0]
                output_data.append(table_df)
                print("Data extracted and added to output.")
                
                # Try to locate and click the next page button
                try:
                    next_page_button = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, elements_xpaths['next_page']))
                    )
                    if next_page_button.is_enabled():
                        next_page_button.click()
                        time.sleep(5)  # Wait for the page to refresh
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

        # Concatenate all dataframes and export to Excel
        final_output_df = pd.concat(output_data, ignore_index=True)
        output_file_name = os.path.join(download_directory, f"{today_mm_dd} {dataset}.xlsx")

        final_output_df.to_excel(output_file_name, index=False)
        print(f"Data for {dataset} saved to {output_file_name}")

    # Ask user to choose the dataset
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
    print(f"\n*******************************\nNavega a la sección elegida: {downloaded_set}\n"
          "El script va a seleccionar 50 resultados por página y continuar...\n")
    
    # Wait for the user to navigate to the desired section manually
    input(f"Please navigate to the {downloaded_set} section and press Enter to continue.")
    
    # Select the number of results per page (50) by clicking the down arrow and then the option
    down_arrow = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, elements_xpaths['down_arrow']))
    )
    down_arrow.click()
    time.sleep(2)
    results_per_page = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, elements_xpaths['50_results_p_page']))
    )
    results_per_page.click()
    time.sleep(5)
    
    # Export the results for the chosen dataset
    export_results(downloaded_set)
    print("Extraction process completed.")