import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def convert_date_format(date):
    return date.replace("/", "-")

def clear_input_field(driver, xpath):
    """Ensure the input field is completely cleared."""
    for attempt in range(5):  # Attempt up to 5 times
        try:
            # Re-locate the element to ensure it's current
            input_field = driver.find_element(By.XPATH, xpath)
            # Clear the field using keys
            input_field.send_keys(Keys.CONTROL, 'a')  # Select all text
            time.sleep(0.2)
            input_field.send_keys(Keys.DELETE)  # Delete selected text
            time.sleep(0.2)
            # Get the current value using get_attribute and JavaScript
            current_value = input_field.get_attribute('value')
            js_value = driver.execute_script("return arguments[0].value;", input_field)
            print(f"Attempt {attempt + 1}: Current value (via JS): '{js_value}'")
            # If either method indicates the field is cleared, exit successfully
            if current_value == '__/__/____' or js_value == '__/__/____':
                print(f"Field cleared successfully on attempt {attempt + 1}")
                return True
        except Exception as e:
            print(f"Attempt {attempt + 1}: Failed to clear input field. Error: {e}")
            time.sleep(1)  # Small delay before retrying
    raise TimeoutException("Failed to clear the input field after multiple attempts.")

def input_date(driver, input_field_xpath, date):
    """Clear the input field and input the new date."""
    try:
        print(f"Date passed: {date}")
        actions = ActionChains(driver)
        # Locate the input field and click it
        input_field = driver.find_element(By.XPATH, input_field_xpath)
        actions.click(input_field).perform()
        time.sleep(0.2)
        # Close the calendar popup if it appears
        actions.send_keys(Keys.ESCAPE).perform()
        time.sleep(0.2)
        # Ensure the field is cleared
        clear_input_field(driver, input_field_xpath)
        time.sleep(0.2)
        # Convert the date to a string and input it
        date_str = str(date)
        print(f"Attempting to input date: {date_str}")
        input_field = driver.find_element(By.XPATH, input_field_xpath)  # Re-locate before sending keys
        input_field.send_keys(date_str)
        print(f"Date '{date_str}' entered successfully into the field.")
        time.sleep(0.5)
    except Exception as e:
        print(f"Error in input_date for field '{input_field_xpath}' with date '{date}': {e}")
        raise

def download_files(driver, df, username, password):
    """
    Uses the provided driver to log into the PREI system and process each date range.
    """
    # Define XPaths for all required elements
    elements_xpaths = {
        'close_button': "/html/body/div[2]/div[1]/a",
        'User': "/html/body/main/div[1]/div[2]/div[2]/form/div[2]/div[1]/div/input",
        'Password': "/html/body/main/div[1]/div[2]/div[2]/form/div[2]/div[2]/div[1]/input",
        'Login_button': "/html/body/main/div[1]/div[2]/div[2]/form/div[2]/div[3]/div/button[1]",
        'fecha_inicial': "/html/body/main/div[3]/div/div/form/div[2]/div[3]/span/div[1]/span[2]/input",
        'fecha_final': "/html/body/main/div[3]/div/div/form/div[2]/div[3]/span/div[2]/span[2]/input",
        'buscar': "/html/body/main/div[3]/div/div/form/div[2]/div[4]/div[1]/button/span",
        'excel': "/html/body/main/div[3]/div/div/form/div[5]/a/img",
        'alerta': "/html/body/main/div[3]/div/div/form/div[3]",
        'menu_pagos': "/html/body/main/nav/div/div[2]/form/div/ul/li[2]/a/span[1]",
        'no_results': "/html/body/main/div[3]/div/div/form/div[4]/div[2]/table/tbody/tr/td",
        'facturasvscr': "/html/body/main/nav/div/div[2]/form/div/ul/li[2]/ul/li[6]/a/span"
    }

    # Navigate to the PREI login page
    driver.get('https://pispdigital.imss.gob.mx/piref/')
    time.sleep(2)

    # Login Process using the provided credentials
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['close_button'])))
    driver.find_element(By.XPATH, elements_xpaths['close_button']).click()
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['Login_button'])))
    driver.find_element(By.XPATH, elements_xpaths['User']).send_keys(username)
    driver.find_element(By.XPATH, elements_xpaths['Password']).send_keys(password)
    driver.find_element(By.XPATH, elements_xpaths['Login_button']).click()
    time.sleep(1)

    # Navigate to the 'facturasvscr' section after login
    menu_pagos_element = WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.XPATH, elements_xpaths['menu_pagos']))
    )
    actions = ActionChains(driver)
    actions.move_to_element(menu_pagos_element).perform()
    time.sleep(2)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['facturasvscr']))).click()
    time.sleep(3)

    # Loop through each row (date range) in the DataFrame
    for index, row in df.iterrows():
        try:
            WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located((By.XPATH, elements_xpaths['fecha_inicial']))
            )
            WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located((By.XPATH, elements_xpaths['fecha_final']))
            )

            print(f"Processing row {index + 1}: Date START = {row['DATE START']}, Date END = {row['DATE END']}")
            print("Calling input_date for fecha_inicial...")
            input_date(driver, elements_xpaths['fecha_inicial'], row['DATE START'])
            print("fecha_inicial set successfully.")

            print("Calling input_date for fecha_final...")
            input_date(driver, elements_xpaths['fecha_final'], row['DATE END'])
            print("fecha_final set successfully.")

            print("Clicking 'buscar' button...")
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['buscar']))).click()
            time.sleep(4)

            alerta_element = driver.find_element(By.XPATH, elements_xpaths['alerta'])
            if "Se encontraron más de 100 coincidencias" in alerta_element.text:
                print(f"Dates {row['DATE START']} to {row['DATE END']} got more than 100 invoices, please modify")
                continue  # Skip to the next row

            excel_element = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['excel'])))
            no_results_elements = driver.find_elements(By.XPATH, elements_xpaths['no_results'])

            if len(no_results_elements) > 0 and no_results_elements[0].text == "No se encontraron resultados.":
                print(f"Dates {row['DATE START']} to {row['DATE END']} got no results, moving to the next set.")
                continue  # Skip to the next row
            elif excel_element:
                try:
                    # Wait until any overlay disappears
                    overlay_xpath = "//*[@id='j_idt26_modal']"
                    WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.XPATH, overlay_xpath)))
                except TimeoutException:
                    print("Overlay did not disappear, refreshing the page...")
                    driver.refresh()
                    continue  # Skip to the next row

                excel_element = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, elements_xpaths['excel'])))
                excel_element.click()
                time.sleep(2)
                print(f"Dates {row['DATE START']} to {row['DATE END']} were processed")
            else:
                print(f"Without results for {row['DATE START']} to {row['DATE END']}.")
        except Exception as e:
            print(f"An error occurred: {e}")
        finally:
            print("Moving to the next set.")
            driver.execute_script("window.scrollTo(0, 0);")

def check_missing_files(df, username):
    missing_files = []
    for index, row in df.iterrows():
        date_start = convert_date_format(row['DATE START'])
        date_end = convert_date_format(row['DATE END'])
        file_name = f'[FacturaVsCR][{username}][{date_start}][{date_end}].xls'
        if not os.path.exists(file_name):
            missing_files.append(row)
    return pd.DataFrame(missing_files)

def PREI_downloader(driver, username, password, download_directory, excel_file):
    """
    Main function to execute the PREI downloader process.
    
    Parameters:
      driver: A Selenium WebDriver instance pre-configured with the desired download directory.
      username: Login username.
      password: Login password.
      download_directory: (Not used inside but passed for consistency—driver already has it configured.)
      excel_file: Path to the Excel file containing date ranges.
    """
    df = pd.read_excel(excel_file)
    download_files(driver, df, username, password)
    missing_df = check_missing_files(df, username)
    if missing_df.empty:
        print("All files are present.")
    else:
        print("Missing files for the following date ranges:")
        for index, row in missing_df.iterrows():
            print(f"{convert_date_format(row['DATE START'])} to {convert_date_format(row['DATE END'])}")
        # Attempt to download missing files
        download_files(driver, missing_df, username, password)