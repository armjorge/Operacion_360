from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def SAI_camunda(driver, username, password):
    """
    Automates the login and preliminary steps on the Camunda task list page.
    
    Parameters:
      driver: Selenium WebDriver instance.
      username: The username for login.
      password: The password for login.
    """
    # Define all the XPath selectors used in this automation
    elements_xpaths = {
        'user': '/html/body/div[2]/div/div/form/input[1]',
        'password': '/html/body/div[2]/div/div/form/input[2]',
        'sign_in': '/html/body/div[2]/div/div/form/button'
    }
    
    # Navigate to the Camunda URL
    camunda_url = ("https://gestorinsumos.salud.gob.mx/camunda/app/tasklist/default/#/?searchQuery=%5B%5D&variableCaseHandling=all&filter=ff2e4ae5-e09a-11e9-ab7a-005056876693&sorting=%5B%7B%22sortBy%22:%22created%22,%22sortOrder%22:%22desc%22%7D%5D&task=a851299e-8503-11ec-b746-005056ace5cb&detailsTab=task-detail-form")
    driver.get(camunda_url)
    
    # Wait for the login fields to be visible and perform login
    WebDriverWait(driver, 120).until(
        EC.visibility_of_element_located((By.XPATH, elements_xpaths['user']))
    )
    driver.find_element(By.XPATH, elements_xpaths['user']).send_keys(username)
    time.sleep(1)
    driver.find_element(By.XPATH, elements_xpaths['password']).send_keys(password)
    driver.find_element(By.XPATH, elements_xpaths['sign_in']).click()
    time.sleep(3)  # allow time for the page to load
    
    # Wait for user to filter tasks manually if needed
    input("Filtra y presiona enter cuando estés listo...")
    