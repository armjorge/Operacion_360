from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
import time
from datetime import datetime

def SAI_download(driver, username, password, range_date):
    """
    Perform the Altas and Ordenes downloads using the provided Selenium driver.
    
    The function encapsulates all the XPath-based interactions, login, and download logic.
    """
    elements_xpaths = {
        'user': "/html/body/main/div[2]/app-root/app-autenticacion/div[3]/form/div[1]/input",
        'password': "/html/body/main/div[2]/app-root/app-autenticacion/div[3]/form/div[2]/input",
        'Login_button': "/html/body/main/div[2]/app-root/app-autenticacion/div[3]/form/button",
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
    start_date, end_date = range_date

    try:
        print("[1] Abriendo URL...")
        driver.get('https://ppsai-abasto.imss.gob.mx/abasto-web/reporteAltas')
        time.sleep(2)

        print("[2] Detectando iframes (antes del captcha)...")
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, "iframe")))
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        #print(f"[2.1] Total iframes encontrados: {len(iframes)}")

        # Si deseas mostrar el captcha más grande, cambia el tamaño del iframe aquí o interactúa si necesario

        input("🔐 Por favor valida el captcha en el navegador y presiona ENTER para continuar...")

        #print("[3] Regresando al contenido principal (por si estaba en un iframe)...")
        driver.switch_to.default_content()

        print("[4] Buscando campo de usuario en el DOM principal...")
        user_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, elements_xpaths['user']))
        )
        user_field.clear()
        user_field.send_keys(username)

        print(" → Usuario enviado")

        print("[4.2] Buscando campo de contraseña...")
        password_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, elements_xpaths['password']))
        )
        password_field.clear()
        password_field.send_keys(password)

        print(" → Contraseña enviada")

        print("[5] Clic en login...")
        login_button = driver.find_element(By.XPATH, elements_xpaths['Login_button'])
        login_button.click()

        print("✅ Login enviado con éxito.")



        # --- Altas section ---
        action = ActionChains(driver)
        menu_element = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, elements_xpaths['Menu']))
        )
        action.move_to_element(menu_element).click().perform()
        time.sleep(1)
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, elements_xpaths['Altas']))
        ).click()
        
        
        
        # Set the start date for Altas
        input_date_element = driver.find_element(By.XPATH, elements_xpaths['Altas_inicial'])
        time.sleep(1)
        input_date_element.send_keys(Keys.ESCAPE)
        time.sleep(1)
        input_date_element.send_keys(start_date)
        time.sleep(1)
        
        # Set the end date for Altas
        input_date_element = driver.find_element(By.XPATH, elements_xpaths['Altas_final'])
        input_date_element.send_keys(Keys.ESCAPE)
        time.sleep(1)
        input_date_element.send_keys(end_date)
        time.sleep(1)
        
        # Execute Altas query and export
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, elements_xpaths['Altas_consultar']))
        ).click()
        time.sleep(5)
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, elements_xpaths['Altas_exportar']))
        ).click()
        try:
            WebDriverWait(driver, 300).until_not(
                EC.element_to_be_clickable((By.XPATH, elements_xpaths['Altas_exportar']))
            )
        except TimeoutException:
            print("Button remained clickable in Altas section. Possible download error.")
        # Wait until the export button becomes clickable again (download completion)
        WebDriverWait(driver, 600).until(
            EC.element_to_be_clickable((By.XPATH, elements_xpaths['Altas_exportar']))
        )
        time.sleep(5)

        # --- Ordenes section ---
        action = ActionChains(driver)
        menu_element = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, elements_xpaths['Menu_ordenes']))
        )
        action.move_to_element(menu_element).click().perform()
        time.sleep(1)
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, elements_xpaths['Ordenes']))
        ).click()
        
        # Set the start date for Ordenes
        input_date_element = driver.find_element(By.XPATH, elements_xpaths['Ordenes_inicial'])
        input_date_element.send_keys(Keys.ESCAPE)
        time.sleep(1)
        input_date_element.send_keys(start_date)
        time.sleep(1)
        
        # Set the end date for Ordenes
        input_date_element = driver.find_element(By.XPATH, elements_xpaths['Ordenes_final'])
        input_date_element.send_keys(Keys.ESCAPE)
        time.sleep(1)
        input_date_element.send_keys(end_date)
        time.sleep(1)
        
        # Execute Ordenes query and export
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, elements_xpaths['Ordenes_consultar']))
        ).click()
        time.sleep(1)
        WebDriverWait(driver, 100).until(
            EC.element_to_be_clickable((By.XPATH, elements_xpaths['Ordenes_exportar']))
        ).click()
        try:
            WebDriverWait(driver, 300).until_not(
                EC.element_to_be_clickable((By.XPATH, elements_xpaths['Ordenes_exportar']))
            )
            time.sleep(5)
        except TimeoutException:
            print("Button remained clickable in Ordenes section. Possible download error.")
        # Wait until the export button becomes clickable again (download completion)
        WebDriverWait(driver, 600).until(
            EC.element_to_be_clickable((By.XPATH, elements_xpaths['Ordenes_exportar']))
        )
        
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # ✅ Verificación manual antes de cerrar Chrome
        input("🧾 Revisa por favor la sección 'Descargas' del navegador Chrome.\nSi ves ambos archivos descargados correctamente, presiona ENTER para cerrar.")
        driver.quit()