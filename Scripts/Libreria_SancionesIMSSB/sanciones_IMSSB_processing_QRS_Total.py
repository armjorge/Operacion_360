import os
import time
import base64
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

from selenium.webdriver.common.by import By
import csv
import sys 
import pandas as pd
import fitz
import numpy as np
import cv2
import platform

def read_QR_infile(file_path):
    #
    #Scan the given PDF for QR codes on every page,
    #returning a list of (modified_filename, qr_data) tuples.
    #
    results = []
    base = os.path.splitext(os.path.basename(file_path))[0]

    try:
        doc = fitz.open(file_path)
    except Exception as e:
        print(f"Error opening {base}.pdf: {e}")
        return results

    for page_number in range(len(doc)):
        try:
            images = doc.get_page_images(page_number)
        except Exception as e:
            print(f"  ⚠️ Skipping page {page_number} of {base}: {e}")
            continue

        for img in images:
            xref = img[0]
            try:
                pix = fitz.Pixmap(doc, xref)
            except Exception as e:
                print(f"  ⚠️ Couldn’t build pixmap for {base}, page {page_number}: {e}")
                continue

            # build a NumPy array from pix.samples
            try:
                raw = np.frombuffer(pix.samples, dtype=np.uint8)
                raw = raw.reshape((pix.height, pix.stride))
                valid_length = pix.width * pix.n
                raw = raw[:, :valid_length]
                raw = raw.reshape((pix.height, pix.width, pix.n))
                if pix.n == 4:
                    cv_img = cv2.cvtColor(raw, cv2.COLOR_RGBA2BGR)
                elif pix.n == 3:
                    cv_img = cv2.cvtColor(raw, cv2.COLOR_RGB2BGR)
                else:
                    cv_img = raw
            except Exception as e:
                print(f"  ⚠️ Error converting image for {base}: {e}")
                pix = None
                continue

            detector = cv2.QRCodeDetector()
            qr_data, _, _ = detector.detectAndDecode(cv_img)
            if qr_data:
                modified_filename = f"{base}_SAT"
                results.append((modified_filename, qr_data))
            else:
                print(f"  • No QR on {base} page {page_number}")
            pix = None

    if results:
        print(f"✓ Found {len(results)} QR(s) in {base}")
    else:
        print(f"✗ No QR in {base}")

    return results


def STEP_A_generate(directory, csv_file_name):  # Generate the CSV(directory)
    #print(f"built csv file path {csv_file_name}")
    with open(csv_file_name, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["Guardar como", "Link"])  # Write CSV header

        # Process each PDF that is not _REM, _TXT, or _SAT
        for filename in os.listdir(directory):
            if filename.endswith(".pdf") and not (
                filename.endswith("_REM.pdf") or 
                filename.endswith("_TXT.pdf") or 
                filename.endswith("_SAT.pdf")
            ):
                #file_path = os.path.join(directory, filename)
                file_path = os.path.join(directory, filename)
                qr_pairs = read_QR_infile(file_path)
                for modified_name, link in qr_pairs:
                    writer.writerow([modified_name, link])

    print(f"Completed writing to {csv_file_name}")
    
    # Load the generated CSV to check if it has content
    df_uploadMergelist = pd.read_csv(csv_file_name)

    # Check if the DataFrame is empty and print the appropriate message
    if df_uploadMergelist.empty:
        print("\n*******************************\nNo hay QRs por descargar \n*******************************\n")
    else:
        print(f"\n*******************************\nHay {len(df_uploadMergelist)} facturas por descargar. \n*******************************\n")


# Function to read the CSV file in the specified directory
def read_csv(directory):
    print("\nReading CSV\n")
    file_path = os.path.join(directory, f"{os.path.basename(directory)}.csv")
    
    if not os.path.exists(file_path):
        print(f"CSV file not found in {directory}")
        return []

    # List all existing PDF files in the directory, removing the .pdf extension
    existing_files = {os.path.splitext(f)[0] for f in os.listdir(directory) if f.endswith(".pdf")}
    
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        header = next(reader)  # Read the header row
        data = [header]  # Start data list with the header

        for row in reader:
            file_name = row[0]  # Assume 'Guardar como' is the first column
            if file_name in existing_files:
                print(f"{file_name} skipped because corresponding PDF already exists.")
                continue  # Skip this row if the PDF file already exists
            data.append(row)  # Add row if the file doesn't exist
    if len(data) == 1:
        print("Your work is done, good")
        sys.exit()  # Exit the script
    return data

# Download PDFs function to simulate PDF download
def download_pdfs(data, directory, chrome_driver):
    driver = chrome_driver  # Initialize Chrome driver
    if not driver:
        print("Driver failed to initialize. Exiting.")
        return
    processed_count = 0
    total_rows = len(data) - 1  # Total rows excluding the header


    for idx, row in enumerate(data[1:], start=1):  # Adjusts index to start from 1 for clarity in logs
        file_name, link = row
        print(f"Starting {idx}/{len(data) - 1}: {file_name} - {link}")
        
        # Ensure link has the correct format
        if not link.startswith("http://") and not link.startswith("https://"):
            link = "https://" + link  # Add protocol if missing
        
        driver.get(link)  # Open the URL
        time.sleep(5)  # Initial wait time for the page to load

        # Wait for manual captcha verification
        print("Please complete the captcha manually.")
        while True:
            try:
                element = driver.find_element(By.XPATH, '//*[@id="ctl00_MainContent_TxtCaptchaNumbers"]')
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                element.click()
                
                driver.find_element(By.XPATH, '//input[@value="Verificar CFDI"]')
                break
            except:
                time.sleep(1)

        # Locate the element by XPath and click on it
        #driver.find_element(By.XPATH, '//*[@id="ctl00_MainContent_TxtCaptchaNumbers"]').click()

        # Wait for "Imprimir" button to appear after captcha verification
        while True:
            try:
                imprimir_button = driver.find_element(By.XPATH, '//input[@value="Imprimir"]')
                #imprimir_button.click()
                print(f"Preparing PDF for {file_name}")
                break
            except:
                time.sleep(1)

        # Generate the PDF using Chrome DevTools Protocol
        try:
            pdf_data = driver.execute_cdp_cmd("Page.printToPDF", {
                "landscape": False,
                "printBackground": True,
                "displayHeaderFooter": True,
                # Márgenes para dejar espacio a header/footer
                "marginTop":    0.5,  
                "marginBottom": 0.5,
                "marginLeft":   0.4,
                "marginRight":  0.4,
                # Plantilla de cabecera: fecha y título centrados
                "headerTemplate": """
                <div style="width:100%; text-align:center; font-size:10px; color: #444;">
                    <span class="date"></span>
                    &nbsp;&nbsp;|&nbsp;&nbsp;
                    <span class="title"></span>
                </div>
                """,
                # Plantilla de pie: URL y paginación centrados
                "footerTemplate": """
                <div style="width:100%; text-align:center; font-size:10px; color: #444;">
                    <span class="url"></span>
                    &nbsp;&nbsp;|&nbsp;&nbsp;
                    Página <span class="pageNumber"></span> de <span class="totalPages"></span>
                </div>
                """
            })
            pdf_content = base64.b64decode(pdf_data['data'])
        except Exception as e:
            print(f"Error generating PDF for {file_name}: {e}")
            continue

        # Set up the PDF path and save the file
        pdf_path = os.path.join(directory, f"{file_name}.pdf")
        try:
            with open(pdf_path, 'wb') as f:
                f.write(pdf_content)
            processed_count += 1
            print(f"Avance {processed_count}/{total_rows}: {file_name} - Saved to {pdf_path}")
        except IOError as e:
            print(f"Error saving PDF file for {file_name}: {e}")
            continue

        # Poll until the file is fully saved, rather than static wait
        wait_time = 0
        while not os.path.exists(pdf_path) and wait_time < 5:
            time.sleep(0.5)
            wait_time += 0.5

    driver.quit()
    print("All PDFs processed.")

# Main function to drive the program
def descargaQRSs(directory, chrome_driver):
    # Construct CSV file name based on the directory name
    directory_name = os.path.basename(directory)
    csv_file_name = os.path.join(directory, f"{directory_name}.csv")

    while True:
        # Mostrar menú
        print("\n¿Qué quieres hacer?")
        print("  1) Generar lista de QR's a descargar")
        print("  2) Descargar acuses del SAT")
        print("  3) Salir")
        
        choice = input("Elige una opción [1/2/3]: ").strip()
        
        if choice == '1':
            # Opción 1: generamos el CSV
            print("\nGenerando CSV...\n")
            STEP_A_generate(directory, csv_file_name)
            time.sleep(2)
        
        elif choice == '2':
            # Opción 2: intentamos leer y descargar
            data = read_csv(directory)
            if not data:
                print("\nNo localizamos CSV a descargar. Prueba a generarlo primero con la opción 1.")
                # Volver al menú
                continue
            
            print("\nDescargando los acuses del SAT...\n")
            download_pdfs(data, directory, chrome_driver)
        
        elif choice == '3':
            # Salir del bucle
            print("\nSaliendo. ¡Hasta luego!")
            break
        
        else:
            # Validación de entrada
            print("\nOpción inválida. Por favor ingresa 1, 2 o 3.")
            continue

