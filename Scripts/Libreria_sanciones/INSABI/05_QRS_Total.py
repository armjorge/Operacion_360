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



def generate(directory):
    # Construct CSV file name based on the directory name
    csv_file_name = os.path.join(directory, f"{directory}.csv")
    
    # Build a list of already processed files (those ending with _SAT.pdf are assumed downloaded)
    directory_downloaded_files = [
        f[:-4] for f in os.listdir(directory) 
        if not (f.endswith("_TXT.pdf") or f.endswith("_REM.pdf"))
    ]

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
                file_path = os.path.join(directory, filename)
                try:
                    doc = fitz.open(file_path)
                except Exception as e:
                    print(f"Error opening {filename}: {e}")
                    continue

                for page_number in range(len(doc)):
                    # Get all image objects on the page
                    for img in doc.get_page_images(page_number):
                        xref = img[0]
                        base = os.path.splitext(filename)[0]
                        modified_filename = f"{base}_SAT"  # Append _SAT

                        if modified_filename in directory_downloaded_files:
                            continue

                        try:
                            pix = fitz.Pixmap(doc, xref)
                            # Create a NumPy array from the pixmap samples.
                            # Note: pix.samples is a bytes object containing all pixel data.
                            # However, each row in the image is padded to a length given by pix.stride.
                            # We first create a 1D array then reshape to (height, stride).
                            raw = np.frombuffer(pix.samples, dtype=np.uint8)
                            raw = raw.reshape((pix.height, pix.stride))
                            
                            # Only the first (width * number_of_components) bytes in each row are valid
                            valid_length = pix.width * pix.n
                            raw = raw[:, :valid_length]
                            
                            # Reshape into (height, width, n) where n is number of color components (3 or 4)
                            raw = raw.reshape((pix.height, pix.width, pix.n))
                            
                            # Convert RGB or RGBA to BGR for OpenCV
                            if pix.n == 4:
                                cv_img = cv2.cvtColor(raw, cv2.COLOR_RGBA2BGR)
                            elif pix.n == 3:
                                cv_img = cv2.cvtColor(raw, cv2.COLOR_RGB2BGR)
                            else:
                                cv_img = raw  # fallback
                        except Exception as e:
                            print(f"Error converting image from {filename}: {e}")
                            continue

                        # Use OpenCV's QRCodeDetector to detect and decode QR codes
                        detector = cv2.QRCodeDetector()
                        qr_data, points, _ = detector.detectAndDecode(cv_img)
                        if qr_data:
                            writer.writerow([modified_filename, qr_data])
                        else:
                            print(f"No QR found in image from {filename} (page {page_number})")
                        
                        pix = None  # Free memory

                print(f"Decoded QR code data from {filename}")
    print(f"Completed writing to {csv_file_name}")

    # Load the generated CSV to check if it has content
    df_uploadMergelist = pd.read_csv(csv_file_name)

    # Check if the DataFrame is empty and print the appropriate message
    if df_uploadMergelist.empty:
        print("\n*******************************\nNo hay QRs por descargar \n*******************************\n")
    else:
        print(f"\n*******************************\nHay {len(df_uploadMergelist)} facturas por descargar. \n*******************************\n")
""" 
def load_chrome(directory):


    chrome_options = Options()
    chrome_binary_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    chrome_options.binary_location = chrome_binary_path

    prefs = {
        "download.default_directory": os.path.abspath(directory),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")

    try:
        driver = webdriver.Chrome(options=chrome_options)
        return driver
    except Exception as e:
        print(f"Failed to initialize Chrome driver: {e}")
        return None
"""

def load_chrome(directory):
    """Launch Chrome with OS-specific paths and consistent configuration."""

    # Detect OS
    system = platform.system()

    # Set Chrome binary and ChromeDriver paths based on OS
    if system == "Windows":
        chrome_binary_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        chromedriver_path = "C:\\Program Files\\Google\\Chrome\\chromedriver.exe"
    elif system == "Darwin":  # macOS
        home = os.path.expanduser("~")
        chrome_binary_path = os.path.join(home, "chrome_testing", "chrome-mac-arm64", "Google Chrome for Testing.app", "Contents", "MacOS", "Google Chrome for Testing")
        chromedriver_path = os.path.join(home, "chrome_testing", "chromedriver-mac-arm64", "chromedriver")
    else:
        print(f"Unsupported OS: {system}")
        return None

    # Set Chrome options
    chrome_options = Options()
    chrome_options.binary_location = chrome_binary_path

    prefs = {
        "download.default_directory": os.path.abspath(directory),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")

    try:
        # Initialize ChromeDriver with the correct service path
        service = Service(chromedriver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        print(f"Failed to initialize Chrome driver: {e}")
        return None

# Function to read the CSV file in the specified directory
def read_csv(directory):
    print("\nReading CSV\n")
    file_path = os.path.join(directory, f"{directory}.csv")
    
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
def download_pdfs(data, directory):
    driver = load_chrome(directory)  # Initialize Chrome driver
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
                "printBackground": True
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
def main():
    directory = input("Enter the folder name: ")
    if not os.path.exists(directory):
        os.makedirs(directory)
    print("generating CSV\n")
    generate(directory)  # Generate the CSV
    time.sleep(2)    
    data = read_csv(directory)
    if not data:
        print("No data found. Exiting.")
        return

    
    download_pdfs(data, directory)

if __name__ == "__main__":
    main()
