import subprocess
import subprocess
import platform
import os
import pandas as pd
from IPython.display import display
import re

def sanitize_filename(filename):
    """Replace invalid characters with '-' and keep only allowed characters."""
    sanitized = re.sub(r'[^a-zA-Z0-9_-]', '-', filename)
    return sanitized

def trim_to_limit(filename, limit=255):
    """Trim filename to the allowed limit, prioritizing digits from the contract value."""
    if len(filename) <= limit:
        return filename
    # Prioritize digits from the contract value
    digits = ''.join(filter(str.isdigit, filename))
    if len(digits) > limit:
        return digits[:limit]
    return digits + filename[len(digits):][:limit - len(digits)]
def open_excel(path):
    """Open an Excel file based on the operating system."""
    if not os.path.exists(path):
        print(f"Error: El archivo '{path}' no existe.")
        return

    try:
        if platform.system() == "Windows":
            os.system(f'start excel "{path}"')
        elif platform.system() == "Darwin":  # macOS
            os.system(f'open -a "Microsoft Excel" "{path}"')
        else:
            print("Sistema operativo no soportado para abrir Excel.")
            return

        print(f"{os.path.basename(path)} abierto con éxito.")
    except Exception as e:
        print(f"Error al intentar abrir el archivo: {e}")

def load_dataframe(path, sheet, columns):
    # Check if the file path exists
    if not os.path.exists(path):
        # Create a new DataFrame and save it as an Excel file if it does not exist
        df = pd.DataFrame(columns=columns)
        with pd.ExcelWriter(path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name=sheet)
        print(f"El {os.path.basename(path)} no existía, creado.")
        return df
    else:
        # Load the DataFrame from the specified Excel sheet if the file exists
        df_loaded = pd.read_excel(path, sheet_name=sheet)
        return df_loaded


def open_pdf(pdf_path):
    """
    Opens the given PDF file in a system-compatible way.
    - Tries to use Adobe Acrobat if available.
    - Otherwise, uses the default PDF viewer.

    Parameters:
        pdf_path (str): The full path to the PDF file.

    Returns:
        None
    """
    if not os.path.exists(pdf_path):
        print(f"❌ Error: El archivo no existe: {pdf_path}")
        return

    system = platform.system()

    try:
        if system == "Windows":
            # Try opening with Acrobat Reader
            acrobat_path = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
            if os.path.exists(acrobat_path):
                subprocess.run([acrobat_path, pdf_path], check=False)
            else:
                os.startfile(pdf_path)  # Open with default PDF viewer
        elif system == "Darwin":  # macOS
            acrobat_path = "/Applications/Adobe Acrobat DC/Adobe Acrobat.app"
            if os.path.exists(acrobat_path):
                subprocess.run(["open", "-a", acrobat_path, pdf_path], check=False)
            else:
                subprocess.run(["open", pdf_path], check=False)  # Default PDF viewer
        elif system == "Linux":
            subprocess.run(["xdg-open", pdf_path], check=False)

    except Exception as e:
        print(f"⚠️ No se pudo abrir el archivo: {e}")

def open_folder(os_path):
    """Opens a folder in the appropriate file explorer depending on the OS."""
    try:
        if os.name == 'nt':  # Windows
            os.startfile(os_path)
        elif os.name == 'posix':  # macOS or Linux
            if "darwin" in os.uname().sysname.lower():  # macOS
                subprocess.run(["open", os_path])
            else:  # Linux
                subprocess.run(["xdg-open", os_path])
        else:
            print(f"Unsupported OS: {os.name}")
    except Exception as e:
        print(f"Error opening folder: {e}")

def create_directory_if_not_exists(path):
    """Creates a directory if it does not exist and prints in Jupyter."""
    if not os.path.exists(path):
        print(f"\nNo se localizó el folder {os.path.basename(path)}, creando.", flush=True)
        os.makedirs(path)
        print(f"\tFolder {os.path.basename(path)} creado.", flush=True)
    else:
        print(f"\tFolder {os.path.basename(path)} encontrado.", flush=True)

