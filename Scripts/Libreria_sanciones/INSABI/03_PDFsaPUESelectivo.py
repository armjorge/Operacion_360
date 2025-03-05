import os
import subprocess
import platform


def A_open_pdfs_in_folder(folder_path):
    # Get all files in the directory
    files = os.listdir(folder_path)
    
    # Filter for PDFs
    pdf_files = [f for f in files if f.lower().endswith('.pdf')]
    pdf_files = A1_selectPDFS(pdf_files)
    for pdf in pdf_files:
        pdf_path = os.path.join(folder_path, pdf)
        if not os.path.exists(pdf_path):
            print(f"❌ Error: El archivo no existe: {pdf_path}")
            continue

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

            else:
                print("Unsupported OS.")
                continue

        except Exception as e:
            print(f"⚠️ No se pudo abrir el archivo: {e}")

        input(f"{pdf} has been opened. Press Enter after you close the file to proceed to the next one...")

    print("All PDFs have been processed.")


def A1_selectPDFS(PDF_filelist):
    print("What do we want to open sequentially?")
    print("1. All the PDF files")
    print("2. Only QR PDF")
    choice = input("Enter your choice (1 or 2): ").strip()
    
    if choice == "1":
        return PDF_filelist
    elif choice == "2":
        filtered_files = [
            f for f in PDF_filelist 
            if f.startswith("NC") and not (f.endswith("_SAT.pdf") or f.endswith("_TXT.pdf"))
        ]
        print("\nSummary of filtering logic:")
        print("Keeping files that start with 'NC' and do not end with '_SAT' or '_TXT'.")
        print(f"Files kept: {len(filtered_files)} out of {len(PDF_filelist)}")
        return filtered_files
    else:
        print("Invalid choice. Returning the original list.")
        return PDF_filelist

def main():
    print("\n**********\nAgrega: \nPUE - Pago en una sola exhibición\nA todas las Notas\n**********\n")
    directory_path = input("Nombre de la carpeta: ")
    A_open_pdfs_in_folder(directory_path)
    
if __name__ == "__main__":
    main()


