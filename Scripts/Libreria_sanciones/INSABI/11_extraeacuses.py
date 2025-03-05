import os
from PyPDF2 import PdfMerger

def check_acuse(root_folder):
    list_with_acuse = []
    list_without_acuse = []
    
    # Iterate over each folder in the root directory
    for folder_name in os.listdir(root_folder):
        folder_path = os.path.join(root_folder, folder_name)
        
        if os.path.isdir(folder_path):  # Only process if it's a directory
            acuse_file = f"{folder_name}_acuse.pdf"
            acuse_path = os.path.join(folder_path, acuse_file)
            
            if os.path.exists(acuse_path):
                list_with_acuse.append(folder_name)
            else:
                list_without_acuse.append(folder_name)
    
    return list_with_acuse, list_without_acuse

def export(list_without_acuse, root_folder):
    merger = PdfMerger()
    export_path = os.path.join(root_folder, 'Acuses_sanciones_INSABI.pdf')
    
    for folder_name in list_without_acuse:
        acuse_path = os.path.join(root_folder, folder_name, f"{folder_name}_acuse.pdf")
        
        # Add PDF if it exists (extra check if folders change during runtime)
        if os.path.exists(acuse_path):
            merger.append(acuse_path)
    
    # Write out the merged PDF
    merger.write(export_path)
    merger.close()
    
    # Open the folder containing the merged PDF file
    os.startfile(root_folder)  # Windows specific; adjust for other OS if needed
    print("Merged PDF generated at:", export_path)

def main():
    root_folder = os.getcwd()  # Set root folder to current working directory
    list_with_acuse, list_without_acuse = check_acuse(root_folder)
    
    print("Folders without 'acuse' PDF:", list_without_acuse)
    user_input = input("Do you want to export all the Acuses? (Yes/No): ")
    
    if user_input.lower() == "yes":
        export(list_with_acuse, root_folder)
    elif user_input.lower() == "no":
        print("Have a nice day!")
    else:
        print("Invalid input. Please run the program again.")

if __name__ == "__main__":
    main()
