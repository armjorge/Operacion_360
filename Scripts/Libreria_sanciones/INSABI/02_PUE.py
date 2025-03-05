import os

def replace_in_file(filename):
    """Replace the target string in the given file."""
    with open(filename, 'r', encoding='utf-8') as file:
        content = file.read()
    
    new_content = content.replace('MetodoPago="PPD"', 'MetodoPago="PUE"')

    with open(filename, 'w', encoding='utf-8') as file:
        file.write(new_content)

def main():
    folder_name = input("Enter the folder name containing XML files: ")
    folder_path = os.path.join('.', folder_name)

    # Get all files in the specified directory
    files_in_directory = os.listdir(folder_path)
    filtered_files = [os.path.join(folder_path, file) for file in files_in_directory if file.endswith(".xml")]

    for file in filtered_files:
        replace_in_file(file)
        print(f"Processed {file}")

if __name__ == "__main__":
    main()

