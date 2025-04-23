import os
import re

def pretty_xml(xml_string):
    # Use regex to add a newline after each closing tag
    return re.sub(r'>(?=<cfdi:)', '>\n', xml_string)

def pretty_xml_files(directory):
    for filename in os.listdir(directory):
        if filename.endswith('.xml'):
            filepath = os.path.join(directory, filename)
            with open(filepath, 'r', encoding='utf-8') as file:
                content = file.read()
                pretty_content = pretty_xml(content)
                
            with open(filepath, 'w', encoding='utf-8') as file:
                file.write(pretty_content)
            print(f"XML estructurado: {filename}")


def replace_in_file(filename):
    """Replace the target string in the given file."""
    with open(filename, 'r', encoding='utf-8') as file:
        content = file.read()
    
    new_content = content.replace('MetodoPago="PPD"', 'MetodoPago="PUE"')

    with open(filename, 'w', encoding='utf-8') as file:
        file.write(new_content)

def XML_PUE(directory):
    folder_path = directory

    # Get all files in the specified directory
    files_in_directory = os.listdir(folder_path)
    filtered_files = [os.path.join(folder_path, file) for file in files_in_directory if file.endswith(".xml")]

    for file in filtered_files:
        replace_in_file(file)
        print(f"PUE Actualizado en el XML {os.path.basename(file)}")