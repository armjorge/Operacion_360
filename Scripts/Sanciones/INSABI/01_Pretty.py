import os
import re

def pretty_xml(xml_string):
    # Use regex to add a newline after each closing tag
    return re.sub(r'>(?=<cfdi:)', '>\n', xml_string)

def process_files_in_directory(directory):
    for filename in os.listdir(directory):
        if filename.endswith('.xml'):
            filepath = os.path.join(directory, filename)
            with open(filepath, 'r', encoding='utf-8') as file:
                content = file.read()
                pretty_content = pretty_xml(content)
                
            with open(filepath, 'w', encoding='utf-8') as file:
                file.write(pretty_content)
            print(f"Processed {filename}")

if __name__ == "__main__":
    folder_name = input("Enter the folder name containing XML files: ")
    process_files_in_directory(folder_name)

