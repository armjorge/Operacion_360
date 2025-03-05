import os
import PyPDF2

def get_pdf_filenames(folder_path, file_prefixes):
    """Generate filenames for each prefix with different suffixes."""
    suffixes = ['', '_SAT', '_TXT']
    file_list = []
    for prefix in file_prefixes:
        for suffix in suffixes:
            filename = f"{prefix}{suffix}.pdf"
            full_path = os.path.join(folder_path, filename)
            if os.path.exists(full_path):
                file_list.append(full_path)
    return file_list

def merge_pdfs(file_list, output_filename):
    """Merge multiple PDFs into one."""
    merger = PyPDF2.PdfMerger()
    
    for file in file_list:
        print(f"Adding file to merger: {file}")  # Debugging line
        try:
            with open(file, 'rb') as pdf:
                merger.append(pdf)
        except Exception as e:
            print(f"Failed to add {file}: {e}")  # Print any errors during the merge process

    try:
        with open(output_filename, 'wb') as output_pdf:
            merger.write(output_pdf)
    except Exception as e:
        print(f"Failed to write merged PDF: {e}")  # Print any errors during writing the merged PDF

if __name__ == "__main__":
    folder_path = input("Enter the folder path: ")
    file_prefixes_input = input("Enter the file list (comma-separated): ")
    file_prefixes = [prefix.strip() for prefix in file_prefixes_input.split(', ')]

    # Generate filenames based on the folder and prefixes
    filenames = get_pdf_filenames(folder_path, file_prefixes)

    # Merge PDFs based on the generated filenames
    output_filename = os.path.basename(folder_path) + "_merged.pdf"
    merge_pdfs(filenames, output_filename)
    print(f"Merged PDFs into {output_filename}")
