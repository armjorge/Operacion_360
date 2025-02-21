import os

# Ask user for the source folder
source_folder = input("Enter the folder name: ")

# Iterate over all files in the directory
for filename in os.listdir(source_folder):
    # Split the filename at the first space
    parts = filename.split(' ')
    new_filename = parts[0]

    # Check if the original filename had an extension
    if '.' in parts[-1]:
        # If the first part already contains the extension, use it as is
        if '.' in new_filename:
            new_filename = new_filename
        else:
            # Otherwise, append the extension from the last part
            extension = parts[-1].split('.')[-1]
            new_filename = f"{new_filename}.{extension}"

    # Create source and destination file paths
    src = os.path.join(source_folder, filename)
    dst = os.path.join(source_folder, new_filename)

    # Rename the file if the new filename is different
    if src != dst:
        os.rename(src, dst)
        print(f"Renamed '{filename}' to '{new_filename}'")
