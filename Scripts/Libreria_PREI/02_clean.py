import os
import re

# Define your source folders
source_folders = ['./XMLs', './PDFs']

for source_folder in source_folders:
    # Iterate over all files in each directory
    for filename in os.listdir(source_folder):
        # Check if the file name matches the pattern
        if re.match(r'P-\d{4}', filename):
            # Search for the "P-####" part and the extension
            match = re.search(r'(P-\d{4}).*(\.\w+)$', filename)
            # Check if a match was found
            if match:
                # Extract the matched groups
                new_filename = ''.join(match.groups())
                # Create source and destination file paths
                src = os.path.join(source_folder, filename)
                dst = os.path.join(source_folder, new_filename)
                # Check if the destination file already exists
                if not os.path.exists(dst):
                    # Rename the file
                    os.rename(src, dst)
                    print(f"Renamed {src} to {dst}")
                else:
                    print(f"File {dst} already exists. Skipping.")
