import os
import os

def clean_names(source_folder):
    """
    Rename files by keeping only the text before the first space (preserving extension),
    but skip any file whose name starts with the folder’s own basename.
    """
    skip_prefix = os.path.basename(source_folder)

    for filename in os.listdir(source_folder):
        # 1) Skip files that start with the folder name
        if filename.startswith(skip_prefix):
            continue

        # 2) Split the filename at the first space
        parts = filename.split(' ')
        new_filename = parts[0]

        # 3) Handle extension if present in the last part
        if '.' in parts[-1]:
            # If the first part already contains an extension, keep as is
            if '.' not in new_filename:
                extension = parts[-1].split('.')[-1]
                new_filename = f"{new_filename}.{extension}"

        # 4) Build full paths and rename if different
        src = os.path.join(source_folder, filename)
        dst = os.path.join(source_folder, new_filename)

        if src != dst:
            os.rename(src, dst)
            print(f"Renamed '{filename}' to '{new_filename}'")
