from math import ceil
import os

def get_folders_pending(sanciones_folder):
    pending_letters = []
    # iterate over entries in the sanciones_folder
    for folder_name in os.listdir(sanciones_folder):
        folder_path = os.path.join(sanciones_folder, folder_name)
        # skip anything that isn’t a directory
        if not os.path.isdir(folder_path):
            continue
        # build the expected PDF filename
        expected_file = f"{folder_name}_acuse.pdf"
        # check whether it exists inside the folder
        if not os.path.exists(os.path.join(folder_path, expected_file)):
            pending_letters.append(folder_name)
    return pending_letters

def print_columns(lst, n_cols=4, col_width=None):
    """
    Print the items of lst in n_cols columns.
    If col_width is None, it will be set to (longest_item_length + 2).
    """
    if col_width is None:
        maxlen = max(len(str(item)) for item in lst)
        col_width = maxlen + 2

    n = len(lst)
    rows = ceil(n / n_cols)

    for r in range(rows):
        line = ""
        for c in range(n_cols):
            idx = c * rows + r
            if idx < n:
                line += f"{lst[idx]:<{col_width}}"
        print(line)
