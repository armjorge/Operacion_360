from math import ceil

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
