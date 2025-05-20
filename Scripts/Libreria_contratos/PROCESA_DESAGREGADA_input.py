
import pandas as pd
import re

def standarized_dataframe_generation(df, standarized_columns):
    """
    Interactively map the raw DataFrame’s columns to your standarized list,
    rename them, select only those columns, and optionally group by one.
    
    Parameters
    ----------
    df : pandas.DataFrame
        The raw input DataFrame.
    standarized_columns : list of str
        The target column names you want.
    group_column : str or None
        If provided, will group the result by this column and sum numeric cols.
    
    Returns
    -------
    pandas.DataFrame
        A DataFrame with columns renamed to standarized_columns (and optionally grouped).
    """
    # Maping columns
    def sanitize(name):
        return re.sub(r'[\W_]+', ' ', name).strip().lower()
    df.columns = [sanitize(col) for col in df.columns]
    raw_columns = list(df.columns)    
    mapping = {}
    print("Raw columns available:", raw_columns)

    for std_col in standarized_columns:
        while True:
            choice = input(f"→ Which raw column maps to '{std_col}'? ")
            #choice = sanitize(choice)
            if choice in raw_columns:
                mapping[choice] = std_col
                break
            print(f"  ✗ '{choice}' not in {raw_columns!r}")

    # Apply renaming
    #   1. build an empty DataFrame and copy over each mapped column
    df_std = pd.DataFrame()
    for raw_col, std_col in mapping.items():
        df_std[std_col] = df[raw_col]

    #   2. ensure the columns are in the exact standard order
    df_std = df_std[standarized_columns]
    for col in df_std.select_dtypes(include=['object']).columns:
        df_std[col] = df_std[col].str.strip()
    # Group the dataframe
    return df_std
