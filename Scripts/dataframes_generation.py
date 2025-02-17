import os
import pandas as pd

def create_dataframe(extension, dataframe_name, columns, output_folder):
    """Creates or loads a dataframe with expected columns, ensuring correct structure."""
    
    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Define file path based on extension
    file_path = os.path.join(output_folder, f"{dataframe_name}.{extension}")

    # Check if the file exists
    if os.path.exists(file_path):
        if extension == "pickle":
            df = pd.read_pickle(file_path)
        elif extension == "xlsx":
            df = pd.read_excel(file_path, dtype=columns)  # Ensure types match YAML definitions
        else:
            print(f"❌ La extensión {extension} no está considerada en este algoritmo.")
            return None

        # Validate column names
        if list(df.columns) == list(columns.keys()):
            print(f"✅ Dataframe {dataframe_name} localizado.")
            return df  # Return the loaded dataframe
        else:
            print(f"❌ ERROR: {dataframe_name} tiene las columnas {list(df.columns)} y esperábamos {list(columns.keys())}.")
            return None

    # If the file does not exist, create a new DataFrame
    df = pd.DataFrame(columns=columns.keys()).astype(columns)

    # Save according to the specified format
    if extension == "pickle":
        df.to_pickle(file_path)
        print(f"🆕 {dataframe_name}.{extension} creado y guardado en {os.path.basename(output_folder)}.")
    elif extension == "xlsx":
        df.to_excel(file_path, index=False)
        print(f"🆕 {dataframe_name}.{extension} creado y guardado en {os.path.basename(output_folder)}.")
    else:
        print(f"❌ La extensión {extension} no está considerada en este algoritmo.")
        return None

    return df