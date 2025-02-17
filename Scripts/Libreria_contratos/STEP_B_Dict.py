
import pandas as pd
import ast
import calendar

def STEP_B_get_string_populated(df_clientes,  tipo, institucion_column, selected_procedimiento): 
    print("\t🛠️ Iniciando la generación del diccionario para el contrato...\n")
    # Crear el diccionario poblado
    orchestration_dict = f"""
    {{
        'Institución': "{STEP_B_populate_from_df(df_clientes, institucion_column)}",
        'Procedimiento': "{selected_procedimiento}",
        'Contrato': "{STEP_B_contrato(tipo, 'Contrato')}",
        'Modificatorio': "{STEP_B_contrato(tipo, 'Modificatorio') }",
        'Fecha Inicio': "{STEP_B_fechas(tipo, 'Fecha Inicio')}",
        'Fecha Fin': "{STEP_B_fechas(tipo, 'Fecha Fin')}",
        'Estatus': "{STEP_B_estatus()}",
        'SKU': ""
    }}
    """

    # Mostrar valores antes de devolver el diccionario
    print("\n🔹 Diccionario generado con valores actuales:")
    print(orchestration_dict)
    computer_dict = STEP_B_dict_validation(orchestration_dict)

    return orchestration_dict, computer_dict

def STEP_B_contrato(tipo, input_field):
    while True:
        if tipo == 'Primigenio' and input_field == 'Contrato':
            pre_contrato = input("Captura el nombre del contrato: ")
            print(pre_contrato)
            respuesta = input("\nConfirmas que es el contrato? Sí o No: ").strip().lower()[0]
            if respuesta == 's':
                return pre_contrato
            elif respuesta == 'n':
                continue  # This will restart the loop if the answer is no

        elif tipo == 'Primigenio' and input_field == 'Modificatorio':
            return ""  # Return empty string for Modificatorio when tipo is Primigenio

        elif tipo == 'Modificatorio' and input_field == 'Modificatorio':
            pre_modificatorio = input("Captura el código del modificatorio: ")
            print(pre_modificatorio)
            respuesta = input("\nConfirmas que es el Modificatorio? Sí o No: ").strip().lower()[0]
            if respuesta == 's':
                return pre_modificatorio
            elif respuesta == 'n':
                continue  # This will restart the loop if the answer is no
        elif tipo == 'Modificatorio' and input_field == 'Contrato':
            print(f"Es un {tipo} por lo que ya debemos tener un contrato cargado, accediendo a los datos previos.")
            break  # Assuming further logic or return is handled here
        else:
            print(f"\n\tCombinación de campo {input_field} y tipo {tipo} no considerado en el arbol de decisiones de Contratos, saliendo del código sin return variable\n")
            break

def input_captura_fechas():
    input_day = "¿Cuál es día? (2 dígitos)?: "            
    input_mes = "¿Cuál es mes? (2 dígitos)?: "            
    input_year = "¿Cuál es el año(4 dígitos)?: "
    mm_dd_digitos = 2
    mes_enero = 1
    mes_diciembre = 12
    year_digits = 4
    year_min = 2025
    year_max = 2060
    def get_specific_digits_as_string(prompt, needed_digits, min_input, max_input):
        """
        Prompts the user for input and validates it using specific conditions:
        - Input must be numeric.
        - Input must have the specified number of digits.
        - Input must fall within the specified range.
        """
        while True:
            user_input = input(prompt)
            error_message = "Vuelve a intentar"
            if user_input.isdigit() and len(user_input) == needed_digits and min_input <= int(user_input) <= max_input:
                return user_input
            print(f"{error_message}, necesitas meter un número de mín de {min_input} y máx {max_input} caracteres")    
    mes = get_specific_digits_as_string(input_mes, mm_dd_digitos, mes_enero, mes_diciembre)
    year = get_specific_digits_as_string(input_year, year_digits, year_min, year_max)
    day_min = 1
    # Función para obtener los días máximos de ese mes. 
    def get_month_days(year, month):
        """
        Returns the maximum number of days in a given month for a specific year.
        
        Parameters:
            year (str or int): The year in 4-digit format.
            month (str): The month in 2-digit string format.

        Returns:
            int: Maximum number of days in the specified month.
        """

        # Convert inputs to integers
        year = int(year)
        month = int(month)

        # Get the last day of the month
        day_max = calendar.monthrange(year, month)[1]

        return day_max
    day_max = get_month_days(year,mes)
    day = get_specific_digits_as_string(input_day, mm_dd_digitos, day_min, day_max)            
    date = f"{day}/{mes}/{year}"
    return date

def STEP_B_fechas(tipo,input_field):
    while True:
        print("Capturaremos la fecha de inicio del contrato")
        if tipo == 'Primigenio' and input_field == 'Fecha Inicio':
            fecha_inicio = input_captura_fechas()
            return fecha_inicio
        elif tipo == 'Primigenio' and input_field == 'Fecha Fin':
            print("Capturaremos la fecha de terminación del contrato")
            fecha_final = input_captura_fechas()
            return fecha_final
        elif tipo == 'Modificatorio' and input_field == 'Fecha Inicio':
            print("Árbol de decisión para {tipo} y {input_field} en desarrollo")
            break
        elif tipo == 'Modificatorio' and input_field == 'Fecha Fin':
            print("Árbol de decisión para {tipo} y {input_field} en desarrollo")
            break
        else:
            print(f"\n\tCombinación de campo {input_field} y tipo {tipo} no considerado en el arbol de decisiones de Contratos, saliendo del código sin return variable\n")
            break

def STEP_B_estatus():
    while True:
        try:
            input_user = int(input("\n1) Formalizado o 2) Copia con firmas incompletas (no formalizado): "))
            if input_user == 1:
                return "Formalizado"
            elif input_user == 2:
                return "Copia con firmas incompletas"
            else:
                print("Opción no válida, intenta de nuevo.")
        except ValueError:
            print("Por favor, ingresa un número.")

def STEP_B_dict_validation(dicct_book):
    try:
        # Step 1: Trim leading/trailing spaces
        dicct_book = dicct_book.strip()

        # Step 2: Fix common syntax issues
        dicct_book = dicct_book.replace("‘", "'").replace("’", "'")  # Smart quotes to normal quotes
        dicct_book = dicct_book.replace("“", '"').replace("”", '"')  # Smart double quotes fix
        
        # Step 3: Count brackets
        open_brackets = dicct_book.count('{')
        close_brackets = dicct_book.count('}')
        
        if open_brackets != close_brackets:
            print(f"⚠️ Desajuste en corchetes: {open_brackets} abiertos, {close_brackets} cerrados.")
            return None

        # Step 4: Try parsing the string into a dictionary
        cleaned_dict = ast.literal_eval(dicct_book)

        if not isinstance(cleaned_dict, dict):
            raise ValueError("⚠️ Error: La estructura no es un diccionario válido.")

        print("✅ Diccionario validado y corregido correctamente.")
        return cleaned_dict

    except (SyntaxError, ValueError) as e:
        print(f"❌ Error al analizar el diccionario: {e}")
        print("⚠️ Revisa los corchetes, comillas y la estructura del diccionario.")
        
        # Suggest a manual fix
        print("\n🔹 Sugerencia: Asegúrate de que el diccionario esté bien formateado:")
        print("{ 'Clave': 'Valor', 'Otra Clave': 'Otro Valor' }")
        return None    

def STEP_B_populate_from_df(df_to_load, column):
    """
    Handles loading a DataFrame from a pickle file, ensuring the specified column exists.
    Allows the user to input or select values.
    
    Parameters:
        df_to_load (str): Path to the pickle file.
        column (str): Column name to retrieve data from.
    
    Returns:
        Selected value from the column.
    """

    # Check if the pickle file exists
    """
    if not os.path.exists(df_to_load):
        print("❌ Archivo no localizado, procedemos a crear uno.")
        df = pd.DataFrame()  # Create an empty DataFrame
        df.to_pickle(df_to_load)
        print("✅ Archivo pickle creado.")
    """
    # Load the DataFrame
    #df = pd.read_pickle(df_to_load)
    if not df_to_load.empty:
        print("✅ Archivo con información cargada.")

    # Ensure the column exists in the DataFrame
    if column not in df_to_load.columns:
        print(f"🔹 Columna '{column}' no encontrada. Abre el excel y generar una nueva columna.")
        #df[column] = pd.Series(dtype="object")  # Add an empty column
        #df.to_pickle(df_to_load)  # Save updated DataFrame
        print(f"✅ Nueva columna '{column}' por crear.")

    # Check if there are any values in the column
    if len(df_to_load[column].dropna()) > 0:
        print("✅ Se encontró al menos un registro:")
        unique_values = df_to_load[column].dropna().unique()
        #for idx, value in enumerate(unique_values[:3]):  
        for idx, value in enumerate(unique_values[:3]):
            print(f"\t{idx}) {value}")
    else:
        print(f"❌ No se encontraron registros en la columna {column}.")

    # Main loop for user input
    while True:
        print("\n📌 Opciones:")
        #print(f"\t1) Agregar más valores o actualizar en la columna {column}")
        print(f"\tSeleccionar un valor existente en la columna {column}")
        
        try:
            choice = 2 #int(input("Selecciona una opción (1/2): "))
        except ValueError:
            print("⚠️ Entrada no válida. Ingresa 1 o 2.")
            continue

        if choice == 1:
            # Save the column to a CSV for user input
            csv_path = os.path.join(os.path.dirname(df_to_load), f"{column}.csv")
            df_to_load[[column]].to_csv(csv_path, index=False)
            print(f"📂 Se generó el archivo: {csv_path}")
            print("📂 Búscalo en tu CODE o desarrolla código para abrirlo en cualquier sistema")
            input("📝 Agrega nuevos valores en el archivo CSV y presiona ENTER cuando termines...")

            # Reload the updated CSV
            updated_df = pd.read_csv(csv_path)

            df_to_load[column] = updated_df[column]  # Replace column data
            #df_to_load.to_pickle(df_to_load)  # Save changes
            print("✅ DataFrame maestro actualizado con estos datos: ")
            print("\n🔹 Valores disponibles:")
            unique_values = df_to_load[column].dropna().unique()
            for idx, value in enumerate(unique_values):
                print(f"\t{idx}) {value}")
            os.remove(csv_path)
            print("✅ CSV de la captura eliminado")
                            
        elif choice == 2:
            # Let the user choose a value from the list
            unique_values = df_to_load[column].dropna().unique()
            if len(unique_values) == 0:
                print("⚠️ No hay valores disponibles para seleccionar.")
                continue
            
            print("\n🔹 Valores disponibles:")
            for idx, value in enumerate(unique_values):
                print(f"\t{idx}) {value}")

            try:
                index = int(input("🔹 Dame el índice del valor: "))
                if index in range(len(unique_values)):
                    return unique_values[index]
                else:
                    print("⚠️ Índice fuera de rango.")
            except ValueError:
                print("⚠️ Entrada no válida. Ingresa un número válido.")

        else:
            print("⚠️ Opción no válida. Intenta de nuevo.")