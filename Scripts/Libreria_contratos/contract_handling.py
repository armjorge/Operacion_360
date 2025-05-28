


def generar_skus(df_clientes, institucion_elegida, folder_path):
    print("🚀 Comenzando la función para capturar máximos por contrato")
    #Index(['Institución', 'Procedimiento', 'Clave', 'Descripción', 'Precio', 'Total']
    #df_contrato= df_clientes[['Institución', 'Clave', 'Precio', 'Total']]['Institución'] == institucion_column
    df_contrato = df_clientes[df_clientes['Institución'] == institucion_elegida][['Institución', 'Clave', 'Precio', 'Total']]
    print("Columna institucional", institucion_elegida)
    print("Retira las claves que no se encuentren en el contrato y modifica el total\nEstá basado en la demanda desagregada")
    print(df_contrato.head())
    # Define the temporary Excel file path
    temporal_excel_file = os.path.join(folder_path, f"Máximos {institucion_elegida}.xlsx")

    # Delete the file if it already exists
    if os.path.exists(temporal_excel_file):
        os.remove(temporal_excel_file)
        print("Eliminando archivo previo.")

    # Save the filtered DataFrame to an Excel file
    df_contrato.to_excel(temporal_excel_file, index=False)
    print(f"Archivo '{os.path.basename(temporal_excel_file)}' creado con éxito.")

    while True:
        open_excel(temporal_excel_file)
        input(f"💡 Por favor, carga los datos de SKU en '{os.path.basename(temporal_excel_file)}' y presiona Enter cuando hayas terminado...")

        try:
            sku_df = pd.read_excel(temporal_excel_file)
            if sku_df.empty:
                print("⚠️ El archivo de Excel está vacío. Por favor, llena los datos requeridos.")
                continue

            print("\n📊 Verifica las cantidades máximas por clave:")
            print(sku_df[['Clave', 'Total']].head())

            is_complete = input("✅ Las cantidades máximas por clave son correctas? (s/n): ").strip().lower()

            if is_complete == "s":
                print("✔️ Confirmación completada con éxito.")
                # Generate SKU string and return inside the loop
                sku_string = ", ".join(
                    f"{{'Clave': '{row['Clave']}', 'Precio': {row['Precio']}, 'Máximo': {row['Total']}}}"
                    for _, row in sku_df.iterrows()
                )
                return f"[{sku_string}]"

            elif is_complete == "n":
                print("🔄 Por favor, actualiza el archivo y vuelve a intentar.")
            else:
                print("⚠️ Respuesta no válida. Introduce 's' para sí o 'n' para no.")

        except Exception as e:
            print(f"❌ Error al cargar el archivo de Excel: {e}")
           