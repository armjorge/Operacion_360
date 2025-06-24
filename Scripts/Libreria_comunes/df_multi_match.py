import pandas as pd

def multi_column_lookup(df_to_fill, df_to_consult, match_columns: dict, return_column, default_value='sin match'):
    """
    Realiza búsqueda con múltiples columnas y retorna valor o advertencias
    Args:
        df_to_fill (pd.DataFrame): DataFrame que queremos llenar.
        df_to_consult (pd.DataFrame): DataFrame que consultamos.
        match_columns (dict): {col_df_to_fill: col_df_to_consult} pares de columnas para hacer match.
        return_column (str): Columna en df_to_consult con el valor a traer.
        default_value (any): Valor si no hay coincidencia.
    Returns:
        pd.Series: Valores para poblar la columna deseada.
    """

    if not isinstance(df_to_consult, pd.DataFrame):
        raise TypeError(f"El argumento 'df_to_consult' debe ser un DataFrame, se recibió: {type(df_to_consult)}")

    result_list = []
    ligas_duplicadas = 0
    ligas_vacías = 0
    for _, row in df_to_fill.iterrows():
        # Construir filtro booleano dinámico
        mask = pd.Series([True] * len(df_to_consult))

        for source_col, consult_col in match_columns.items():
            mask &= df_to_consult[consult_col] == row[source_col]

        filtered = df_to_consult[mask]

        if filtered.empty:
            result_list.append(default_value)
            #print("\tEncontramos renglones sin poder ligar")
            ligas_vacías += 1
        elif len(filtered) > 1:
            resultados_duplicados = ", ".join(filtered[return_column].astype(str).tolist())
            result_list.append(f'Duplicados:  {resultados_duplicados}')
            ligas_duplicadas += 1
        else:
            result_list.append(filtered.iloc[0][return_column])
    success_message = "✅ Se ligaron el 100% de los renglones y no hubo duplicados."
    if ligas_duplicadas == 0 and ligas_vacías == 0:
        print(f"\n{'*'*len(success_message)}\n{success_message}\n{'*'*len(success_message)}\n")
    elif ligas_duplicadas == 0 and ligas_vacías > 0:
        print(f"\n⚠️ Hay {ligas_vacías} renglones que no se pudieron llenar con el dataframe consultado.\n")
    elif ligas_duplicadas > 0 and ligas_vacías == 0:
        print(f"\n⚠️ Hay renglones {ligas_duplicadas} para los que se encontraron más de un resultado en el dataframe de consulta.\n")
    else:
        print(f"\n⚠️ Hay {ligas_vacías} renglones vacíos y {ligas_duplicadas}  renglones con más de un resultado.\n")

    return pd.Series(result_list, index=df_to_fill.index)