import os
import zipfile

def G_generatezip(folder_name, suffixes, preffix):
    """
    Genera un archivo ZIP que incluye, para cada nota de crédito única encontrada en la carpeta:
      - Cada archivo formado por la nota base + cada sufijo de `suffixes`
    Además, se incluye el archivo de formato INSABI.
    """
    # Validación de parámetros
    if not isinstance(suffixes, list) or not suffixes:
        print("❌ Error: 'suffixes' debe ser una lista no vacía.")
        return
    if not isinstance(preffix, str) or not preffix:
        print("❌ Error: 'preffix' debe ser una cadena no vacía.")
        return

    # Definir rutas
    base_name = os.path.basename(folder_name)
    base_dir = folder_name
    zip_path = os.path.join(base_dir, f"{base_name}.zip")
    formato_file = os.path.join(base_dir, f"{base_name}_formato_INSABI.xlsx")
    
    # Mensaje inicial con íconos
    print(f"🔍 Buscaremos que cada valor individual '{preffix}' tenga su correspondiente archivo {suffixes}.")

    # Obtener todos los archivos en la carpeta que comiencen con el preffix
    try:
        all_files = os.listdir(base_dir)
    except Exception as e:
        print(f"❌ Error al acceder a la carpeta {base_dir}: {e}")
        return

    credit_files = [f for f in all_files if f.startswith(preffix) and os.path.isfile(os.path.join(base_dir, f))]

    # Extraer las notas de crédito únicas eliminando los sufijos conocidos.
    unique_notes = set()
    # Ordenar los sufijos de mayor a menor longitud para remover primero el más específico.
    sorted_suffixes = sorted(suffixes, key=len, reverse=True)
    for file in credit_files:
        for suf in sorted_suffixes:
            if file.endswith(suf):
                base_name = file[:-len(suf)]
                unique_notes.add(base_name)
                break
        else:
            unique_notes.add(file)
    
    unique_notes = sorted(unique_notes)
    
    print(f"✅ Se han detectado {len(unique_notes)} Notas de crédito esperadas:")
    for note in unique_notes:
        print(f"   📄 {note}")
    
    # Construir la lista de archivos esperados para cada nota única y cada sufijo
    expected_files = []
    for note in unique_notes:
        for suf in suffixes:
            expected_files.append(os.path.join(base_dir, note + suf))
    
    # Incluir el archivo formato INSABI en la lista
    expected_files.append(formato_file)
    
    # Verificar la existencia de cada archivo esperado
    missing_files = []
    for file in expected_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print("⚠️ Algunos archivos requeridos están faltando:")
        for m in missing_files:
            print(f"   ❌ {m}")
        print("Por favor, asegúrese de que todos los archivos requeridos estén presentes en la carpeta.")
        return
    else:
        print("🎉 Todos los archivos requeridos se encuentran. Procediendo a crear el archivo ZIP...")
    
    # Crear el archivo ZIP
    try:
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in expected_files:
                zipf.write(file, os.path.basename(file))
        print(f"✅ ZIP creado exitosamente: {zip_path}")
    except Exception as e:
        print(f"❌ Ocurrió un error al crear el ZIP: {e}")

def genera_zip(folder_name):
    preffix = 'NC'
    suffixes = ['.pdf', '_SAT.pdf', '_TXT.pdf', '.xml']
    G_generatezip(folder_name, suffixes, preffix)
    print("\nPaso G Generación de Zip's completada\n")