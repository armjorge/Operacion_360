import os
import pandas as pd
from lxml import etree


def smart_xml_extraction(invoice_paths, xlsx_database):
    # Si la base existe, la cargamos; si no, creamos el DataFrame con todas las columnas, incluida la nueva 'Fecha'
    if os.path.exists(xlsx_database):
        df_database = pd.read_excel(xlsx_database)
    else:
        df_database = pd.DataFrame(columns=[
            'UUID', 'Folio', 'Fecha', 'Nombre', 'Rfc',
            'Descripcion', 'Cantidad', 'Importe', 'Archivo'
        ])

    data = []

    for folder in invoice_paths:
        print(f"\nExplorando carpeta: {folder}")
        
        for root_dir, dirs, files in os.walk(folder):
            for file in files:
                if not file.endswith('.xml'):
                    continue
                full_path = os.path.join(root_dir, file)

                try:
                    tree = etree.parse(full_path)
                    root_element = tree.getroot()

                    # Detectar namespace CFDI y TimbreFiscalDigital
                    ns = None
                    for ns_url in root_element.nsmap.values():
                        if "cfd/3" in ns_url:
                            ns = {
                                "cfdi": "http://www.sat.gob.mx/cfd/3",
                                "tfd": "http://www.sat.gob.mx/TimbreFiscalDigital"
                            }
                            break
                        elif "cfd/4" in ns_url:
                            ns = {
                                "cfdi": "http://www.sat.gob.mx/cfd/4",
                                "tfd": "http://www.sat.gob.mx/TimbreFiscalDigital"
                            }
                            break
                    if ns is None:
                        continue

                    # Extraer Folio y Serie, y construir Folio completo
                    folio = root_element.get('Folio')
                    serie = root_element.get('Serie')
                    folio_completo = f"{serie}-{folio}"

                    # Extraer Fecha del <cfdi:Comprobante>
                    fecha = root_element.get('Fecha')

                    # Extraer UUID desde <tfd:TimbreFiscalDigital>
                    uuid = None
                    complemento = root_element.find('./cfdi:Complemento', ns)
                    if complemento is not None:
                        timbre = complemento.find('./tfd:TimbreFiscalDigital', ns)
                        if timbre is not None:
                            uuid = timbre.get('UUID')

                    # Saltar si ya existe (por UUID o por Folio+Archivo)
                    if uuid:
                        if (df_database['UUID'] == uuid).any():
                            continue
                    else:
                        if ((df_database['Folio'] == folio_completo) & 
                            (df_database['Archivo'] == file)).any():
                            continue

                    # Extraer receptor
                    rec = root_element.find('./cfdi:Receptor', ns)
                    if rec is None:
                        continue
                    nombre = rec.get('Nombre')
                    rfc = rec.get('Rfc')

                    # Extraer cada concepto
                    for concepto in root_element.findall('./cfdi:Conceptos/cfdi:Concepto', ns):
                        descripcion = concepto.get('Descripcion')
                        cantidad    = concepto.get('Cantidad')
                        importe     = concepto.get('Importe')

                        data.append([
                            uuid,
                            folio_completo,
                            fecha,
                            nombre,
                            rfc,
                            descripcion,
                            cantidad,
                            importe,
                            file
                        ])

                except Exception as e:
                    print(f"[ERROR] Al procesar {file}: {e}")

    # Si hay nuevos registros, los agregamos y salvamos
    if data:
        df_nuevos = pd.DataFrame(data, columns=[
            'UUID', 'Folio', 'Fecha', 'Nombre', 'Rfc',
            'Descripcion', 'Cantidad', 'Importe', 'Archivo'
        ])
        df_database = pd.concat([df_database, df_nuevos], ignore_index=True)
        df_database[['Cantidad', 'Importe']] = df_database[['Cantidad', 'Importe']].astype(float)
        df_database.to_excel(xlsx_database, engine='openpyxl', index=False)
        print(f"\n✅ Se agregaron {len(df_nuevos)} nuevos registros a {xlsx_database}")
    else:
        print("\n✔️ No se encontraron nuevos XMLs para agregar.")


if __name__ == "__main__":
    invoice_paths = [
        r'C:\Users\arman\Dropbox\FACT 2023',
        r'C:\Users\arman\Dropbox\FACT 2024',
        r'C:\Users\arman\Dropbox\FACT 2025'
    ]
    folder_root = r"C:\Users\arman\Dropbox\3. Armando Cuaxospa\Adjudicaciones\Licitaciones 2025\E115 360"
    xlsx_database = os.path.join(folder_root, "Implementación", "Facturas", 'xmls_extraidos.xlsx')

    smart_xml_extraction(invoice_paths, xlsx_database)
