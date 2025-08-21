# Proyecto: Adjudicaciones y Licitaciones 2025 - E115 360

Este repositorio contiene un conjunto de herramientas y scripts diseñados para automatizar y gestionar procesos relacionados con adjudicaciones y licitaciones. El núcleo del proyecto está compuesto por varios Jupyter Notebooks que orquestan tareas específicas, desde la descarga de datos hasta la generación de reportes.

## Estructura del Proyecto

### Jupyter Notebooks Principales

1. **00 Contratos.ipynb**  
   - Captura y gestión de contratos.
   - Generación de bases de datos estandarizadas con información de instituciones, piezas, procedimientos y precios.

2. **01 Downloader.ipynb**  
   - Orquestador para la descarga de datos desde sistemas como SAI, CAMUNDA, SAGI, ZOHO y PISP.
   - Incluye funciones para fusionar y auditar archivos descargados.

3. **02 Sanciones IMSSB.ipynb**  
   - Procesamiento de sanciones relacionadas con el IMSS Bienestar.
   - Generación de auditorías y entregables.

4. **03 Generador de cartas poder.ipynb**  
   - Automatización para la generación de cartas poder.
   - Permite elegir entre cartas con o sin testigos.

5. **04 Facturas.ipynb**  
   - Validación y procesamiento de facturas.
   - Asegura la consistencia de referencias, importes y fechas.

6. **05 Cobranza.ipynb**  
   - Gestión de cobranza y atención a proveedores.
   - Organización de datos relacionados con IMSS-Bienestar.

7. **06 Integración ejecutiva.ipynb**  
   - Integración de información de altas, facturas y PREI.
   - Generación de reportes ejecutivos.

### Carpetas y Archivos Adicionales

- **Implementación/**  
  Contiene scripts y datos auxiliares utilizados por los notebooks.

- **Contratos temp/**  
  Carpeta temporal para la gestión de contratos.

- **Scripts/**  
  Librerías internas organizadas por funcionalidad (e.g., `Libreria_contratos`, `Libreria_facturas`).

- **Archivos de configuración**  
  - `.gitignore`: Define los archivos y carpetas que no se deben incluir en el control de versiones.

## Requisitos

- **Python**: Asegúrate de tener instalada la versión 3.13 o superior.
- **Dependencias**: Las librerías necesarias están incluidas en los scripts y deben instalarse antes de ejecutar los notebooks.

## Uso

1. Clona este repositorio:
   ```bash
   git clone <URL_DEL_REPOSITORIO>