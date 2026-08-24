# XML to Excel Automation Pipeline

Script en Python diseñado para automatizar la extracción, parsing y estructuración de datos desde archivos XML (facturas/comprobantes electrónicos) hacia hojas de cálculo de Microsoft Excel (`.xlsx`). 

Este desarrollo elimina la carga de trabajo manual y reduce el error humano en el procesamiento de comprobantes fiscales y financieros.

## 🚀 Características
* **Parsing de XML:** Lectura e interpretación de nodos XML (CFDI / comprobantes).
* **Limpieza y Estructuración:** Extracción de campos clave (Folio, Fecha, Emisor, Receptor, Importes, Impuestos).
* **Exportación a Excel:** Generación o actualización automática de libros Excel con formato estructurado.
* **Manejo de Errores:** Validación de archivos corruptos o estructuras XML incompletas.

## 🛠️ Stack Técnico
* **Lenguaje:** Python 3.x
* **Librerías:** `openpyxl`, `xml.etree.ElementTree`

## 📋 Requisitos Previos
Tener instalado Python 3.8+ en el sistema.

```bash
pip install pandas openpyxl


⚙️ Uso e Instalación
Clonar el repositorio: 

Bash
git clone [https://github.com/rbarrientos182/scritpxml.git]
cd scritpxml

Colocar los archivos XML:
Ubica los archivos .xml a procesar en la carpeta /data o ruta configurada.

Ejecutar el script:
Bash
python main.py

📄 Licencia
Este proyecto es de uso personal y educativo licencia MIT.

***

**Detalles a revisar antes de publicar:**
* Asegúrate de agregar un archivo `.gitignore` para omitir carpetas virtuales (`venv/`), archivos de caché (`__pycache__/`) y cualquier archivo `.xml` o `.xlsx` real con datos sensibles.
* Si el script procesa carpetas completas con múltiples XML, destacarlo en la sección de características como *procesamiento por lotes (batch processing)* agrega valor técnico.