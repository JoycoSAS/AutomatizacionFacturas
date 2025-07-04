import os
import sys

# Ruta base del proyecto (nivel raíz donde está main.py)
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

# Ruta a la carpeta de datos
DATA_DIR = os.path.join(BASE_DIR, "data")

# Carpetas internas
CARPETA_ADJUNTOS = os.path.join(DATA_DIR, "adjuntos")
CARPETA_EXTRAIDOS = os.path.join(DATA_DIR, "extraidos")
TEMP_CHECK_DIR = os.path.join(DATA_DIR, "temp_check")

# Archivos Excel
ARCHIVO_EXCEL = os.path.join(DATA_DIR, "facturas.xlsx")
HISTORIAL_EXCEL = os.path.join(DATA_DIR, "historial_ejecuciones.xlsx")

# Configuración Outlook
STORE_NAME = "auxiliar.infraestructura@joyco.com.co"
