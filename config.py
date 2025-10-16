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

# === Disparo por carpeta de aprobados (Power Automate) ===
# Puedes usar el nombre visible de la carpeta dentro del buzón (debajo de Inbox) o el ID directo.
APROB_FOLDER_NAME = "Facturas aprobadas"   # nombre de la carpeta donde PA reenvía el PDF aprobado

# Ventana de búsqueda de correos con ZIP para intentar el match (en días hacia atrás)
APROB_SEARCH_SINCE_DAYS = 60

# Clave de matching (estricto por CUFE; si no hay CUFE en PDF, usa (Número + Fecha))
MATCH_PRIORIDAD = ["CUFE", "NUMERO_FECHA"]

# (Opcional) Categorías para marcar trazabilidad en el correo de aprobadas
APROB_CAT_OK    = "AprobMatchOK"
APROB_CAT_ERROR = "AprobMatchError"

