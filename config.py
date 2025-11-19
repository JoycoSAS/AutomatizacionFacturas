import os
import sys

# Ruta base del proyecto (nivel raíz donde está main_*.py)
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

# Ruta a la carpeta de datos
DATA_DIR = os.path.join(BASE_DIR, "data")

# Carpetas internas
CARPETA_ADJUNTOS = os.path.join(DATA_DIR, "adjuntos")
CARPETA_EXTRAIDOS = os.path.join(DATA_DIR, "extraidos")

# --- TEMP: unificamos nombre como TMP_DIR (alias para compatibilidad) ---
TEMP_CHECK_DIR = os.path.join(DATA_DIR, "temp_check")
TMP_DIR = TEMP_CHECK_DIR  # ⬅️ alias usado por los controladores

# Archivos Excel
ARCHIVO_EXCEL = os.path.join(DATA_DIR, "facturas.xlsx")
HISTORIAL_EXCEL = os.path.join(DATA_DIR, "historial_ejecuciones.xlsx")

# Configuración Outlook
STORE_NAME = "auxiliar.infraestructura@joyco.com.co"

# === Disparo por carpeta de aprobados (Power Automate) ===
APROB_FOLDER_NAME = "Facturas aprobadas"   # nombre de la carpeta donde PA reenvía el PDF aprobado

# Ventana de búsqueda de correos con ZIP para intentar el match (en días hacia atrás)
APROB_SEARCH_SINCE_DAYS = 4

# Clave de matching (estricto por CUFE; si no hay CUFE en PDF, usa (Número + Fecha))
MATCH_PRIORIDAD = ["CUFE", "NUMERO_FECHA"]

# (Opcional) Categorías para marcar trazabilidad en el correo de aprobadas
APROB_CAT_OK    = "AprobMatchOK"
APROB_CAT_ERROR = "AprobMatchError"

# =====================================
# OPTIMIZACIÓN AUTOMÁTICA DE TIEMPO
# =====================================
# Mínimo de PDFs procesados antes de permitir cortes automáticos
AUTO_STOP_MIN_PROCESADOS = 2
# Cortar si hay esta cantidad de PDFs seguidos sin match
AUTO_STOP_SIN_MATCH_CONSEC = 2
# Cortar si hay esta cantidad de PDFs seguidos con match pero sin nuevos en Excel
AUTO_STOP_SIN_NUEVOS_CONSEC = 2

# =====================================
# CRUCE CON APROBACIONES (SharePoint)
# =====================================
# Ruta *en SharePoint* al Excel que rellena Power Automate (carpeta "excel")
# Ejemplo según tus capturas:
#   SOPORTES/Temporal Vehiculos/Prueba de Facturas Daniel/excel/Aprobaciones_Facturas.xlsx
APROBACIONES_SP_RELATIVE_PATH = "SOPORTES/Temporal Vehiculos/Prueba de Facturas Daniel/excel/Aprobaciones_Facturas.xlsx"

# Nombre de hoja donde PA escribe (si cambias el nombre, ajústalo aquí)
APROBACIONES_SHEET_NAME = "Hoja1"

# Nombres de columnas en el Excel de PA
APROB_COL_NUMERO = "NumeroFactura"
APROB_COL_RAD = "Radicado"
APROB_COL_PROY = "ProyectoProceso"

# Nombres de columnas en tu facturas.xlsx
FACT_COL_NUMERO = "Número de factura"   # si cambias el nombre en tu archivo, actualízalo aquí
FACT_COL_RAD = "Radicado"               # se crea si no existe
FACT_COL_PROY = "ProyectoProceso"       # se crea si no existe
