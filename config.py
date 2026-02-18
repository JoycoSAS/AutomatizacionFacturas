# config.py
import os
from pathlib import Path

# ============================================================
# Paths base (más estable que sys.argv[0] para tests)
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

CARPETA_ADJUNTOS = os.path.join(DATA_DIR, "adjuntos")
CARPETA_EXTRAIDOS = os.path.join(DATA_DIR, "extraidos")

TEMP_CHECK_DIR = os.path.join(DATA_DIR, "temp_check")
TMP_DIR = TEMP_CHECK_DIR

ARCHIVO_EXCEL = os.path.join(DATA_DIR, "facturas.xlsx")
HISTORIAL_EXCEL = os.path.join(DATA_DIR, "historial_ejecuciones.xlsx")

# ============================================================
# Correo / carpeta aprobadas
# ============================================================
STORE_NAME = "radicacion@joyco.com.co"

APROB_FOLDER_NAME = "solo aprobadas"

# ⚠️ Para “cargar histórico” puedes dejarlo alto (ej: 365).
# Luego bájalo a 4 cuando ya estés en operación diaria.
APROB_SEARCH_SINCE_DAYS = 365

MATCH_PRIORIDAD = ["CUFE", "NUMERO_FECHA"]

APROB_CAT_OK = "AprobMatchOK"
APROB_CAT_ERROR = "AprobMatchError"

AUTO_STOP_MIN_PROCESADOS = 2
AUTO_STOP_SIN_MATCH_CONSEC = 2
AUTO_STOP_SIN_NUEVOS_CONSEC = 2

# ============================================================
# ✅ RADICADOS: Control Correspondencia (SharePoint / Drive nuevo)
# ============================================================
# Este es el archivo real en SharePoint (ruta relativa dentro del drive)
# OJO: según tu captura, el root del drive YA contiene la carpeta:
# "Control de correspondencia Oficina Principal"
RADICADOS_SP_RELATIVE_PATH = (
    "Control de correspondencia Oficina Principal/"
    "Control correspondencia Oficina Principal.xlsx"
)

# Hoja a leer dentro del Excel de radicados
RADICADOS_SHEET_NAME = "Recibida"

# Columnas a leer del Excel de radicados
RAD_COL_ASUNTO = "Asunto"
RAD_COL_RADICADO = "Consecutivo de entrada"
RAD_COL_PROY = "Proyecto o Proceso"

# Columnas destino dentro de facturas.xlsx
FACT_COL_NUMERO = "Número de factura"
FACT_COL_RAD = "Radicado"
FACT_COL_PROY = "ProyectoProceso"

# ============================================================
# ✅ RADICADOS: ruta local estándar (para tests y ejecución)
# ============================================================
# El downloader debe guardar SIEMPRE aquí para que:
# - tests/sp_debug_radicados_headers.py lo encuentre
# - el pipeline lo lea sin depender de rutas viejas
RADICADOS_LOCAL_DIR = os.path.join(TEMP_CHECK_DIR, "radicados")
RADICADOS_LOCAL_FILENAME = "Control correspondencia Oficina Principal.xlsx"
RADICADOS_LOCAL_PATH = os.path.join(RADICADOS_LOCAL_DIR, RADICADOS_LOCAL_FILENAME)

# (Opcional) asegurar que existan carpetas base (solo constantes acá; creación real en runtime)
# Path(RADICADOS_LOCAL_DIR).mkdir(parents=True, exist_ok=True)

# ============================================================
# STATE (ProcessedStore)
# ============================================================
STATE_DIR = os.path.join(DATA_DIR, "state")
PROCESSED_MESSAGES_PATH = os.path.join(STATE_DIR, "processed_messages.json")
PROCESSED_MESSAGES_TTL_DAYS = 2000

# ============================================================
# AttachmentIndexStore (ZIP históricos)
# ============================================================
ATTACHMENT_INDEX_PATH = os.path.join(STATE_DIR, "attachment_index_store.json")
ATTACHMENT_INDEX_TTL_DAYS = 765
