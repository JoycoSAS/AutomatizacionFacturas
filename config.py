import os
from pathlib import Path

# ============================================================
# Paths base
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
APROB_SEARCH_SINCE_DAYS = 90

MATCH_PRIORIDAD = ["CUFE", "NUMERO_FECHA"]

APROB_CAT_OK = "AprobMatchOK"
APROB_CAT_ERROR = "AprobMatchError"

AUTO_STOP_MIN_PROCESADOS = 9
AUTO_STOP_SIN_MATCH_CONSEC = 9
AUTO_STOP_SIN_NUEVOS_CONSEC = 9

# ============================================================
# FLUJO ESPECIAL DIAN
# ============================================================
APROB_DIAN_KEYWORD = "DIAN"

# Más amplio y tolerante
INBOX_DIAN_SUBJECT_CANDIDATES = [
    "DIAN",
    "VALIDACION DIAN",
    "VALIDACIONES DIAN",
    "VALIDACION",
    "VALIDACIONES",
    "02-VALIDACION DIAN",
    "02 VALIDACION DIAN",
    "02-VALIDACIONES DIAN",
    "02 VALIDACIONES DIAN",
    "VALIDACION DIAN JOYCO",
    "VALIDACIONES DIAN JOYCO",
    "DIAN VIAL",
    "VALIDACION JOYCO",
    "VALIDACIONES JOYCO",
    "VALIDACION JOYCO SAS",
    "VALIDACIONES JOYCO SAS",
]

# True = si hay preview, también valida ahí; si no hay, cae a asunto
REQUIRE_DIAN_IN_BODY_PREVIEW = True

# ============================================================
# RADICADOS
# ============================================================
RADICADOS_SP_RELATIVE_PATH = (
    "Control de correspondencia Oficina Principal/"
    "Control correspondencia Oficina Principal.xlsx"
)

RADICADOS_SHEET_NAME = "Recibida"

RAD_COL_ASUNTO = "Asunto"
RAD_COL_RADICADO = "Consecutivo de entrada"
RAD_COL_PROY = "Proyecto o Proceso"

FACT_COL_NUMERO = "Número de factura"
FACT_COL_RAD = "Radicado"
FACT_COL_PROY = "ProyectoProceso"

# ============================================================
# RADICADOS local
# ============================================================
RADICADOS_LOCAL_DIR = os.path.join(TEMP_CHECK_DIR, "radicados")
RADICADOS_LOCAL_FILENAME = "Control correspondencia Oficina Principal.xlsx"
RADICADOS_LOCAL_PATH = os.path.join(RADICADOS_LOCAL_DIR, RADICADOS_LOCAL_FILENAME)

# ============================================================
# STATE
# ============================================================
STATE_DIR = os.path.join(DATA_DIR, "state")
PROCESSED_MESSAGES_PATH = os.path.join(STATE_DIR, "processed_messages.json")
PROCESSED_MESSAGES_TTL_DAYS = 20000

# ============================================================
# AttachmentIndexStore
# ============================================================
ATTACHMENT_INDEX_PATH = os.path.join(STATE_DIR, "attachment_index_store.json")
ATTACHMENT_INDEX_TTL_DAYS = 965

# ============================================================
# Auditoría CSV
# ============================================================
AUDIT_DIR = os.path.join(DATA_DIR, "audit")
AUDIT_RUNS_PREFIX = "audit_runs"
AUDIT_DETALLE_PREFIX = "audit_detalle"
AUDIT_WRITE_ONLY_IF_ACTIVITY = True

# ============================================================
# Lock
# ============================================================
LOCK_FILE_APROBADAS = os.path.join(STATE_DIR, "aprobadas.lock")

# 4 horas
LOCK_TTL_SECONDS = 240 * 60