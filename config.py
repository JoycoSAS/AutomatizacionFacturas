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

AUTO_STOP_MIN_PROCESADOS = 9999
AUTO_STOP_SIN_MATCH_CONSEC = 9999
AUTO_STOP_SIN_NUEVOS_CONSEC = 9999

# ============================================================
# ✅ FLUJO ESPECIAL "DIAN" (PDF-only desde "Validación(s) DIAN")
# ============================================================
# Palabra clave que activa el flujo especial dentro de "solo aprobadas"
APROB_DIAN_KEYWORD = "DIAN"

# Asuntos aceptados del correo contenedor en INBOX donde están "todos los PDFs"
# (se normaliza: sin tildes, sin dobles espacios, etc.)
INBOX_DIAN_SUBJECT_CANDIDATES = [
    "VALIDACION DIAN",
    "VALIDACIONES DIAN",
    "VALIDACIÓN DIAN",
    "VALIDACIONES DIAN",
]

# Si el mensaje NO trae bodyPreview desde Graph, se hace fallback a solo asunto.
# True = exigir DIAN también en bodyPreview si existe.
REQUIRE_DIAN_IN_BODY_PREVIEW = True

# ============================================================
# ✅ RADICADOS: Control Correspondencia (SharePoint / Drive nuevo)
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
# ✅ RADICADOS: ruta local estándar (para tests y ejecución)
# ============================================================
RADICADOS_LOCAL_DIR = os.path.join(TEMP_CHECK_DIR, "radicados")
RADICADOS_LOCAL_FILENAME = "Control correspondencia Oficina Principal.xlsx"
RADICADOS_LOCAL_PATH = os.path.join(RADICADOS_LOCAL_DIR, RADICADOS_LOCAL_FILENAME)

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

# ============================================================
# ✅ AUDITORÍA CSV (rotación diaria)
# ============================================================
AUDIT_DIR = os.path.join(DATA_DIR, "audit")
AUDIT_RUNS_PREFIX = "audit_runs"         # audit_runs_YYYY-MM-DD.csv
AUDIT_DETALLE_PREFIX = "audit_detalle"   # audit_detalle_YYYY-MM-DD.csv

# True = no escribir CSV si no hubo nada que procesar (modo recomendado para "cada minuto")
AUDIT_WRITE_ONLY_IF_ACTIVITY = True

# ============================================================
# ✅ LOCK (anti-ejecución simultánea)
# ============================================================
LOCK_FILE_APROBADAS = os.path.join(STATE_DIR, "aprobadas.lock")

# TTL del lock: si por alguna razón un proceso muere y deja el lock,
# pasados X segundos se considera stale y se reemplaza.
LOCK_TTL_SECONDS = 60 * 30  # 30 minutos