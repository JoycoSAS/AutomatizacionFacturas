import os

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass


# ============================================================
# HELPERS DE CONFIGURACIÓN
# ============================================================

def _env_str(name: str, default: str = "") -> str:
    value = os.getenv(name)
    if value is None:
        return default

    value = str(value).strip()
    return value if value else default


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name)

    if raw is None or str(raw).strip() == "":
        return int(default)

    try:
        return int(str(raw).strip())
    except Exception:
        print(f"⚠️ config.py: variable {name} inválida={raw!r}. Uso default={default}.")
        return int(default)


def _env_bool(name: str, default: bool = False) -> bool:
    raw = os.getenv(name)

    if raw is None or str(raw).strip() == "":
        return bool(default)

    value = str(raw).strip().lower()

    if value in {"1", "true", "yes", "y", "si", "sí", "on"}:
        return True

    if value in {"0", "false", "no", "n", "off"}:
        return False

    print(f"⚠️ config.py: variable {name} inválida={raw!r}. Uso default={default}.")
    return bool(default)


def _path_from_base_or_absolute(path_value: str) -> str:
    """
    Permite usar rutas relativas o absolutas desde .env.

    Ejemplos válidos:
    - data/facturas.xlsx
    - data/prod/facturas.xlsx
    - C:/ruta/local/facturas.xlsx
    - /opt/joyco/facturas_procesador/data/prod/facturas.xlsx
    """
    if os.path.isabs(path_value):
        return path_value

    return os.path.join(BASE_DIR, path_value)


# ============================================================
# Paths base
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

CARPETA_ADJUNTOS = os.path.join(DATA_DIR, "adjuntos")
CARPETA_EXTRAIDOS = os.path.join(DATA_DIR, "extraidos")

TEMP_CHECK_DIR = os.path.join(DATA_DIR, "temp_check")
TMP_DIR = TEMP_CHECK_DIR


# ============================================================
# Excel local runtime
# ============================================================
# Este Excel local es el archivo de trabajo interno del automatizador.
#
# El Excel oficial visible para usuarios debe ser el Excel Web actualizado
# por Workbook API.
#
# Por ahora queda:
#   data/facturas.xlsx
#
# Más adelante en VPS se podrá usar, por ejemplo:
#   ARCHIVO_EXCEL_LOCAL=data/prod/facturas.xlsx
#   HISTORIAL_EXCEL_LOCAL=data/prod/historial_ejecuciones.xlsx
# ============================================================

ARCHIVO_EXCEL = _path_from_base_or_absolute(
    _env_str("ARCHIVO_EXCEL_LOCAL", os.path.join("data", "facturas.xlsx"))
)

HISTORIAL_EXCEL = _path_from_base_or_absolute(
    _env_str("HISTORIAL_EXCEL_LOCAL", os.path.join("data", "historial_ejecuciones.xlsx"))
)


# ============================================================
# Correo / carpeta aprobadas
# ============================================================
# STORE_NAME:
#   Buzón que consulta el sistema.
#
# APROB_FOLDER_NAME:
#   Carpeta donde llegan las facturas aprobadas.
#
# APROB_SEARCH_SINCE_DAYS:
#   Ventana de días para búsqueda en aprobadas.
#   Por defecto se sincroniza con FACTURAS_SINCE_DAYS.
#
# Producción actual:
#   FACTURAS_SINCE_DAYS=6
# ============================================================

STORE_NAME = _env_str("MAILBOX_UPN", "radicacion@joyco.com.co")

APROB_FOLDER_NAME = _env_str("APROB_FOLDER_NAME", "solo aprobadas")

APROB_SEARCH_SINCE_DAYS = _env_int(
    "APROB_SEARCH_SINCE_DAYS",
    _env_int("FACTURAS_SINCE_DAYS", 6),
)

# ============================================================
# Modo histórico / reproceso controlado
# ============================================================
# Estas variables NO cambian el comportamiento de producción normal.
#
# Uso esperado para corrida histórica:
#   FACTURAS_MODO=HISTORICO
#   FACTURAS_ORDEN_HISTORICO=ASC
#   FACTURAS_PROCESAR_ANTIGUOS_PRIMERO=1
#   FACTURAS_HISTORICO_DESDE=2026-04-01
#   FACTURAS_HISTORICO_HASTA=2026-06-09
#   FACTURAS_DISABLE_AUTOSTOP=1
#   FACTURAS_HISTORICO_DRY_RUN=1
#
# FACTURAS_HISTORICO_DRY_RUN:
#   True  = prueba segura del orden, sin escribir Excel ni procesar adjuntos.
#   False = corrida real. Usar solo cuando ya validemos el orden.
# ============================================================

FACTURAS_MODO = _env_str("FACTURAS_MODO", "PRODUCCION").upper()
if FACTURAS_MODO not in {"HISTORICO", "DIARIO", "PRODUCCION", "PRODUCCIÓN"}:
    print(f"⚠️ config.py: FACTURAS_MODO inválido={FACTURAS_MODO!r}. Uso PRODUCCION.")
    FACTURAS_MODO = "PRODUCCION"

FACTURAS_ORDEN_HISTORICO = _env_str("FACTURAS_ORDEN_HISTORICO", "").upper()
if FACTURAS_ORDEN_HISTORICO not in {"", "ASC", "DESC"}:
    print(
        "⚠️ config.py: FACTURAS_ORDEN_HISTORICO inválido="
        f"{FACTURAS_ORDEN_HISTORICO!r}. Uso vacío."
    )
    FACTURAS_ORDEN_HISTORICO = ""

FACTURAS_PROCESAR_ANTIGUOS_PRIMERO = _env_bool(
    "FACTURAS_PROCESAR_ANTIGUOS_PRIMERO",
    FACTURAS_ORDEN_HISTORICO == "ASC",
)

FACTURAS_HISTORICO_DESDE = _env_str("FACTURAS_HISTORICO_DESDE", "")
FACTURAS_HISTORICO_HASTA = _env_str("FACTURAS_HISTORICO_HASTA", "")

FACTURAS_DISABLE_AUTOSTOP = _env_bool("FACTURAS_DISABLE_AUTOSTOP", False)
FACTURAS_HISTORICO_DRY_RUN = _env_bool("FACTURAS_HISTORICO_DRY_RUN", False)

FACTURAS_HISTORICO_ASC_ACTIVO = (
    FACTURAS_MODO == "HISTORICO"
    and (
        FACTURAS_ORDEN_HISTORICO == "ASC"
        or FACTURAS_PROCESAR_ANTIGUOS_PRIMERO
    )
)


# ============================================================
# Match
# ============================================================
# MATCH_PRIORIDAD:
#   Orden de búsqueda para relacionar PDF aprobado con XML/ZIP.
#
# No tocar por rendimiento sin revisar la lógica del controller.
# ============================================================

MATCH_PRIORIDAD = ["CUFE", "NUMERO_FECHA"]

APROB_CAT_OK = _env_str("APROB_CAT_OK", "AprobMatchOK")
APROB_CAT_ERROR = _env_str("APROB_CAT_ERROR", "AprobMatchError")


# ============================================================
# Auto stop / corte temprano
# ============================================================
# Estas variables ayudan a que producción no se quede recorriendo
# demasiados correos cuando ya encuentra registros repetidos.
#
# AUTO_STOP_MIN_PROCESADOS:
#   Mínimo de facturas procesadas antes de permitir el corte.
#
# AUTO_STOP_SIN_NUEVOS_CONSEC:
#   Facturas seguidas con match pero sin filas nuevas.
#   Normalmente significa que ya estaban registradas.
#
# AUTO_STOP_SIN_MATCH_CONSEC:
#   Facturas seguidas sin match antes de cortar.
#   Se deja un poco más alto para no cortar por casos difíciles.
#
# Recomendado producción:
#   AUTO_STOP_MIN_PROCESADOS=3
#   AUTO_STOP_SIN_NUEVOS_CONSEC=3
#   AUTO_STOP_SIN_MATCH_CONSEC=5
# ============================================================

if FACTURAS_DISABLE_AUTOSTOP:
    # En histórico no conviene cortar por repetidos/sin match,
    # porque se busca recorrer una ventana grande y cronológica.
    AUTO_STOP_MIN_PROCESADOS = 999999999
    AUTO_STOP_SIN_NUEVOS_CONSEC = 999999999
    AUTO_STOP_SIN_MATCH_CONSEC = 999999999
else:
    AUTO_STOP_MIN_PROCESADOS = _env_int("AUTO_STOP_MIN_PROCESADOS", 3)
    AUTO_STOP_SIN_NUEVOS_CONSEC = _env_int("AUTO_STOP_SIN_NUEVOS_CONSEC", 3)
    AUTO_STOP_SIN_MATCH_CONSEC = _env_int("AUTO_STOP_SIN_MATCH_CONSEC", 5)


# ============================================================
# FLUJO ESPECIAL DIAN
# ============================================================

APROB_DIAN_KEYWORD = _env_str("APROB_DIAN_KEYWORD", "DIAN")

# Más amplio y tolerante.
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

# True:
#   Si hay bodyPreview, también valida ahí.
#   Si no hay, cae al asunto.
REQUIRE_DIAN_IN_BODY_PREVIEW = _env_bool("REQUIRE_DIAN_IN_BODY_PREVIEW", True)


# ============================================================
# RADICADOS
# ============================================================

RADICADOS_SP_RELATIVE_PATH = _env_str(
    "RADICADOS_SP_RELATIVE_PATH",
    (
        "Control de correspondencia Oficina Principal/"
        "Control correspondencia Oficina Principal.xlsx"
    ),
)

RADICADOS_SHEET_NAME = _env_str("RADICADOS_SHEET_NAME", "Recibida")

RAD_COL_ASUNTO = _env_str("RAD_COL_ASUNTO", "Asunto")
RAD_COL_RADICADO = _env_str("RAD_COL_RADICADO", "Consecutivo de entrada")
RAD_COL_PROY = _env_str("RAD_COL_PROY", "Proyecto o Proceso")

FACT_COL_NUMERO = _env_str("FACT_COL_NUMERO", "Número de factura")
FACT_COL_RAD = _env_str("FACT_COL_RAD", "Radicado")
FACT_COL_PROY = _env_str("FACT_COL_PROY", "ProyectoProceso")


# ============================================================
# RADICADOS local
# ============================================================

RADICADOS_LOCAL_DIR = os.path.join(TEMP_CHECK_DIR, "radicados")

RADICADOS_LOCAL_FILENAME = _env_str(
    "RADICADOS_LOCAL_FILENAME",
    "Control correspondencia Oficina Principal.xlsx",
)

RADICADOS_LOCAL_PATH = os.path.join(RADICADOS_LOCAL_DIR, RADICADOS_LOCAL_FILENAME)


# ============================================================
# STATE
# ============================================================

STATE_DIR = os.path.join(DATA_DIR, "state")

PROCESSED_MESSAGES_PATH = os.path.join(
    STATE_DIR,
    _env_str("PROCESSED_MESSAGES_FILENAME", "processed_messages.json"),
)

# TTL del ProcessedStore.
#
# Producción recomendado:
#   365 días.
#
# Esto evita que processed_messages.json crezca sin control durante años.
PROCESSED_MESSAGES_TTL_DAYS = _env_int("PROCESSED_MESSAGES_TTL_DAYS", 365)


# ============================================================
# AttachmentIndexStore / AIDX
# ============================================================

ATTACHMENT_INDEX_PATH = os.path.join(
    STATE_DIR,
    _env_str("ATTACHMENT_INDEX_FILENAME", "attachment_index_store.json"),
)

# TTL del índice de adjuntos/ZIP/XML.
#
# Producción recomendado:
#   365 días.
#
# Si se necesita más histórico:
#   730 días.
#
# No dejar infinito si el sistema va a correr por años.
ATTACHMENT_INDEX_TTL_DAYS = _env_int("ATTACHMENT_INDEX_TTL_DAYS", 365)


# ============================================================
# Auditoría CSV
# ============================================================

AUDIT_DIR = os.path.join(DATA_DIR, "audit")

AUDIT_RUNS_PREFIX = _env_str("AUDIT_RUNS_PREFIX", "audit_runs")
AUDIT_DETALLE_PREFIX = _env_str("AUDIT_DETALLE_PREFIX", "audit_detalle")

# True:
#   Solo escribe audit si hubo actividad.
#
# False:
#   Escribe audit aunque no haya actividad.
#   Puede servir más adelante para evidencia de cada corrida.
AUDIT_WRITE_ONLY_IF_ACTIVITY = _env_bool("AUDIT_WRITE_ONLY_IF_ACTIVITY", True)


# ============================================================
# Lock
# ============================================================

LOCK_FILE_APROBADAS = os.path.join(STATE_DIR, "aprobadas.lock")

# TTL del lock.
#
# Sirve para evitar que dos ejecuciones corran al mismo tiempo.
#
# Si una ejecución se queda pegada o el proceso muere, este TTL permite
# considerar vencido el lock después de cierto tiempo.
#
# Producción:
#   3600 = 1 hora.
#
# Reproceso histórico largo:
#   14400 = 4 horas.
#
# No poner demasiado bajo si una ejecución puede durar más que ese tiempo,
# porque una segunda ejecución podría creer que el lock ya venció.
LOCK_TTL_SECONDS = _env_int(
    "LOCK_TTL_SECONDS",
    14400 if FACTURAS_MODO == "HISTORICO" else 3600,
)


# ============================================================
# Compatibilidad opcional con variable antigua
# ============================================================
# Esta variable existía en algunas pruebas antiguas.
# Se deja disponible por si algún módulo viejo la consulta.
# El control oficial actual debe ser AUTO_STOP_*.
# ============================================================

STOP_AFTER_ALREADY_PROCESSED = _env_int("STOP_AFTER_ALREADY_PROCESSED", 3)


# ============================================================
# Crear carpetas base si no existen
# ============================================================

for _folder in [
    DATA_DIR,
    CARPETA_ADJUNTOS,
    CARPETA_EXTRAIDOS,
    TEMP_CHECK_DIR,
    STATE_DIR,
    AUDIT_DIR,
]:
    try:
        os.makedirs(_folder, exist_ok=True)
    except Exception:
        pass


# ============================================================
# Sello de versión
# ============================================================

print("🔥 CONFIG VERSION ACTIVA: 2026-06-09-HISTORICO-ASC-DRYRUN-AUTOSTOP")