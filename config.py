# config.py
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
DATA_DIR = os.path.join(BASE_DIR, "data")

CARPETA_ADJUNTOS = os.path.join(DATA_DIR, "adjuntos")
CARPETA_EXTRAIDOS = os.path.join(DATA_DIR, "extraidos")

TEMP_CHECK_DIR = os.path.join(DATA_DIR, "temp_check")
TMP_DIR = TEMP_CHECK_DIR

ARCHIVO_EXCEL = os.path.join(DATA_DIR, "facturas.xlsx")
HISTORIAL_EXCEL = os.path.join(DATA_DIR, "historial_ejecuciones.xlsx")

STORE_NAME = "auxiliar.infraestructura@joyco.com.co"
APROB_FOLDER_NAME = "Facturas aprobadas"
APROB_SEARCH_SINCE_DAYS = 4

MATCH_PRIORIDAD = ["CUFE", "NUMERO_FECHA"]

APROB_CAT_OK = "AprobMatchOK"
APROB_CAT_ERROR = "AprobMatchError"

AUTO_STOP_MIN_PROCESADOS = 2
AUTO_STOP_SIN_MATCH_CONSEC = 2
AUTO_STOP_SIN_NUEVOS_CONSEC = 2

APROBACIONES_SP_RELATIVE_PATH = (
    "Innovacion/08. Pruebas proyectos/autoFacturas/Aprobaciones_Facturas.xlsx"
)
APROBACIONES_SHEET_NAME = "Hoja1"
APROB_COL_NUMERO = "NumeroFactura"
APROB_COL_RAD = "Radicado"
APROB_COL_PROY = "ProyectoProceso"

FACT_COL_NUMERO = "Número de factura"
FACT_COL_RAD = "Radicado"
FACT_COL_PROY = "ProyectoProceso"

# ==========================
# STATE (ProcessedStore)
# ==========================
STATE_DIR = os.path.join(DATA_DIR, "state")
PROCESSED_MESSAGES_PATH = os.path.join(STATE_DIR, "processed_messages.json")

# (opcional) TTL por defecto (días)
PROCESSED_MESSAGES_TTL_DAYS = 20
