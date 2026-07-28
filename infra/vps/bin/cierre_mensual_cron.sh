#!/usr/bin/env bash

set -u
set -o pipefail
umask 027

export HOME="/home/innovacion"
export PATH="/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin"
export LANG="C.UTF-8"
export LC_ALL="C.UTF-8"
export TZ="America/Bogota"
export PYTHONUNBUFFERED="1"

VERSION_WRAPPER="2026-07-28-CIERRE-MENSUAL-CRON-V1"

APP_DIR="/opt/joyco/facturas-procesador/app"
PYTHON_BIN="/opt/joyco/facturas-procesador/venv/bin/python"
ENV_FILE="/etc/joyco/facturas-procesador/facturas.env"

SCRIPT_CIERRE="$APP_DIR/scripts/cierre_mensual_facturas.py"
SCRIPT_SUBIDA="$APP_DIR/scripts/subir_cierre_mensual_sharepoint.py"

CIERRE_ROOT="/var/lib/joyco/facturas-procesador/cierres/diarios"

LOG_DIR="/var/log/joyco/facturas-procesador/cron"
LOCK_DIR="/var/lib/joyco/facturas-procesador/state/locks"

LOCK_PRINCIPAL="$LOCK_DIR/cron_facturas_principal.lock"
LOCK_MENSUAL="$LOCK_DIR/cron_cierre_mensual.lock"
LOCK_APP="/var/lib/joyco/facturas-procesador/state/aprobadas.lock"

ESPERA_LOCK_PRINCIPAL_SEGUNDOS=21600

MODE="${1:-run}"
FECHA_EJECUCION="$(date +%F)"

if ! mkdir -p \
  "$LOG_DIR" \
  "$LOCK_DIR" \
  "$CIERRE_ROOT"; then

  echo "ERROR: no fue posible preparar directorios operativos."
  exit 70
fi

calcular_mes_anterior() {
  "$PYTHON_BIN" - <<'PY'
from datetime import date, timedelta


MESES_ES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

hoy = date.today()
fin = hoy.replace(day=1) - timedelta(days=1)
inicio = fin.replace(day=1)

periodo = inicio.strftime("%Y-%m")
anio = inicio.strftime("%Y")
mes_directorio = f"{periodo}_{MESES_ES[inicio.month]}"

print(
    inicio.isoformat(),
    fin.isoformat(),
    periodo,
    anio,
    mes_directorio,
)
PY
}

if ! read -r INICIO FIN PERIODO ANIO MES_DIRECTORIO < <(
  calcular_mes_anterior
); then
  echo "ERROR: no fue posible calcular el mes anterior."
  exit 70
fi

CIERRE_DIR="$CIERRE_ROOT/$ANIO/$MES_DIRECTORIO/Mensual"
VALIDACION_LOCAL="$CIERRE_DIR/05_Validaciones/validacion_local_mensual_${PERIODO}.json"
VALIDACION_REMOTA="$CIERRE_DIR/05_Validaciones/validacion_remota_mensual_${PERIODO}.json"
ESTADO_RETENCION="$CIERRE_DIR/05_Validaciones/estado_retencion_mensual_${PERIODO}.json"

SELLO_LOG="$(date +%Y%m%d_%H%M%S)"
LOG_FILE="$LOG_DIR/cierre_mensual_${PERIODO}_${SELLO_LOG}.log"

check_environment() {
  local requerido

  for requerido in \
    "$APP_DIR" \
    "$ENV_FILE" \
    "$SCRIPT_CIERRE" \
    "$SCRIPT_SUBIDA" \
    "$CIERRE_ROOT" \
    "/usr/bin/flock"; do

    if [[ ! -e "$requerido" ]]; then
      echo "ERROR: falta $requerido"
      return 1
    fi
  done

  if [[ ! -x "$PYTHON_BIN" ]]; then
    echo "ERROR: Python no es ejecutable: $PYTHON_BIN"
    return 1
  fi

  if ! "$PYTHON_BIN" -m py_compile \
    "$SCRIPT_CIERRE" \
    "$SCRIPT_SUBIDA"; then

    echo "ERROR: uno de los scripts mensuales no compila."
    return 1
  fi

  return 0
}

if [[ "$MODE" == "--check" ]]; then
  echo "============================================================"
  echo "COMPROBACIÓN DEL WRAPPER MENSUAL"
  echo "============================================================"
  echo "Versión: $VERSION_WRAPPER"

  if ! check_environment; then
    echo "ERROR: comprobación del entorno fallida."
    exit 70
  fi

  echo "Fecha de comprobación: $FECHA_EJECUCION"
  echo "Último mes completo calculado: $INICIO a $FIN"
  echo "Periodo: $PERIODO"
  echo "Carpeta esperada: $CIERRE_DIR"
  echo "Python: $PYTHON_BIN"
  echo "Generador: $SCRIPT_CIERRE"
  echo "Uploader: $SCRIPT_SUBIDA"
  echo "Espera máxima para obtener el bloqueo: $ESPERA_LOCK_PRINCIPAL_SEGUNDOS segundos"
  echo "Log previsto: $LOG_FILE"
  echo "OK: comprobación terminada sin generar ni subir archivos."
  echo "============================================================"

  exit 0
fi

if [[ "$MODE" != "run" ]]; then
  echo "Uso: $0 [run|--check]"
  exit 64
fi

{
  echo
  echo "===================================================================================================="
  echo "CRON_CIERRE_MENSUAL_INICIO|$(date --iso-8601=seconds)|periodo=$PERIODO|inicio=$INICIO|fin=$FIN|pid=$$"
  echo "VERSION_WRAPPER|$VERSION_WRAPPER"
  echo "===================================================================================================="

  if ! check_environment; then
    echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=70|paso=entorno"
    exit 70
  fi

  if ! cd "$APP_DIR"; then
    echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=71|paso=directorio"
    exit 71
  fi

  exec 9>"$LOCK_MENSUAL"

  if ! /usr/bin/flock -n 9; then
    echo "CRON_CIERRE_MENSUAL_SKIP|$(date --iso-8601=seconds)|motivo=mensual_activo"
    exit 0
  fi

  echo "Esperando acceso exclusivo al flujo principal y demás cierres..."

  exec 8>"$LOCK_PRINCIPAL"

  if ! /usr/bin/flock -w "$ESPERA_LOCK_PRINCIPAL_SEGUNDOS" 8; then
    echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=75|paso=espera_principal"
    exit 75
  fi

  echo "OK: acceso exclusivo obtenido."

  if [[ -e "$LOCK_APP" ]]; then
    echo "ERROR: existe aprobadas.lock después de obtener el bloqueo principal."
    echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=76|paso=lock_aplicacion"
    exit 76
  fi

  if [[ -d "$CIERRE_DIR" ]]; then
    echo
    echo "CIERRE_MENSUAL_EXISTENTE|periodo=$PERIODO|ruta=$CIERRE_DIR"
    echo "No se reemplazará ni modificará el paquete local existente."
  else
    echo
    echo "GENERANDO CIERRE MENSUAL LOCAL..."

    "$PYTHON_BIN" - \
      "$SCRIPT_CIERRE" \
      "$ENV_FILE" \
      "$INICIO" \
      "$FIN" <<'PY'
import os
import runpy
import sys
from pathlib import Path

from dotenv import load_dotenv


script = Path(sys.argv[1])
env_file = Path(sys.argv[2])
inicio = sys.argv[3]
fin = sys.argv[4]

os.environ["FACTURAS_ENV_FILE"] = str(env_file)

if not load_dotenv(
    dotenv_path=env_file,
    override=True,
    encoding="utf-8",
):
    raise RuntimeError(
        "No fue posible cargar facturas.env."
    )

sys.argv = [
    str(script),
    "--inicio",
    inicio,
    "--fin",
    fin,
    "--permitir-vacio",
]

runpy.run_path(
    str(script),
    run_name="__main__",
)
PY

    cierre_rc=$?

    echo "CRON_CIERRE_MENSUAL_LOCAL_FIN|$(date --iso-8601=seconds)|rc=$cierre_rc"

    if [[ "$cierre_rc" -ne 0 ]]; then
      echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=$cierre_rc|paso=cierre_local"
      exit "$cierre_rc"
    fi
  fi

  if [[ ! -d "$CIERRE_DIR" ]]; then
    echo "ERROR: no existe la carpeta mensual esperada: $CIERRE_DIR"
    echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=77|paso=carpeta_local"
    exit 77
  fi

  if [[ ! -s "$VALIDACION_LOCAL" ]]; then
    echo "ERROR: falta la validación local mensual."
    echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=78|paso=validacion_local"
    exit 78
  fi

  "$PYTHON_BIN" - "$VALIDACION_LOCAL" "$PERIODO" <<'PY'
import json
import sys
from pathlib import Path


path = Path(sys.argv[1])
periodo = sys.argv[2]

payload = json.loads(
    path.read_text(
        encoding="utf-8"
    )
)

if payload.get("ok") is not True:
    raise RuntimeError(
        "La validación local mensual no está en estado OK."
    )

if payload.get("periodo") != periodo:
    raise RuntimeError(
        "El periodo de la validación local mensual no coincide."
    )

print(
    "VALIDACION_LOCAL_MENSUAL_OK"
    f"|archivo={path.name}"
    f"|bytes={path.stat().st_size}"
    f"|periodo={periodo}"
)
PY

  validacion_local_rc=$?

  if [[ "$validacion_local_rc" -ne 0 ]]; then
    echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=$validacion_local_rc|paso=validacion_local_json"
    exit "$validacion_local_rc"
  fi

  modo_reintento=0

  if [[ -s "$VALIDACION_REMOTA" ]]; then
    modo_reintento=1
    echo
    echo "SUBIENDO Y VERIFICANDO CIERRE MENSUAL EN MODO REINTENTO..."
  else
    echo
    echo "SUBIENDO Y VERIFICANDO CIERRE MENSUAL COMPLETO..."
  fi

  "$PYTHON_BIN" - \
    "$SCRIPT_SUBIDA" \
    "$ENV_FILE" \
    "$INICIO" \
    "$FIN" \
    "$modo_reintento" <<'PY' 2>&1 |
import os
import runpy
import sys
from pathlib import Path

from dotenv import load_dotenv


script = Path(sys.argv[1])
env_file = Path(sys.argv[2])
inicio = sys.argv[3]
fin = sys.argv[4]
modo_reintento = sys.argv[5] == "1"

os.environ["FACTURAS_ENV_FILE"] = str(env_file)

if not load_dotenv(
    dotenv_path=env_file,
    override=True,
    encoding="utf-8",
):
    raise RuntimeError(
        "No fue posible cargar facturas.env."
    )

argumentos = [
    str(script),
    "--inicio",
    inicio,
    "--fin",
    fin,
]

if modo_reintento:
    argumentos.append(
        "--reintentar-fallidos"
    )

sys.argv = argumentos

runpy.run_path(
    str(script),
    run_name="__main__",
)
PY
  sed -E \
    -e 's/(Drive backup validado:[^|]*\|).*/\1 ***REDACTADO***/' \
    -e 's#(Repositorio destino:[^|]*\|).*#\1 ***REDACTADO***#' \
    -e 's#https://[^[:space:]]+#***URL_REDACTADA***#g'

  subida_rc="${PIPESTATUS[0]}"

  echo "CRON_CIERRE_MENSUAL_SUBIDA_FIN|$(date --iso-8601=seconds)|rc=$subida_rc|reintento=$modo_reintento"

  if [[ "$subida_rc" -ne 0 ]]; then
    echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=$subida_rc|paso=subida_remota"
    exit "$subida_rc"
  fi

  if [[ ! -s "$VALIDACION_REMOTA" ]]; then
    echo "ERROR: falta la validación remota mensual."
    echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=79|paso=validacion_remota"
    exit 79
  fi

  if [[ ! -s "$ESTADO_RETENCION" ]]; then
    echo "ERROR: falta el estado de retención mensual."
    echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=80|paso=estado_retencion"
    exit 80
  fi

  "$PYTHON_BIN" - \
    "$VALIDACION_REMOTA" \
    "$ESTADO_RETENCION" \
    "$PERIODO" <<'PY'
import json
import sys
from pathlib import Path


validacion_path = Path(sys.argv[1])
retencion_path = Path(sys.argv[2])
periodo = sys.argv[3]

validacion = json.loads(
    validacion_path.read_text(
        encoding="utf-8"
    )
)

retencion = json.loads(
    retencion_path.read_text(
        encoding="utf-8"
    )
)


def entero(payload, nombre):
    valor = payload.get(nombre)

    if (
        not isinstance(valor, int)
        or isinstance(valor, bool)
    ):
        raise RuntimeError(
            f"Campo entero inválido: {nombre}"
        )

    return valor


if validacion.get("ok") is not True:
    raise RuntimeError(
        "La validación remota mensual no está en estado OK."
    )

if validacion.get("periodo") != periodo:
    raise RuntimeError(
        "El periodo de la validación remota no coincide."
    )

esperados = entero(
    validacion,
    "total_archivos_esperados",
)
verificados = entero(
    validacion,
    "total_archivos_verificados",
)
fallidos = entero(
    validacion,
    "total_archivos_fallidos",
)
total_resultados = entero(
    validacion,
    "total_resultados",
)

resultados = validacion.get("resultados")

if not isinstance(resultados, list):
    raise RuntimeError(
        "La validación remota no contiene una lista de resultados."
    )

if esperados <= 0:
    raise RuntimeError(
        "La validación remota no contiene archivos esperados."
    )

if verificados != esperados:
    raise RuntimeError(
        "La cantidad verificada no coincide con la esperada."
    )

if total_resultados != esperados:
    raise RuntimeError(
        "La cantidad de resultados no coincide con la esperada."
    )

if len(resultados) != esperados:
    raise RuntimeError(
        "La lista de resultados no coincide con la cantidad esperada."
    )

if fallidos != 0:
    raise RuntimeError(
        "La validación remota mensual contiene archivos fallidos."
    )

if retencion.get("ok") is not True:
    raise RuntimeError(
        "El estado de retención mensual no está en estado OK."
    )

if retencion.get("periodo") != periodo:
    raise RuntimeError(
        "El periodo del estado de retención no coincide."
    )

if (
    retencion.get(
        "validacion_remota_publicada_y_verificada"
    )
    is not True
):
    raise RuntimeError(
        "La validación remota no figura como publicada y verificada."
    )

dias_retencion = entero(
    retencion,
    "retencion_local_dias",
)

if dias_retencion <= 0:
    raise RuntimeError(
        "La retención local mensual es inválida."
    )

if not retencion.get(
    "eliminacion_local_permitida_desde"
):
    raise RuntimeError(
        "Falta la fecha permitida de eliminación local."
    )

print(
    "VALIDACION_REMOTA_MENSUAL_OK"
    f"|archivo={validacion_path.name}"
    f"|bytes={validacion_path.stat().st_size}"
    f"|esperados={esperados}"
    f"|verificados={verificados}"
    f"|fallidos={fallidos}"
)

print(
    "RETENCION_MENSUAL_OK"
    f"|archivo={retencion_path.name}"
    f"|bytes={retencion_path.stat().st_size}"
    f"|dias={dias_retencion}"
    f"|eliminacion_desde="
    f"{retencion.get('eliminacion_local_permitida_desde')}"
)
PY

  validacion_remota_rc=$?

  if [[ "$validacion_remota_rc" -ne 0 ]]; then
    echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=$validacion_remota_rc|paso=validacion_remota_json"
    exit "$validacion_remota_rc"
  fi

  echo "CRON_CIERRE_MENSUAL_FIN|$(date --iso-8601=seconds)|rc=0|periodo=$PERIODO|inicio=$INICIO|fin=$FIN"
  echo "===================================================================================================="

  exit 0

} >> "$LOG_FILE" 2>&1
