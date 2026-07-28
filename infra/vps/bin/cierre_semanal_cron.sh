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

APP_DIR="/opt/joyco/facturas-procesador/app"
PYTHON_BIN="/opt/joyco/facturas-procesador/venv/bin/python"
ENV_FILE="/etc/joyco/facturas-procesador/facturas.env"

SCRIPT_CIERRE="$APP_DIR/scripts/cierre_semanal_facturas.py"
SCRIPT_SUBIDA="$APP_DIR/scripts/subir_cierre_semanal_sharepoint.py"

CIERRE_ROOT="/var/lib/joyco/facturas-procesador/cierres/diarios"

LOG_DIR="/var/log/joyco/facturas-procesador/cron"
LOCK_DIR="/var/lib/joyco/facturas-procesador/state/locks"

LOCK_PRINCIPAL="$LOCK_DIR/cron_facturas_principal.lock"
LOCK_SEMANAL="$LOCK_DIR/cron_cierre_semanal.lock"
LOCK_APP="/var/lib/joyco/facturas-procesador/state/aprobadas.lock"

MODE="${1:-run}"
FECHA_EJECUCION="$(date +%F)"
LOG_FILE="$LOG_DIR/cierre_semanal_${FECHA_EJECUCION}.log"

mkdir -p \
  "$LOG_DIR" \
  "$LOCK_DIR"

calcular_semana_anterior() {
  "$PYTHON_BIN" - <<'PY'
from datetime import date, timedelta


today = date.today()

fin = today - timedelta(
    days=today.weekday() + 1
)

inicio = fin - timedelta(days=6)

print(
    inicio.isoformat(),
    fin.isoformat(),
)
PY
}

read -r INICIO FIN < <(
  calcular_semana_anterior
)

check_environment() {
  local requerido

  for requerido in \
    "$APP_DIR" \
    "$PYTHON_BIN" \
    "$ENV_FILE" \
    "$SCRIPT_CIERRE" \
    "$SCRIPT_SUBIDA"; do

    if [[ ! -e "$requerido" ]]; then
      echo "ERROR: falta $requerido"
      return 1
    fi
  done

  "$PYTHON_BIN" -m py_compile \
    "$SCRIPT_CIERRE" \
    "$SCRIPT_SUBIDA"

  return 0
}

if [[ "$MODE" == "--check" ]]; then
  echo "============================================================"
  echo "COMPROBACIÓN DEL WRAPPER SEMANAL"
  echo "============================================================"

  check_environment

  echo "Fecha de comprobación: $FECHA_EJECUCION"
  echo "Última semana completa calculada: $INICIO a $FIN"
  echo "Python: $PYTHON_BIN"
  echo "Generador: $SCRIPT_CIERRE"
  echo "Uploader: $SCRIPT_SUBIDA"
  echo "Log: $LOG_FILE"
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
  echo "============================================================"
  echo "CRON_CIERRE_SEMANAL_INICIO|$(date --iso-8601=seconds)|inicio=$INICIO|fin=$FIN|pid=$$"
  echo "============================================================"

  if ! check_environment; then
    echo "CRON_CIERRE_SEMANAL_FIN|$(date --iso-8601=seconds)|rc=70|paso=entorno"
    exit 70
  fi

  if ! cd "$APP_DIR"; then
    echo "CRON_CIERRE_SEMANAL_FIN|$(date --iso-8601=seconds)|rc=71|paso=directorio"
    exit 71
  fi

  exec 9>"$LOCK_SEMANAL"

  if ! /usr/bin/flock -n 9; then
    echo "CRON_CIERRE_SEMANAL_SKIP|$(date --iso-8601=seconds)|motivo=semanal_activo"
    exit 0
  fi

  echo "Esperando acceso exclusivo al flujo principal y cierres..."

  exec 8>"$LOCK_PRINCIPAL"

  if ! /usr/bin/flock -w 3600 8; then
    echo "CRON_CIERRE_SEMANAL_FIN|$(date --iso-8601=seconds)|rc=75|paso=espera_principal"
    exit 75
  fi

  echo "OK: acceso exclusivo obtenido."

  if [[ -e "$LOCK_APP" ]]; then
    echo "ERROR: existe aprobadas.lock."
    echo "CRON_CIERRE_SEMANAL_FIN|$(date --iso-8601=seconds)|rc=76|paso=lock_aplicacion"
    exit 76
  fi

  echo
  echo "GENERANDO CIERRE SEMANAL LOCAL..."

  "$PYTHON_BIN" - \
    "$SCRIPT_CIERRE" \
    "$ENV_FILE" \
    "$INICIO" \
    "$FIN" <<'PY'
import runpy
import sys
from pathlib import Path

from dotenv import load_dotenv


script = Path(sys.argv[1])
env_file = Path(sys.argv[2])
inicio = sys.argv[3]
fin = sys.argv[4]

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

  echo "CRON_CIERRE_SEMANAL_LOCAL_FIN|$(date --iso-8601=seconds)|rc=$cierre_rc"

  if [[ "$cierre_rc" -ne 0 ]]; then
    echo "CRON_CIERRE_SEMANAL_FIN|$(date --iso-8601=seconds)|rc=$cierre_rc|paso=cierre_local"
    exit "$cierre_rc"
  fi

  mapfile -t cierres_encontrados < <(
    find "$CIERRE_ROOT" \
      -type d \
      -path "*/SEMANA_${INICIO}_a_${FIN}/Semanal" \
      -print
  )

  if [[ "${#cierres_encontrados[@]}" -ne 1 ]]; then
    echo "ERROR: carpetas semanales encontradas=${#cierres_encontrados[@]}"
    echo "CRON_CIERRE_SEMANAL_FIN|$(date --iso-8601=seconds)|rc=77|paso=carpeta_local"
    exit 77
  fi

  cierre_dir="${cierres_encontrados[0]}"

  validacion_local="$(
    find "$cierre_dir" \
      -type f \
      -name "validacion_local_semanal_${INICIO}_a_${FIN}.json" \
      -print \
      -quit
  )"

  if [[ -z "$validacion_local" ]] || [[ ! -s "$validacion_local" ]]; then
    echo "CRON_CIERRE_SEMANAL_FIN|$(date --iso-8601=seconds)|rc=78|paso=validacion_local"
    exit 78
  fi

  "$PYTHON_BIN" - "$validacion_local" <<'PY'
import json
import sys
from pathlib import Path


path = Path(sys.argv[1])

payload = json.loads(
    path.read_text(
        encoding="utf-8"
    )
)

if payload.get("ok") is not True:
    raise RuntimeError(
        "La validación local semanal no está en estado OK."
    )

print(
    "VALIDACION_LOCAL_SEMANAL_OK"
    f"|archivo={path.name}"
    f"|bytes={path.stat().st_size}"
)
PY

  validacion_local_rc=$?

  if [[ "$validacion_local_rc" -ne 0 ]]; then
    echo "CRON_CIERRE_SEMANAL_FIN|$(date --iso-8601=seconds)|rc=$validacion_local_rc|paso=validacion_local_json"
    exit "$validacion_local_rc"
  fi

  echo
  echo "SUBIENDO Y VERIFICANDO CIERRE SEMANAL..."

  "$PYTHON_BIN" - \
    "$SCRIPT_SUBIDA" \
    "$ENV_FILE" \
    "$INICIO" \
    "$FIN" <<'PY' 2>&1 |
import runpy
import sys
from pathlib import Path

from dotenv import load_dotenv


script = Path(sys.argv[1])
env_file = Path(sys.argv[2])
inicio = sys.argv[3]
fin = sys.argv[4]

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
]

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

  echo "CRON_CIERRE_SEMANAL_SUBIDA_FIN|$(date --iso-8601=seconds)|rc=$subida_rc"

  if [[ "$subida_rc" -ne 0 ]]; then
    echo "CRON_CIERRE_SEMANAL_FIN|$(date --iso-8601=seconds)|rc=$subida_rc|paso=subida_remota"
    exit "$subida_rc"
  fi

  validacion_remota="$(
    find "$cierre_dir" \
      -type f \
      -name "validacion_remota_semanal_${INICIO}_a_${FIN}.json" \
      -print \
      -quit
  )"

  if [[ -z "$validacion_remota" ]] || [[ ! -s "$validacion_remota" ]]; then
    echo "CRON_CIERRE_SEMANAL_FIN|$(date --iso-8601=seconds)|rc=79|paso=validacion_remota"
    exit 79
  fi

  "$PYTHON_BIN" - "$validacion_remota" <<'PY'
import json
import sys
from pathlib import Path


path = Path(sys.argv[1])

payload = json.loads(
    path.read_text(
        encoding="utf-8"
    )
)


def values(value, target):
    found = []

    if isinstance(value, dict):
        for key, child in value.items():
            if str(key).casefold() == target.casefold():
                found.append(child)

            found.extend(values(child, target))

    elif isinstance(value, list):
        for child in value:
            found.extend(values(child, target))

    return found


ok_values = values(payload, "ok")

error_values = (
    values(payload, "errores")
    + values(payload, "errors")
    + values(payload, "fallidos")
)

if any(value is False for value in ok_values):
    raise RuntimeError(
        "La validación remota semanal contiene OK=False."
    )

nonempty = [
    value
    for value in error_values
    if value not in (
        None,
        "",
        [],
        {},
        0,
    )
]

if nonempty:
    raise RuntimeError(
        "La validación remota semanal contiene errores."
    )

print(
    "VALIDACION_REMOTA_SEMANAL_OK"
    f"|archivo={path.name}"
    f"|bytes={path.stat().st_size}"
    f"|valores_ok={len(ok_values)}"
)
PY

  validacion_remota_rc=$?

  if [[ "$validacion_remota_rc" -ne 0 ]]; then
    echo "CRON_CIERRE_SEMANAL_FIN|$(date --iso-8601=seconds)|rc=$validacion_remota_rc|paso=validacion_remota_json"
    exit "$validacion_remota_rc"
  fi

  echo "CRON_CIERRE_SEMANAL_FIN|$(date --iso-8601=seconds)|rc=0|inicio=$INICIO|fin=$FIN"
  echo "============================================================"

  exit 0

} >> "$LOG_FILE" 2>&1
