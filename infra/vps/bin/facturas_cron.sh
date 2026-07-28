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
RUNNER="/opt/joyco/facturas-procesador/bin/facturas_runner.py"

LOG_DIR="/var/log/joyco/facturas-procesador/cron"
LOCK_FILE="/var/lib/joyco/facturas-procesador/state/locks/cron_facturas_principal.lock"

mkdir -p "$LOG_DIR"

LOG_FILE="$LOG_DIR/facturas_$(date +%F).log"

{
  echo
  echo "============================================================"
  echo "CRON_FACTURAS_INICIO|$(date --iso-8601=seconds)|pid=$$"
  echo "============================================================"

  if ! cd "$APP_DIR"; then
    echo "CRON_FACTURAS_FIN|$(date --iso-8601=seconds)|rc=71|error=cd"
    exit 71
  fi

  exec 9>"$LOCK_FILE"

  if ! /usr/bin/flock -n 9; then
    echo "CRON_FACTURAS_SKIP|$(date --iso-8601=seconds)|motivo=ejecucion_anterior_activa"
    exit 0
  fi

  "$RUNNER" run --confirm-production
  rc=$?

  echo "CRON_FACTURAS_FIN|$(date --iso-8601=seconds)|rc=$rc"
  echo "============================================================"

  exit "$rc"

} >> "$LOG_FILE" 2>&1
