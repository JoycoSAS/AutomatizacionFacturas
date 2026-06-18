# -*- coding: utf-8 -*-
"""
JOYCO - Facturas Procesador
Wrapper central para ejecutar procesos con alertas/notificaciones.

Uso general:
python scripts/ejecutar_con_alertas.py --proceso FACTURAS_PRODUCCION --asunto "Proceso facturas" -- python main_aprobadas_integrado.py

Política:
- Si el comando falla: notifica ERROR.
- Si el comando sale OK:
  - Solo notifica éxito para procesos de control definidos en utils/notificador.py:
    backup mensual, cierre trimestral, reemplazo Excel activo.
  - Para procesos frecuentes, no notifica éxito.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from utils.notificador import (
    VERSION_NOTIFICADOR,
    detalle_excepcion,
    notificar_exito_control,
    notificar_fallo,
)

VERSION = "2026-06-18-EJECUTOR-CON-ALERTAS-V1"


def truncar(txt: str | None, limite: int = 8000) -> str:
    if not txt:
        return ""

    txt = str(txt)

    if len(txt) <= limite:
        return txt

    return txt[-limite:]


def limpiar_comando(cmd: list[str]) -> list[str]:
    if cmd and cmd[0] == "--":
        return cmd[1:]
    return cmd


def imprimir_bloque(titulo: str, contenido: str | None) -> None:
    contenido = contenido or ""

    if not contenido.strip():
        return

    print("-" * 100)
    print(titulo)
    print("-" * 100)
    print(contenido.rstrip())


def ejecutar_comando(cmd: list[str]) -> tuple[int, str, str, float]:
    inicio = time.perf_counter()

    proceso = subprocess.run(
        cmd,
        cwd=str(ROOT),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        shell=False,
    )

    duracion = time.perf_counter() - inicio

    return proceso.returncode, proceso.stdout or "", proceso.stderr or "", duracion


def detalle_ejecucion(
    *,
    cmd: list[str],
    exit_code: int,
    duracion_seg: float,
    stdout: str,
    stderr: str,
) -> dict[str, Any]:
    return {
        "root": str(ROOT),
        "comando": cmd,
        "exit_code": exit_code,
        "duracion_segundos": round(duracion_seg, 2),
        "stdout_ultimos_caracteres": truncar(stdout, 8000),
        "stderr_ultimos_caracteres": truncar(stderr, 8000),
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Ejecuta un proceso del proyecto y notifica fallas o éxitos de control."
    )

    parser.add_argument("--proceso", required=True, help="Nombre lógico del proceso.")
    parser.add_argument("--asunto", required=True, help="Asunto de la alerta/notificación.")
    parser.add_argument(
        "--mensaje-ok",
        default="El proceso finalizó correctamente.",
        help="Mensaje usado si el proceso sale OK y aplica notificación informativa.",
    )
    parser.add_argument(
        "--mensaje-error",
        default="El proceso falló durante la ejecución.",
        help="Mensaje usado si el proceso falla.",
    )
    parser.add_argument(
        "comando",
        nargs=argparse.REMAINDER,
        help="Comando a ejecutar después de --",
    )

    args = parser.parse_args()

    proceso = str(args.proceso or "").strip().upper()
    asunto = str(args.asunto or "").strip()
    cmd = limpiar_comando(args.comando)

    print("=" * 100)
    print("EJECUTOR CON ALERTAS")
    print("=" * 100)
    print(f"Versión ejecutor: {VERSION}")
    print(f"Versión notificador: {VERSION_NOTIFICADOR}")
    print(f"Root: {ROOT}")
    print(f"Proceso: {proceso}")
    print(f"Asunto: {asunto}")
    print(f"Comando: {' '.join(cmd) if cmd else '(vacío)'}")
    print("-" * 100)

    if not cmd:
        print("❌ No se recibió comando para ejecutar. Usa -- antes del comando real.")
        notificar_fallo(
            proceso=proceso,
            asunto=asunto or "Comando vacío",
            mensaje="No se recibió comando para ejecutar en el wrapper de alertas.",
            detalle={"root": str(ROOT), "ayuda": "Usa: -- python archivo.py"},
        )
        return 2

    try:
        exit_code, stdout, stderr, duracion = ejecutar_comando(cmd)

        imprimir_bloque("STDOUT", stdout)
        imprimir_bloque("STDERR", stderr)

        detalle = detalle_ejecucion(
            cmd=cmd,
            exit_code=exit_code,
            duracion_seg=duracion,
            stdout=stdout,
            stderr=stderr,
        )

        print("-" * 100)
        print(f"ExitCode: {exit_code}")
        print(f"Duración: {duracion:.2f} segundos")

        if exit_code != 0:
            print("❌ Proceso finalizó con error. Enviando/registrando alerta...")
            res = notificar_fallo(
                proceso=proceso,
                asunto=asunto,
                mensaje=args.mensaje_error,
                detalle=detalle,
            )
            print("Resultado notificación:")
            print(json.dumps(res, ensure_ascii=False, indent=2, default=str))
            print("=" * 100)
            return exit_code

        print("✅ Proceso finalizó correctamente.")

        res = notificar_exito_control(
            proceso=proceso,
            asunto=asunto,
            mensaje=args.mensaje_ok,
            detalle=detalle,
        )

        if res.get("motivo") == "EXITO_NO_INFORMATIVO_POR_POLITICA":
            print("🔕 Éxito no notificado por política del sistema.")
        else:
            print("Resultado notificación de éxito:")
            print(json.dumps(res, ensure_ascii=False, indent=2, default=str))

        print("=" * 100)
        return 0

    except Exception as exc:
        print(f"❌ Error inesperado en ejecutor con alertas: {exc}")

        res = notificar_fallo(
            proceso=proceso,
            asunto=asunto or "Error inesperado en ejecutor",
            mensaje="El wrapper de ejecución con alertas falló inesperadamente.",
            detalle=detalle_excepcion(exc),
        )

        print("Resultado notificación:")
        print(json.dumps(res, ensure_ascii=False, indent=2, default=str))
        print("=" * 100)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
