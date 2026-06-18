# -*- coding: utf-8 -*-
"""
Prueba controlada del módulo de notificaciones.

Por defecto corre en dry-run y NO envía correo.
Para enviar realmente:
python scripts/probar_notificador.py --tipo fallo --enviar
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from utils.notificador import (
    VERSION_NOTIFICADOR,
    enviar_notificacion,
    notificar_exito_control,
    notificar_fallo,
)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--tipo",
        choices=["fallo", "backup-ok", "trimestral-ok", "info"],
        default="fallo",
    )
    parser.add_argument(
        "--enviar",
        action="store_true",
        help="Envía realmente según .env. Si no se usa, corre en dry-run.",
    )
    args = parser.parse_args()

    dry_run = not args.enviar

    print("=" * 100)
    print("PRUEBA NOTIFICADOR")
    print("=" * 100)
    print(f"Versión: {VERSION_NOTIFICADOR}")
    print(f"Dry run: {dry_run}")
    print("-" * 100)

    if args.tipo == "fallo":
        res = notificar_fallo(
            proceso="FACTURAS_PRODUCCION",
            asunto="Prueba de alerta de falla",
            mensaje="Esta es una prueba controlada de alerta por falla. No corresponde a una falla real.",
            detalle={
                "archivo": "main_aprobadas_integrado.py",
                "tipo_prueba": "fallo_controlado",
                "accion_requerida": "Validar recepción del correo o salida dry-run.",
            },
            dry_run=dry_run,
        )

    elif args.tipo == "backup-ok":
        res = notificar_exito_control(
            proceso="BACKUP_MENSUAL_LOCAL",
            asunto="Prueba informativa de backup mensual exitoso",
            mensaje="Esta es una prueba controlada de correo informativo de backup mensual exitoso.",
            detalle={
                "archivo": "backup_mensual_YYYY_MM.zip",
                "verificacion": "SHA256 exacto",
                "tipo_prueba": "exito_controlado",
            },
            dry_run=dry_run,
        )

    elif args.tipo == "trimestral-ok":
        res = notificar_exito_control(
            proceso="CIERRE_TRIMESTRAL_SHAREPOINT",
            asunto="Prueba informativa de cierre trimestral exitoso",
            mensaje="Esta es una prueba controlada de correo informativo de cierre trimestral exitoso.",
            detalle={
                "periodo": "2026-T2",
                "verificacion": "Excel por datos internos y soportes por SHA256",
                "tipo_prueba": "exito_controlado",
            },
            dry_run=dry_run,
        )

    else:
        res = enviar_notificacion(
            proceso="SISTEMA",
            estado="INFO",
            asunto="Prueba informativa general",
            mensaje="Esta es una prueba general del módulo de notificaciones.",
            detalle={"tipo_prueba": "info_controlado"},
            dry_run=dry_run,
        )

    print("Resultado:")
    print(res)
    print("=" * 100)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

