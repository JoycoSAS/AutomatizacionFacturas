# -*- coding: utf-8 -*-
"""
JOYCO - Facturas Procesador
Utilidad compartida para resolver el trimestre operativo activo.

La jerarquía oficial de cierres es:

    data/cierres_diarios/<AÑO_INICIO_TRIMESTRE>/
        TRIMESTRE_<FECHA_INICIO>_A_<FECHA_FIN>/
            <MES>/
                <SEMANA>/
                    <DIARIO>/

El periodo se obtiene exclusivamente de:
    data/state/cierre_trimestral_state.json

No se calcula un trimestre calendario por cuenta propia. Esto evita que los
cierres diario, semanal, mensual y trimestral terminen en rutas distintas.
"""

from __future__ import annotations

import datetime as _dt
import json
from pathlib import Path
from typing import Any


CAMPOS_REQUERIDOS = (
    "periodo_activo",
    "fecha_inicio_periodo_activo",
    "proximo_cierre_estimado",
    "estado",
)


def _parse_fecha(valor: Any, campo: str) -> _dt.date:
    texto = str(valor or "").strip()
    try:
        return _dt.date.fromisoformat(texto)
    except ValueError as exc:
        raise RuntimeError(
            f"Fecha inválida en {campo}: {texto!r}. Se esperaba YYYY-MM-DD."
        ) from exc


def cargar_trimestre_activo(root: Path, fecha_operacion: _dt.date) -> dict[str, Any]:
    """Carga y valida el trimestre activo aplicable a ``fecha_operacion``.

    La función falla de forma segura cuando el estado no existe, está
    incompleto, no está ACTIVO o la fecha solicitada queda fuera del rango.
    Nunca usa silenciosamente la estructura antigua.
    """

    state_path = Path(root) / "data" / "state" / "cierre_trimestral_state.json"

    if not state_path.exists() or not state_path.is_file():
        raise RuntimeError(
            "No existe el estado trimestral requerido: "
            f"{state_path}. No se creará el cierre fuera de un trimestre."
        )

    try:
        estado = json.loads(state_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"El estado trimestral no contiene JSON válido: {state_path}"
        ) from exc

    if not isinstance(estado, dict):
        raise RuntimeError(
            f"El estado trimestral debe ser un objeto JSON: {state_path}"
        )

    faltantes = [campo for campo in CAMPOS_REQUERIDOS if not str(estado.get(campo) or "").strip()]
    if faltantes:
        raise RuntimeError(
            "El estado trimestral está incompleto. Faltan: "
            + ", ".join(faltantes)
        )

    estado_operativo = str(estado["estado"]).strip().upper()
    if estado_operativo != "ACTIVO":
        raise RuntimeError(
            "El trimestre no está ACTIVO. "
            f"Estado encontrado: {estado_operativo!r}."
        )

    inicio = _parse_fecha(
        estado["fecha_inicio_periodo_activo"],
        "fecha_inicio_periodo_activo",
    )
    fin = _parse_fecha(
        estado["proximo_cierre_estimado"],
        "proximo_cierre_estimado",
    )

    if inicio > fin:
        raise RuntimeError(
            "El estado trimestral es inválido: la fecha de inicio es posterior "
            "a la fecha de cierre."
        )

    if not (inicio <= fecha_operacion <= fin):
        raise RuntimeError(
            "La fecha solicitada no pertenece al trimestre activo. "
            f"Fecha={fecha_operacion.isoformat()} | "
            f"Trimestre={inicio.isoformat()} a {fin.isoformat()}."
        )

    nombre_carpeta = (
        f"TRIMESTRE_{inicio.isoformat()}_A_{fin.isoformat()}"
    )

    return {
        "periodo_activo": str(estado["periodo_activo"]).strip(),
        "estado": estado_operativo,
        "fecha_inicio": inicio.isoformat(),
        "fecha_fin": fin.isoformat(),
        "anio": str(inicio.year),
        "nombre_carpeta": nombre_carpeta,
        "ruta_relativa": (Path(str(inicio.year)) / nombre_carpeta).as_posix(),
        "state_path": str(state_path),
    }
