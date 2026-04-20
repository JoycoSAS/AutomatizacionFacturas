import os
import csv
import datetime
from typing import Dict, List, Optional


def _today_str() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d")


def _ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def _csv_path(base_dir: str, prefix: str, date_str: Optional[str] = None) -> str:
    date_str = date_str or _today_str()
    return os.path.join(base_dir, f"{prefix}_{date_str}.csv")


def _append_row(csv_path: str, fieldnames: List[str], row: Dict[str, object]):
    is_new = not os.path.exists(csv_path)
    with open(csv_path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(
            f,
            fieldnames=fieldnames,
            delimiter=",",
            quoting=csv.QUOTE_MINIMAL,
            extrasaction="ignore"
        )
        if is_new:
            w.writeheader()
        w.writerow(row)


def _ensure_header(csv_path: str, fieldnames: List[str]):
    if os.path.exists(csv_path):
        return
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(
            f,
            fieldnames=fieldnames,
            delimiter=",",
            quoting=csv.QUOTE_MINIMAL,
            extrasaction="ignore"
        )
        w.writeheader()


def append_run_summary(audit_dir: str, prefix: str, row: Dict[str, object]):
    _ensure_dir(audit_dir)
    path = _csv_path(audit_dir, prefix)

    fieldnames = [
        "run_id",
        "inicio",
        "fin",
        "duracion_s",
        "carpeta",
        "since_days",
        "max_aprobados",
        "max_zip_buscar",
        "msgs_leidos",
        "msgs_pendientes",
        "msgs_procesados",
        "ok",
        "sin_match",
        "ya_registrado",
        "sin_pdf",
        "errores",
        "dian_pdf_only",
        "nuevos_total",
        "enriquecidas_total",

        # NUEVO RESUMEN EXPLÍCITO
        "match_total",
        "ok_match",
        "dian_match",
        "facturas_con_filas",
        "facturas_sin_filas",
        "ok_con_filas",
        "ok_sin_filas",
        "dian_con_filas",
        "dian_sin_filas",

        "nota",
    ]

    _append_row(path, fieldnames, row)


def append_detalle_rows(
    audit_dir: str,
    prefix: str,
    rows: List[Dict[str, object]],
    create_if_empty: bool = True
):
    _ensure_dir(audit_dir)
    path = _csv_path(audit_dir, prefix)

    fieldnames = [
        "run_id",
        "fecha_hora",
        "msg_id",
        "subject",
        "pdf_elegido",
        "cufe",
        "numero",
        "fecha_factura",
        "zip_match",
        "estado",

        # NUEVAS COLUMNAS DETALLE
        "tipo_resultado",
        "filas_generadas",
        "motivo_no_registro",

        "duracion_s",
        "nuevos",
        "enriquecidas",
        "fuente",
        "error",
    ]

    if create_if_empty:
        _ensure_header(path, fieldnames)

    if not rows:
        return

    for r in rows:
        _append_row(path, fieldnames, r)