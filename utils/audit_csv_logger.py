import csv
import os
import datetime
from typing import Dict, List, Any, Iterable


# Orden preferido para audit_detalle.
DETALLE_BASE_COLUMNS = [
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
    "tipo_resultado",
    "filas_generadas",
    "motivo_no_registro",

    # Calidad / completitud
    "calidad_datos",
    "score_calidad",
    "campos_faltantes",
    "alertas_calidad",
    "total_detectado_calidad",

    "duracion_s",
    "nuevos",
    "enriquecidas",
    "fuente",
    "error",
]


# Orden preferido para audit_runs.
RUNS_BASE_COLUMNS = [
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
    "filas_local_total",
    "filas_web_total",

    "total_match",
    "match_total",

    "ok_total",
    "ok_match",
    "ok_registradas",
    "ok_con_filas",
    "ok_no_registrables",
    "ok_sin_filas",

    "dian_total",
    "dian_match",
    "dian_registradas",
    "dian_con_filas",
    "dian_no_registrables",
    "dian_sin_filas",

    "facturas_con_filas",
    "facturas_sin_registro",
    "facturas_sin_filas",

    # Calidad / completitud
    "calidad_promedio_pct",
    "facturas_calidad_total",
    "facturas_calidad_completa",
    "facturas_calidad_parcial",
    "facturas_calidad_minima",
    "facturas_calidad_sin_filas",
    "facturas_total_cero_o_vacio",
    "facturas_sin_cufe",
    "facturas_sin_nit",
    "facturas_sin_cliente",

    "nota",
]


def _today_str() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d")


def _csv_path(audit_dir: str, prefix: str) -> str:
    os.makedirs(audit_dir, exist_ok=True)
    return os.path.join(audit_dir, f"{prefix}_{_today_str()}.csv")


def _normalizar_valor_csv(valor: Any) -> Any:
    if valor is None:
        return ""

    if isinstance(valor, bool):
        return "true" if valor else "false"

    if isinstance(valor, (list, tuple, set)):
        return ";".join(str(x) for x in valor)

    if isinstance(valor, dict):
        return str(valor)

    return valor


def _read_existing(path: str) -> tuple[list[str], list[dict]]:
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        return [], []

    try:
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            fieldnames = list(reader.fieldnames or [])
            rows = [dict(r) for r in reader]
            return fieldnames, rows
    except Exception:
        # Fallback para archivos creados sin BOM o con alguna codificación previa.
        with open(path, "r", encoding="utf-8", errors="replace", newline="") as f:
            reader = csv.DictReader(f)
            fieldnames = list(reader.fieldnames or [])
            rows = [dict(r) for r in reader]
            return fieldnames, rows


def _ordered_union(
    *,
    preferred: list[str],
    existing: list[str],
    incoming_rows: Iterable[dict],
) -> list[str]:
    incoming_keys: list[str] = []
    seen_incoming = set()

    for row in incoming_rows:
        for k in row.keys():
            if k not in seen_incoming:
                seen_incoming.add(k)
                incoming_keys.append(k)

    result: list[str] = []
    seen = set()

    for source in (preferred, existing, incoming_keys):
        for col in source:
            if not col:
                continue
            if col in seen:
                continue
            seen.add(col)
            result.append(col)

    return result


def _clean_row(row: Dict[str, Any], fieldnames: list[str]) -> dict:
    row = row or {}
    return {col: _normalizar_valor_csv(row.get(col, "")) for col in fieldnames}


def _append_rows_dynamic(
    *,
    path: str,
    rows: list[dict],
    preferred_columns: list[str],
) -> None:
    if not rows:
        return

    rows = [dict(r or {}) for r in rows if isinstance(r, dict)]
    if not rows:
        return

    existing_cols, existing_rows = _read_existing(path)

    # Unión dinámica:
    # Si el controller agrega columnas nuevas, el CSV se reescribe con esas columnas.
    fieldnames = _ordered_union(
        preferred=preferred_columns,
        existing=existing_cols,
        incoming_rows=rows,
    )

    all_rows = existing_rows + rows

    tmp_path = f"{path}.tmp"
    with open(tmp_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=fieldnames,
            extrasaction="ignore",
            lineterminator="\n",
        )
        writer.writeheader()
        for row in all_rows:
            writer.writerow(_clean_row(row, fieldnames))

    os.replace(tmp_path, path)


def append_detalle_rows(audit_dir: str, prefix: str, rows: List[Dict[str, Any]]) -> None:
    """
    Agrega filas al audit_detalle diario.

    Importante:
    - Soporta columnas nuevas dinámicamente.
    - Si el CSV ya existía con encabezados antiguos, lo reescribe conservando
      las filas previas y agregando las columnas nuevas.
    """
    if not rows:
        return

    path = _csv_path(audit_dir, prefix)
    _append_rows_dynamic(
        path=path,
        rows=rows,
        preferred_columns=DETALLE_BASE_COLUMNS,
    )


def append_run_summary(audit_dir: str, prefix: str, row: Dict[str, Any]) -> None:
    """
    Agrega una fila al audit_runs diario.

    Importante:
    - Soporta columnas nuevas dinámicamente.
    - No descarta métricas nuevas como calidad_promedio_pct,
      facturas_calidad_completa, etc.
    """
    if not row:
        return

    path = _csv_path(audit_dir, prefix)
    _append_rows_dynamic(
        path=path,
        rows=[row],
        preferred_columns=RUNS_BASE_COLUMNS,
    )
