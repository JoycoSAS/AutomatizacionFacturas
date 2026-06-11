import csv
import os
import datetime
from typing import Dict, List, Any, Iterable


AUDIT_LOGGER_VERSION = "2026-06-11-AUDIT-TIEMPOS-CALIDAD-NOEMPTY"


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

    # Calidad / completitud por factura.
    "calidad_datos",
    "score_calidad",
    "campos_faltantes",
    "alertas_calidad",
    "total_detectado_calidad",

    # Tiempos por factura.
    "duracion_s",
    "tiempo_factura_s",
    "tiempo_match_y_registro_s",

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

    # Tiempos por etapa de ejecución.
    "tiempo_total_s",
    "tiempo_aidx_s",
    "tiempo_procesamiento_facturas_s",
    "tiempo_limpieza_s",
    "tiempo_audit_s",
    "tiempo_promedio_factura_s",

    # Métricas del AttachmentIndexStore / AIDX.
    "aidx_cufe_count",
    "aidx_num_count",
    "aidx_match_count",
    "audit_detalle_rows",

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

    # Calidad / completitud del lote.
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


_NUMERIC_DEFAULT_ZERO_FIELDS = {
    "duracion_s",
    "tiempo_total_s",
    "tiempo_aidx_s",
    "tiempo_procesamiento_facturas_s",
    "tiempo_limpieza_s",
    "tiempo_audit_s",
    "tiempo_promedio_factura_s",
    "aidx_cufe_count",
    "aidx_num_count",
    "aidx_match_count",
    "audit_detalle_rows",
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
    "score_calidad",
    "total_detectado_calidad",
    "filas_generadas",
    "nuevos",
    "enriquecidas",
    "tiempo_factura_s",
    "tiempo_match_y_registro_s",
}


_TRUE_VALUES = {"1", "true", "yes", "si", "sí", "on"}


def _env_bool(name: str, default: str = "0") -> bool:
    value = str(os.getenv(name, default) or default).strip().lower()
    return value in _TRUE_VALUES


def _today_str() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d")


def _csv_path(audit_dir: str, prefix: str) -> str:
    os.makedirs(audit_dir, exist_ok=True)
    return os.path.join(audit_dir, f"{prefix}_{_today_str()}.csv")


def _to_int(value: Any, default: int = 0) -> int:
    if value is None:
        return default
    if isinstance(value, bool):
        return int(value)
    try:
        text = str(value).strip()
        if not text:
            return default
        return int(float(text.replace(",", ".")))
    except Exception:
        return default


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
    cleaned: dict[str, Any] = {}

    for col in fieldnames:
        value = row.get(col, "")

        # Mantiene columnas numéricas como 0 cuando la métrica no llegó.
        # Esto facilita filtros, cierres diarios y tablas dinámicas.
        if col in _NUMERIC_DEFAULT_ZERO_FIELDS and (value is None or str(value).strip() == ""):
            value = 0

        cleaned[col] = _normalizar_valor_csv(value)

    return cleaned


def _detalle_row_tiene_actividad(row: Dict[str, Any]) -> bool:
    """
    Evita llenar audit_detalle con filas vacías/no útiles.

    Se conserva una fila de detalle cuando:
    - generó filas nuevas,
    - reportó filas_generadas,
    - o representa un error real.
    """
    if not row:
        return False

    if _to_int(row.get("nuevos")) > 0:
        return True

    if _to_int(row.get("filas_generadas")) > 0:
        return True

    estado = str(row.get("estado") or "").strip().lower()
    tipo = str(row.get("tipo_resultado") or "").strip().lower()
    error = str(row.get("error") or "").strip()

    if "error" in estado or tipo == "error":
        return True

    # Errores de descarga, Graph, Excel o SharePoint deben quedar trazados,
    # aunque no hayan generado filas.
    if error and (
        "error" in error.lower()
        or "fall" in error.lower()
        or "exception" in error.lower()
        or "graph" in error.lower()
        or "sharepoint" in error.lower()
        or "workbook" in error.lower()
        or "excel" in error.lower()
        or "descarga" in error.lower()
    ):
        return True

    return False


def _run_summary_tiene_actividad(row: Dict[str, Any]) -> bool:
    """
    Evita crear filas en audit_runs cuando la ejecución no encontró nada nuevo.

    Política definida para producción:
    - Si no hay nuevos y no hay error relevante, no se registra audit_runs.
    - Si hay filas nuevas, errores o detalle real, sí se registra.
    """
    if not row:
        return False

    actividad_numerica = [
        "nuevos_total",
        "filas_local_total",
        "filas_web_total",
        "errores",
        "audit_detalle_rows",
        "facturas_con_filas",
        "facturas_calidad_total",
    ]
    if any(_to_int(row.get(k)) > 0 for k in actividad_numerica):
        return True

    nota = str(row.get("nota") or "").strip().lower()
    if nota and any(x in nota for x in ["error", "fall", "alerta", "warning", "lock", "graph", "sharepoint", "workbook", "excel"]):
        return True

    return False


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
    - Por política de producción, no guarda filas vacías cuando no hubo nuevos
      ni error real.
    """
    if not rows:
        return

    rows_limpias = [dict(r or {}) for r in rows if isinstance(r, dict) and _detalle_row_tiene_actividad(r)]
    if not rows_limpias:
        return

    path = _csv_path(audit_dir, prefix)
    _append_rows_dynamic(
        path=path,
        rows=rows_limpias,
        preferred_columns=DETALLE_BASE_COLUMNS,
    )


def append_run_summary(audit_dir: str, prefix: str, row: Dict[str, Any]) -> None:
    """
    Agrega una fila al audit_runs diario.

    Importante:
    - Soporta columnas nuevas dinámicamente.
    - No descarta métricas nuevas como calidad_promedio_pct,
      facturas_calidad_completa, tiempo_aidx_s, etc.
    - Por política de producción, no guarda corridas vacías cuando no hubo
      facturas nuevas ni error relevante.
    """
    if not row:
        return

    # Permite desactivar el filtro desde .env si algún día se necesita auditar
    # absolutamente todas las corridas, incluso las vacías.
    write_only_if_activity = _env_bool("AUDIT_WRITE_ONLY_IF_ACTIVITY", "1")
    if write_only_if_activity and not _run_summary_tiene_actividad(row):
        return

    path = _csv_path(audit_dir, prefix)
    _append_rows_dynamic(
        path=path,
        rows=[row],
        preferred_columns=RUNS_BASE_COLUMNS,
    )
