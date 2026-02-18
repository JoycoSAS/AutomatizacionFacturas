# services/aprobaciones_service.py
"""
Sincroniza Radicado / ProyectoProceso desde el Excel de Radicados (SharePoint)
hacia facturas.xlsx.

✅ FIX CLAVE:
- Ya NO detecta columnas con heurísticas propias (eso era lo que fallaba y daba:
  "[RAD] Columnas no detectadas correctamente")
- Ahora descarga el Excel de radicados al RADICADOS_LOCAL_PATH (config.py)
  y reutiliza la lógica robusta de:
    services/radicados_service.cargar_mapa_radicados()

Resultado:
- Si existe match por "Número de factura" (normalizado), llena:
    Radicado, ProyectoProceso
- Si no existe match, NO rompe, solo no actualiza.
"""

from __future__ import annotations

import os
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv

from config import (
    ARCHIVO_EXCEL,
    TMP_DIR,
    # Radicados (SharePoint + local estandar)
    RADICADOS_SP_RELATIVE_PATH,
    RADICADOS_LOCAL_PATH,
    # Columnas Facturas.xlsx
    FACT_COL_NUMERO,
    FACT_COL_RAD,
    FACT_COL_PROY,
)

from services.m365.sp_graph import download_small_file
from services.radicados_service import cargar_mapa_radicados
from utils.safe_io import safe_save_pandas
from utils.normalizacion_facturas import normalizar_texto_basico

load_dotenv()
SP_DRIVE_ID_RADICADOS = (os.getenv("SP_DRIVE_ID_RADICADOS") or "").strip()


# -------------------------------------------------
# Helpers (Excel facturas)
# -------------------------------------------------
def _find_col(ws, header_row: int, wanted_name: str) -> int | None:
    wanted = normalizar_texto_basico(wanted_name)
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=header_row, column=c).value
        if normalizar_texto_basico(h) == wanted:
            return c
    return None


def _ensure_column(ws, header_name: str) -> int:
    # busca en la fila 1 (headers)
    col = _find_col(ws, 1, header_name)
    if col:
        return col
    new_c = ws.max_column + 1
    ws.cell(row=1, column=new_c, value=header_name)
    return new_c


def _ordenar_y_formatear_facturas():
    """
    Mantiene el Excel bonito:
    - Radicado y ProyectoProceso al inicio
    - Orden por Radicado (si es numérico)
    - Tabla + freeze panes
    """
    if not os.path.exists(ARCHIVO_EXCEL):
        return

    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Facturas", engine="openpyxl")

    prioridad = [c for c in [FACT_COL_RAD, FACT_COL_PROY] if c in df.columns]
    resto = [c for c in df.columns if c not in prioridad]
    df = df[prioridad + resto]

    if FACT_COL_RAD in df.columns:
        def _rad_sort(x):
            s = str(x).strip()
            return int(s) if s.isdigit() else 10**12

        df["_rad_sort"] = df[FACT_COL_RAD].apply(_rad_sort)
        df = df.sort_values("_rad_sort").drop(columns="_rad_sort").reset_index(drop=True)

    safe_save_pandas(df, ARCHIVO_EXCEL, sheet_name="Facturas")

    wb = load_workbook(ARCHIVO_EXCEL)
    ws = wb["Facturas"]

    # recrear tabla
    ws._tables = []
    ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    tbl = Table(displayName="TblFacturas", ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)

    ws.add_table(tbl)
    ws.freeze_panes = "A2"
    wb.save(ARCHIVO_EXCEL)


# -------------------------------------------------
# FUNCIÓN PRINCIPAL
# -------------------------------------------------
def sincronizar_aprobaciones_en_facturas(force_reload_radicados: bool = True) -> int:
    """
    1) Descarga el Excel de radicados desde SharePoint a RADICADOS_LOCAL_PATH
    2) Construye mapa con cargar_mapa_radicados()
    3) Aplica a facturas.xlsx (solo llena si está vacío)
    """
    if not os.path.exists(ARCHIVO_EXCEL):
        return 0

    if not SP_DRIVE_ID_RADICADOS:
        print("[RAD] Falta SP_DRIVE_ID_RADICADOS en .env")
        return 0

    # Asegura carpetas
    Path(TMP_DIR).mkdir(parents=True, exist_ok=True)
    Path(os.path.dirname(RADICADOS_LOCAL_PATH)).mkdir(parents=True, exist_ok=True)

    # 1) Descargar radicados al path estándar que ya usan los tests
    ok = download_small_file(
        sp_relative_path=RADICADOS_SP_RELATIVE_PATH,
        local_path=RADICADOS_LOCAL_PATH,
        drive_id=SP_DRIVE_ID_RADICADOS,
    )
    if not ok:
        print("[RAD] No se pudo descargar el Excel de radicados.")
        return 0

    # 2) Mapa robusto (usa detección real del header row)
    try:
        mapa = cargar_mapa_radicados(force_reload=force_reload_radicados)
    except Exception as e:
        print(f"[RAD] Error cargando mapa de radicados: {e}")
        return 0

    if not mapa:
        print("[RAD] Mapa de radicados vacío (0).")
        return 0

    # 3) Aplicar a facturas.xlsx
    wb_f = load_workbook(ARCHIVO_EXCEL)
    ws_f = wb_f["Facturas"]

    col_num = _find_col(ws_f, 1, FACT_COL_NUMERO)
    if not col_num:
        print(f"[RAD] No existe columna '{FACT_COL_NUMERO}' en facturas.xlsx")
        return 0

    col_rad_f = _ensure_column(ws_f, FACT_COL_RAD)
    col_proy_f = _ensure_column(ws_f, FACT_COL_PROY)

    updated = 0

    # Importante: la key del mapa ya viene normalizada desde radicados_service
    # (se guarda como _norm_factura(num_raw)). Por eso aquí SOLO buscamos por
    # el valor de la celda "Número de factura" pero usando la misma normalización
    # indirectamente: radicados_service la aplica internamente al consultar.
    #
    # Para no re-importar buscar_radicado_y_proyecto por fila (más lento),
    # usamos el mapa directo con normalización equivalente:
    from services.radicados_service import _norm_factura as _norm_fact

    for r in range(2, ws_f.max_row + 1):
        val = ws_f.cell(r, col_num).value
        if not val:
            continue

        key = _norm_fact(str(val))
        if not key:
            continue

        if key not in mapa:
            continue

        rad, proy = mapa[key]
        changed = False

        if not ws_f.cell(r, col_rad_f).value and rad:
            ws_f.cell(r, col_rad_f, rad)
            changed = True

        if not ws_f.cell(r, col_proy_f).value and proy:
            ws_f.cell(r, col_proy_f, proy)
            changed = True

        if changed:
            updated += 1

    wb_f.save(ARCHIVO_EXCEL)

    if updated > 0:
        _ordenar_y_formatear_facturas()
        print(f"[RAD] ✅ Filas actualizadas en facturas.xlsx: {updated}")
    else:
        print("[RAD] ℹ️ No hubo cambios (sin match o ya estaban llenos).")

    return updated
