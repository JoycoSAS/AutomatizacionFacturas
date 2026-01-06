# services/aprobaciones_service.py
"""
Sincroniza datos de Aprobaciones (Radicado, Proyecto) desde el Excel de PA en SharePoint
hacia el Excel local `facturas.xlsx`, cruzando por "Número de factura".

Mejoras incluidas:
- Normalización fuerte y multi-candidatos del "Número de factura" (tolerante a espacios/guiones/formato).
- Asegura que las columnas 'Radicado' y 'ProyectoProceso' queden como primeras columnas.
- Ordena todas las filas por 'Radicado' (numérico ascendente).
- Reconstruye la tabla de Excel (TblFacturas) para evitar archivos dañados.
"""

import os
import pandas as pd

from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from config import (
    ARCHIVO_EXCEL, TMP_DIR,
    APROBACIONES_SP_RELATIVE_PATH, APROBACIONES_SHEET_NAME,
    APROB_COL_NUMERO, APROB_COL_RAD, APROB_COL_PROY,
    FACT_COL_NUMERO, FACT_COL_RAD, FACT_COL_PROY
)
from services.m365.sp_graph import download_small_file
from utils.safe_io import safe_save_pandas
from utils.normalizacion_facturas import (
    normalizar_texto_basico,
    claves_normalizadas_factura,
)


# ---------------------------
# Helpers columnas (encabezados)
# ---------------------------

def _flatten_names(names):
    for n in names:
        if isinstance(n, (list, tuple, set)):
            for sub in n:
                yield sub
        else:
            yield n


def _find_col_idx(ws, wanted_names):
    """
    Busca el índice de columna (1-based) en la primera fila de la hoja.
    Usa normalizar_texto_basico para comparar.
    """
    wanted = {normalizar_texto_basico(x) for x in _flatten_names(wanted_names)}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=c).value
        if normalizar_texto_basico(h) in wanted:
            return c
    return None


def _ensure_column(ws, header_name: str) -> int:
    """
    Devuelve el índice de la columna 'header_name'; si no existe, la crea al final.
    """
    hdrs = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    for i, h in enumerate(hdrs, start=1):
        if normalizar_texto_basico(h) == normalizar_texto_basico(header_name):
            return i
    new_idx = ws.max_column + 1
    ws.cell(row=1, column=new_idx, value=header_name)
    return new_idx


# ------------------------------------------
# Post-proceso: ordenar y formatear con pandas
# ------------------------------------------

def _ordenar_y_formatear_por_radicado() -> None:
    """
    - Lee facturas.xlsx con pandas
    - Mueve 'Radicado' y 'ProyectoProceso' al inicio (si existen)
    - Ordena por 'Radicado' numérico ascendente (vacíos al final)
    - Guarda con safe_save_pandas
    - Reconstruye la tabla TblFacturas en la hoja 'Facturas'
    """
    if not os.path.exists(ARCHIVO_EXCEL):
        return

    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Facturas", engine="openpyxl")

    cols = list(df.columns)

    prioridad = [c for c in ["Radicado", "ProyectoProceso"] if c in cols]
    resto = [c for c in cols if c not in prioridad]
    if prioridad:
        df = df[prioridad + resto]

    if "Radicado" in df.columns:
        def _rad_to_int(x):
            if pd.isna(x):
                return 10**12
            s = str(x).strip()
            if not s:
                return 10**12
            try:
                return int(s)
            except Exception:
                return 10**12

        df = (
            df.sort_values(
                by="Radicado",
                key=lambda s: s.map(_rad_to_int),
                kind="mergesort"   # estable
            )
            .reset_index(drop=True)
        )

    safe_save_pandas(
        df,
        ARCHIVO_EXCEL,
        sheet_name="Facturas",
        header=True,
        index=False,
    )

    wb = load_workbook(ARCHIVO_EXCEL)
    ws = wb["Facturas"]

    if hasattr(ws, "_tables") and ws._tables:
        ws._tables = []

    max_row = ws.max_row
    max_col = ws.max_column
    last_col = get_column_letter(max_col)
    table_ref = f"A1:{last_col}{max_row}"

    tbl = Table(displayName="TblFacturas", ref=table_ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tbl)
    ws.freeze_panes = "A2"

    wb.save(ARCHIVO_EXCEL)


# -------------------------------------
# Función principal de sincronización
# -------------------------------------

def sincronizar_aprobaciones_en_facturas() -> int:
    """
    Descarga el Excel de Aprobaciones de SP a TMP_DIR,
    cruza por Número de factura y completa 'Radicado' y 'ProyectoProceso'
    en facturas.xlsx (solo si esas celdas están vacías).

    Luego:
      - Reordena columnas para que Radicado/ProyectoProceso sean las primeras.
      - Ordena todas las filas por Radicado (numérico).
      - Reconstruye la tabla TblFacturas.

    Devuelve la cantidad de filas actualizadas.
    """
    if not os.path.exists(ARCHIVO_EXCEL):
        return 0

    # 1) Descargar Excel de Aprobaciones
    local_pa = os.path.join(TMP_DIR, "Aprobaciones_Facturas.xlsx")
    ok = download_small_file(APROBACIONES_SP_RELATIVE_PATH, local_pa)
    if not ok or not os.path.exists(local_pa):
        print("[APROB] No se pudo descargar el Excel de aprobaciones; se omite cruce.")
        return 0

    # 2) Leer aprobaciones (PA)
    wb_a = load_workbook(local_pa, data_only=True, read_only=True)
    ws_a = wb_a[APROBACIONES_SHEET_NAME] if APROBACIONES_SHEET_NAME in wb_a.sheetnames else wb_a.active

    col_num_a = _find_col_idx(ws_a, [APROB_COL_NUMERO, "numero de factura", "numerofactura"])
    col_rad_a = _find_col_idx(ws_a, [APROB_COL_RAD, "radicado"])
    col_proy_a = _find_col_idx(ws_a, [APROB_COL_PROY, "proyectoproceso", "proyecto/proceso", "proyecto proceso"])

    if not (col_num_a and col_rad_a and col_proy_a):
        print("[APROB] No se localizaron columnas esperadas en Aprobaciones.")
        return 0

    # 2.1) Construir mapa: clave_normalizada -> (radicado, proyecto)
    #      OJO: generamos varias claves por fila (tolerancia máxima)
    mapa = {}
    for r in ws_a.iter_rows(min_row=2, values_only=True):
        raw_num = r[col_num_a - 1]
        if raw_num is None:
            continue

        rad = r[col_rad_a - 1]
        proy = r[col_proy_a - 1]

        # Generar múltiples claves normalizadas
        for k in claves_normalizadas_factura(str(raw_num)):
            # nos quedamos con el último (si PA repite factura, gana el último)
            mapa[k] = (rad, proy)

    # 3) Actualizar facturas.xlsx
    wb_f = load_workbook(ARCHIVO_EXCEL)
    ws_f = wb_f["Facturas"] if "Facturas" in wb_f.sheetnames else wb_f.active

    col_num_f = _find_col_idx(ws_f, [FACT_COL_NUMERO, "numero de factura", "numerofactura"])
    if not col_num_f:
        print("[APROB] No se encontró la columna de Número de factura en facturas.xlsx.")
        return 0

    col_rad_f = _ensure_column(ws_f, FACT_COL_RAD)
    col_proy_f = _ensure_column(ws_f, FACT_COL_PROY)

    actualizadas = 0

    for row in range(2, ws_f.max_row + 1):
        num_val = ws_f.cell(row=row, column=col_num_f).value
        if num_val is None:
            continue

        # Claves candidatas del valor local (facturas.xlsx)
        claves_locales = claves_normalizadas_factura(str(num_val))
        if not claves_locales:
            continue

        hit = None
        for k in claves_locales:
            if k in mapa:
                hit = mapa[k]
                break

        if not hit:
            continue

        rad, proy = hit
        wrote = False

        if ws_f.cell(row=row, column=col_rad_f).value in (None, "") and rad not in (None, ""):
            ws_f.cell(row=row, column=col_rad_f, value=rad)
            wrote = True

        if ws_f.cell(row=row, column=col_proy_f).value in (None, "") and proy not in (None, ""):
            ws_f.cell(row=row, column=col_proy_f, value=proy)
            wrote = True

        if wrote:
            actualizadas += 1

    wb_f.save(ARCHIVO_EXCEL)

    # 4) Reordenar columnas, ordenar por Radicado y reconstruir tabla
    _ordenar_y_formatear_por_radicado()

    return actualizadas
