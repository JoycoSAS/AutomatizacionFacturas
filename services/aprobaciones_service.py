# services/aprobaciones_service.py
"""
Sincroniza datos de Aprobaciones (Radicado, Proyecto) desde el Excel de PA en SharePoint
hacia el Excel local `facturas.xlsx`, cruzando por "Número de factura".

Además:
- Asegura que las columnas 'Radicado' y 'ProyectoProceso' queden como primeras columnas.
- Ordena todas las filas por 'Radicado' (numérico ascendente).
- Reconstruye la tabla de Excel (TblFacturas) para evitar archivos dañados.
"""

import os
import re
import unicodedata

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


# ---------------------------
# Helpers de normalización
# ---------------------------

def _norm(s) -> str:
    """
    Normaliza un texto (número de factura, encabezado, etc.) para compararlo:
    - convierte a str
    - quita espacios iniciales/finales
    - elimina acentos
    - deja solo [a-z0-9]
    - pasa a minúsculas
    """
    if s is None:
        s = ""
    s = str(s).strip()
    s = "".join(
        c for c in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(c)
    )
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _extract_numero_from_pa(value: str) -> str:
    """
    En el Excel de PA, la columna NumeroFactura puede venir como:
      '2025-11-12T07:11:4; FE-21394'
      '2025-11-19T15:24:; FE 94381'
      'FEC756'
      'FE94381'
      etc.

    La idea es devolver el número de factura completo (incluyendo prefijo)
    para que, al normalizarlo con _norm, coincida con el de facturas.xlsx.
    """
    if value is None:
        return ""
    s = str(value).strip()

    # 1) Caso explícito 'Factura: XXXX'
    m = re.search(r"Factura:\s*([A-Za-z0-9\-\/\.]{3,})", s, flags=re.IGNORECASE)
    if m:
        return m.group(1)

    # 2) Patrón tipo 'FE 94381', 'FE-94381', 'DISL 1595' al final del texto
    m = re.search(r"([A-Za-z]{1,10}[\s\-]*\d{3,})\s*$", s)
    if m:
        return m.group(1)

    # 3) Fallback: último token alfanumérico
    m = re.search(r"([A-Za-z0-9\-\/\.]{3,})\s*$", s)
    return m.group(1) if m else s


def _flatten_names(names):
    """
    Acepta una lista que puede contener strings o listas/tuplas de strings
    y devuelve un generador plano de nombres.
    """
    for n in names:
        if isinstance(n, (list, tuple, set)):
            for sub in n:
                yield sub
        else:
            yield n


def _find_col_idx(ws, wanted_names):
    """
    Busca el índice de columna (1-based) en la primera fila de la hoja.
    - `wanted_names` puede contener strings o listas de alias.
    Usa _norm para comparar (tolerante a espacios, mayúsculas, acentos, etc.).
    """
    wanted = {_norm(x) for x in _flatten_names(wanted_names)}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=c).value
        if _norm(h) in wanted:
            return c
    return None


def _ensure_column(ws, header_name: str) -> int:
    """
    Devuelve el índice de la columna 'header_name'; si no existe, la crea al final.
    Usa _norm para comparar encabezados.
    """
    hdrs = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    for i, h in enumerate(hdrs, start=1):
        if _norm(h) == _norm(header_name):
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
    - Ordena por 'Radicado' numérico ascendente
    - Guarda con safe_save_pandas
    - Reconstruye la tabla TblFacturas en la hoja 'Facturas'
    """
    if not os.path.exists(ARCHIVO_EXCEL):
        return

    # Leer con pandas
    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Facturas", engine="openpyxl")

    cols = list(df.columns)

    # 1) Reordenar columnas (Radicado, ProyectoProceso primero)
    prioridad = []
    for c in ["Radicado", "ProyectoProceso"]:
        if c in cols:
            prioridad.append(c)

    resto = [c for c in cols if c not in prioridad]
    if prioridad:
        df = df[prioridad + resto]

    # 2) Ordenar por Radicado de forma numérica
    if "Radicado" in df.columns:
        def _rad_to_int(x):
            if pd.isna(x):
                return 10**12  # los vacíos al final
            try:
                return int(str(x).strip())
            except Exception:
                return 10**12

        df = df.sort_values(
            by="Radicado",
            key=lambda s: s.map(_rad_to_int),
            kind="mergesort"   # estable: mantiene el orden de conceptos por factura
        ).reset_index(drop=True)

    # 3) Guardar de forma segura
    safe_save_pandas(
        df,
        ARCHIVO_EXCEL,
        sheet_name="Facturas",
        header=True,
        index=False,
    )

    # 4) Reconstruir la tabla TblFacturas
    wb = load_workbook(ARCHIVO_EXCEL)
    ws = wb["Facturas"]

    # Eliminar cualquier tabla previa para evitar archivos corruptos
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
    Descarga el Excel de Aprobaciones de SP a /data/temp_check,
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

    # 2) Leer aprobaciones (PA) con openpyxl
    wb_a = load_workbook(local_pa, data_only=True, read_only=True)
    ws_a = wb_a[APROBACIONES_SHEET_NAME] if APROBACIONES_SHEET_NAME in wb_a.sheetnames else wb_a.active

    col_num_a = _find_col_idx(ws_a, [APROB_COL_NUMERO, "numero de factura", "numerofactura"])
    col_rad_a = _find_col_idx(ws_a, [APROB_COL_RAD, "radicado"])
    col_proy_a = _find_col_idx(ws_a, [APROB_COL_PROY, "proyectoproceso", "proyecto/proceso"])

    if not (col_num_a and col_rad_a and col_proy_a):
        print("[APROB] No se localizaron columnas esperadas en Aprobaciones.")
        return 0

    # mapa: numero_normalizado -> (radicado, proyecto)
    mapa = {}
    for r in ws_a.iter_rows(min_row=2, values_only=True):
        raw = r[col_num_a - 1]
        if raw is None:
            continue
        num = _extract_numero_from_pa(str(raw))
        clave = _norm(num)
        if not clave:
            continue
        mapa[clave] = (r[col_rad_a - 1], r[col_proy_a - 1])

    # 3) Actualizar facturas.xlsx con openpyxl (solo completa celdas vacías)
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
        clave = _norm(num_val) if num_val is not None else ""
        if not clave:
            continue

        if clave in mapa:
            rad, proy = mapa[clave]
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

    # 4) Reordenar columnas, ordenar por Radicado y reconstruir tabla con pandas
    _ordenar_y_formatear_por_radicado()

    return actualizadas
