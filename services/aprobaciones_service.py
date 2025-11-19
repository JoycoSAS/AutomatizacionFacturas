# services/aprobaciones_service.py
"""
Sincroniza datos de Aprobaciones (Radicado, Proyecto) desde el Excel de PA en SharePoint
hacia el Excel local `facturas.xlsx`, cruzando por "Número de factura".
"""

import os
import re
import unicodedata
from openpyxl import load_workbook

from config import (
    ARCHIVO_EXCEL, TMP_DIR,
    APROBACIONES_SP_RELATIVE_PATH, APROBACIONES_SHEET_NAME,
    APROB_COL_NUMERO, APROB_COL_RAD, APROB_COL_PROY,
    FACT_COL_NUMERO, FACT_COL_RAD, FACT_COL_PROY
)
from services.m365.sp_graph import download_small_file


# ---------------------------
# Helpers de normalización
# ---------------------------

def _norm(s: str) -> str:
    """
    Normaliza un número de factura para poder compararlo:
    - quita espacios iniciales/finales
    - elimina acentos
    - deja solo [a-z0-9]
    - pasa a minúsculas
    """
    s = (s or "").strip()
    s = "".join(
        c for c in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(c)
    )
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _extract_numero_from_pa(value: str) -> str:
    """
    En el Excel de PA, la columna NumeroFactura puede venir como:
      '2025-11-12T07:11:4; FEC756'
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

    # 2) Patrón tipo 'FE 94381', 'FE-94381', 'DISL 1595', etc. al final del texto
    #    (letras + espacios/guiones opcionales + dígitos)
    m = re.search(r"([A-Za-z]{1,10}[\s\-]*\d{3,})\s*$", s)
    if m:
        return m.group(1)

    # 3) Fallback: último token alfanumérico (comportamiento anterior)
    m = re.search(r"([A-Za-z0-9\-\/\.]{3,})\s*$", s)
    return m.group(1) if m else s


def _find_col_idx(ws, wanted_names):
    """
    Busca el índice de columna (1-based) en la primera fila de la hoja.
    ws: Worksheet (no tupla). Evita el error de 'tuple has no attribute max_column'.
    """
    wanted = {_norm(x) for x in wanted_names}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=c).value
        if _norm(str(h)) in wanted:
            return c
    return None


def _ensure_column(ws, header_name: str) -> int:
    """
    Devuelve el índice de la columna 'header_name'; si no existe, la crea al final.
    """
    hdrs = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    for i, h in enumerate(hdrs, start=1):
        if _norm(str(h)) == _norm(header_name):
            return i
    new_idx = ws.max_column + 1
    ws.cell(row=1, column=new_idx, value=header_name)
    return new_idx


# ------------------------------------------
# Nuevo: reordenar columnas y ordenar filas
# ------------------------------------------

def _reordenar_y_ordenar_facturas(ws) -> None:
    """
    Reordena las columnas de la hoja 'Facturas' para que queden así:
      1. Radicado
      2. ProyectoProceso
      3. El resto de columnas en el orden original

    Además, ordena las filas por Radicado (numérico si se puede, si no, alfabético).
    """
    # Leer todas las filas (incluyendo encabezado)
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return

    headers = list(rows[0])
    data = list(rows[1:])

    # Construir orden deseado de columnas
    preferidas = ["Radicado", "ProyectoProceso"]
    orden = []

    for col in preferidas:
        if col in headers:
            orden.append(col)

    for h in headers:
        if h not in orden:
            orden.append(h)

    # Mapeo header -> índice original
    idx_map = {h: i for i, h in enumerate(headers)}

    # Reconstruir filas según el nuevo orden
    nuevas_filas = []
    for row in data:
        nueva = []
        for col in orden:
            orig_idx = idx_map.get(col)
            val = row[orig_idx] if orig_idx is not None and orig_idx < len(row) else None
            nueva.append(val)
        nuevas_filas.append(nueva)

    # Ordenar por Radicado (si existe)
    try:
        rad_idx = orden.index("Radicado")
    except ValueError:
        rad_idx = None

    if rad_idx is not None:
        def _key(r):
            v = r[rad_idx]
            if v is None:
                return (2, "")
            s = str(v).strip()
            # Intentar numérico primero
            try:
                return (0, int(s))
            except Exception:
                return (1, s)
        nuevas_filas.sort(key=_key)

    # Limpiar hoja y escribir de nuevo
    ws.delete_rows(1, ws.max_row)
    # Encabezados
    for c, h in enumerate(orden, start=1):
        ws.cell(row=1, column=c, value=h)
    # Datos
    for r_idx, row in enumerate(nuevas_filas, start=2):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)


# -------------------------------------
# Función principal de sincronización
# -------------------------------------

def sincronizar_aprobaciones_en_facturas() -> int:
    """
    Descarga el Excel de Aprobaciones de SP a /data/temp_check,
    cruza por Número de factura y completa 'Radicado' y 'ProyectoProceso'
    en facturas.xlsx (solo si esas celdas están vacías).
    Luego reordena columnas y ordena por Radicado.
    Devuelve la cantidad de filas actualizadas.
    """
    if not os.path.exists(ARCHIVO_EXCEL):
        return 0

    local_pa = os.path.join(TMP_DIR, "Aprobaciones_Facturas.xlsx")
    ok = download_small_file(APROBACIONES_SP_RELATIVE_PATH, local_pa)
    if not ok or not os.path.exists(local_pa):
        print("[APROB] No se pudo descargar el Excel de aprobaciones; se omite cruce.")
        return 0

    # --- Lee aprobaciones: mapa numero->(radicado, proyecto)
    wb_a = load_workbook(local_pa, data_only=True, read_only=True)
    ws_a = wb_a[APROBACIONES_SHEET_NAME] if APROBACIONES_SHEET_NAME in wb_a.sheetnames else wb_a.active

    col_num_a = _find_col_idx(ws_a, [APROB_COL_NUMERO, "numero de factura", "numerofactura"])
    col_rad_a = _find_col_idx(ws_a, [APROB_COL_RAD, "radicado"])
    col_proy_a = _find_col_idx(ws_a, [APROB_COL_PROY, "proyectoproceso", "proyecto/proceso"])

    if not (col_num_a and col_rad_a and col_proy_a):
        print("[APROB] No se localizaron columnas esperadas en Aprobaciones.")
        return 0

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

    # --- Actualiza facturas
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
        clave = _norm(str(num_val)) if num_val is not None else ""
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

    # --- Reordenar columnas y ordenar por Radicado ---
    _reordenar_y_ordenar_facturas(ws_f)

    wb_f.save(ARCHIVO_EXCEL)
    return actualizadas
