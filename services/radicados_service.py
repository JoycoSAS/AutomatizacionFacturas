# services/radicados_service.py
from __future__ import annotations

import os
import re
from typing import Dict, Tuple, Optional

import pandas as pd

from config import (
    RADICADOS_LOCAL_PATH,
    RADICADOS_SHEET_NAME,
    RAD_COL_ASUNTO,
    RAD_COL_RADICADO,
    RAD_COL_PROY,
)

# Cache en memoria para no releer el Excel cada vez
_CACHE_MAP: Optional[Dict[str, Tuple[str, str]]] = None

# =========================
# Normalización
# =========================

def _norm_text(s: str) -> str:
    """Normaliza texto para comparar headers (sin espacios raros, lower)."""
    if s is None:
        return ""
    s = str(s).replace("\u00A0", " ").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def _norm_factura(num: str) -> str:
    """
    Normaliza un número de factura a una clave.
    - FEED6164 -> FEED6164
    - FVE-2080 -> FVE2080
    - FE 266 -> FE266
    - N0761734 -> N0761734 (y se usará también 0761734 como alias)
    """
    if num is None:
        return ""
    s = str(num).upper().strip()
    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    # deja solo A-Z0-9
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s


def _alias_keys(key: str) -> list[str]:
    """
    Devuelve posibles alias de la clave (para casos tipo N0761734 vs 0761734).
    """
    keys = []
    if key:
        keys.append(key)
        # Si es N + dígitos, agrega alias sin la N
        if re.fullmatch(r"N\d{4,}", key):
            keys.append(key[1:])
    # únicos conservando orden
    out = []
    for k in keys:
        if k and k not in out:
            out.append(k)
    return out


# =========================
# Extracción desde Asunto
# =========================

def _extract_numero_factura_from_asunto(asunto: str) -> str:
    """
    Extrae el "número de factura" desde el campo Asunto del Excel de radicados.
    Espera cosas tipo:
      - "Factura: FEED6164 - ..."
      - "Factura: N0761734"
      - "Factura FE 1943 - ..."
      - "FACTURA:F-002-1943 - ..."
    """
    if asunto is None:
        return ""

    s = str(asunto).replace("\u00A0", " ").strip()
    if not s:
        return ""

    # 1) Caso principal: contiene la palabra "Factura"
    m = re.search(r"(?i)\bfactura\b\s*[:\-]?\s*(.+)$", s)
    if not m:
        return ""

    tail = m.group(1).strip()

    # si está separado por " - ", nos quedamos con lo primero
    if " - " in tail:
        tail = tail.split(" - ", 1)[0].strip()

    # Si tiene "Radicado ..." después, cortar
    tail = re.split(r"(?i)\bradicad[oa]\b", tail)[0].strip()

    tail = tail.replace("\u00A0", " ")
    tail = re.sub(r"\s+", " ", tail).strip()
    return tail


# =========================
# Header detection (fila real)
# =========================

def _find_header_row(df_raw: pd.DataFrame, required_headers: list[str], max_scan_rows: int = 120) -> int:
    """
    Busca la fila donde aparecen los headers reales.
    df_raw viene con header=None (valores crudos).
    """
    req = [_norm_text(h) for h in required_headers]
    scan_rows = min(max_scan_rows, len(df_raw))

    for i in range(scan_rows):
        row_vals = [_norm_text(v) for v in df_raw.iloc[i].tolist()]
        if all(r in row_vals for r in req):
            return i

    return -1


# =========================
# API pública
# =========================

def cargar_mapa_radicados(force_reload: bool = False) -> Dict[str, Tuple[str, str]]:
    """
    Lee el Excel de radicados y construye mapa:
      KEY(normalizada) -> (radicado, proyecto)
    También guarda alias para Nxxxx -> xxxx.
    """
    global _CACHE_MAP

    if _CACHE_MAP is not None and not force_reload:
        return _CACHE_MAP

    if not os.path.exists(RADICADOS_LOCAL_PATH):
        raise FileNotFoundError(f"No existe radicados local: {RADICADOS_LOCAL_PATH}")

    required = [RAD_COL_ASUNTO, RAD_COL_RADICADO, RAD_COL_PROY]

    # 1) leer crudo sin header
    df_raw = pd.read_excel(
        RADICADOS_LOCAL_PATH,
        sheet_name=RADICADOS_SHEET_NAME,
        header=None,
        engine="openpyxl",
    )

    # 2) encontrar fila header real
    header_row = _find_header_row(df_raw, required_headers=required, max_scan_rows=120)
    if header_row == -1:
        preview = df_raw.head(20).fillna("").astype(str).values.tolist()
        raise ValueError(
            f"[RAD] No pude detectar fila de headers en '{RADICADOS_SHEET_NAME}'. "
            f"Busqué: {required}. Preview primeras 20 filas: {preview}"
        )

    # 3) construir df ya con headers
    headers = df_raw.iloc[header_row].tolist()
    df = df_raw.iloc[header_row + 1 :].copy()
    df.columns = [str(c).replace("\u00A0", " ").strip() for c in headers]
    df = df.reset_index(drop=True)

    # 4) validar columnas requeridas (por nombre normalizado)
    df_cols_norm = {_norm_text(c): c for c in df.columns}  # norm -> real
    missing = [h for h in required if _norm_text(h) not in df_cols_norm]
    if missing:
        raise ValueError(
            f"[RAD] Columnas no detectadas. Faltan: {missing}. Detectadas: {list(df.columns)}"
        )

    col_asunto = df_cols_norm[_norm_text(RAD_COL_ASUNTO)]
    col_rad = df_cols_norm[_norm_text(RAD_COL_RADICADO)]
    col_proy = df_cols_norm[_norm_text(RAD_COL_PROY)]

    mapa: Dict[str, Tuple[str, str]] = {}

    for _, row in df.iterrows():
        asunto = row.get(col_asunto, "")
        radicado = row.get(col_rad, "")
        proy = row.get(col_proy, "")

        num_raw = _extract_numero_factura_from_asunto(asunto)
        key = _norm_factura(num_raw)
        if not key:
            continue

        rad_str = "" if pd.isna(radicado) else str(radicado).strip()
        proy_str = "" if pd.isna(proy) else str(proy).strip()

        if not (rad_str or proy_str):
            continue

        # guarda key + alias (Nxxxx -> xxxx)
        for k in _alias_keys(key):
            # Si ya existe, no sobre-escribimos a menos que el nuevo tenga info y el viejo esté vacío
            if k not in mapa:
                mapa[k] = (rad_str, proy_str)
            else:
                old_rad, old_proy = mapa[k]
                if (not old_rad and rad_str) or (not old_proy and proy_str):
                    mapa[k] = (rad_str or old_rad, proy_str or old_proy)

    _CACHE_MAP = mapa
    return mapa


def buscar_radicado_y_proyecto(numero_factura: str, force_reload: bool = False) -> Tuple[str, str]:
    """
    Busca (radicado, proyecto) por numero_factura.
    Intenta key normal y alias (Nxxxx -> xxxx) automáticamente.
    """
    mapa = cargar_mapa_radicados(force_reload=force_reload)
    key = _norm_factura(numero_factura)

    for k in _alias_keys(key):
        if k in mapa:
            return mapa[k]

    return ("", "")
