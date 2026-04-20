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
    """Normaliza texto para comparar headers."""
    if s is None:
        return ""
    s = str(s).replace("\u00A0", " ").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def _norm_factura(num: str) -> str:
    """
    Normaliza número de factura para usarlo como clave de match.
    Conserva letras y números, elimina espacios/guiones/símbolos.
    """
    if num is None:
        return ""

    s = str(num).strip().upper()
    if not s:
        return ""

    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()

    # reparar posible notación científica si llega dañada
    sci = s.replace(",", ".")
    if "E+" in sci or "E-" in sci:
        try:
            s = "{:.0f}".format(float(sci))
        except Exception:
            pass

    s = re.sub(r"[^A-Z0-9]", "", s)
    return s


def _build_possible_keys(num: str) -> list[str]:
    """
    Genera varias claves posibles para aumentar la probabilidad de match.
    """
    raw = str(num or "").strip().upper()
    if not raw:
        return []

    keys = []
    seen = set()

    def add(x: str):
        x = str(x or "").strip().upper()
        if not x:
            return
        if x not in seen:
            seen.add(x)
            keys.append(x)

    add(raw)
    add(raw.replace(" ", ""))
    add(raw.replace("-", ""))
    add(raw.replace("_", ""))
    add(re.sub(r"[^A-Z0-9]", "", raw))

    limpio = re.sub(r"[^A-Z0-9]", "", raw)

    # alias N0761734 -> 0761734
    if re.fullmatch(r"N\d{4,}", limpio):
        add(limpio[1:])

    # FE-1245 -> FE1245
    m = re.match(r"^([A-Z]+)(\d+)$", limpio)
    if m:
        pref, dig = m.groups()
        add(f"{pref}{dig}")
        add(f"{pref}-{dig}")
        add(f"{pref} {dig}")

    return keys


def _alias_keys(key: str) -> list[str]:
    """
    Devuelve posibles alias de una clave ya normalizada.
    """
    out = []
    seen = set()

    def add(x: str):
        x = str(x or "").strip().upper()
        if x and x not in seen:
            seen.add(x)
            out.append(x)

    add(key)

    if re.fullmatch(r"N\d{4,}", key):
        add(key[1:])

    return out


# =========================
# Extracción desde Asunto
# =========================

def _extract_numero_factura_from_asunto(asunto: str) -> str:
    """
    Extrae el número de factura desde el campo Asunto del Excel de radicados.
    """
    if asunto is None:
        return ""

    s = str(asunto).replace("\u00A0", " ").strip()
    if not s:
        return ""

    m = re.search(r"(?i)\bfactura\b\s*[:\-]?\s*(.+)$", s)
    if not m:
        return ""

    tail = m.group(1).strip()

    if " - " in tail:
        tail = tail.split(" - ", 1)[0].strip()

    tail = re.split(r"(?i)\bradicad[oa]\b", tail)[0].strip()
    tail = tail.replace("\u00A0", " ")
    tail = re.sub(r"\s+", " ", tail).strip()

    return tail


# =========================
# Header detection
# =========================

def _find_header_row(df_raw: pd.DataFrame, required_headers: list[str], max_scan_rows: int = 120) -> int:
    """
    Busca la fila donde aparecen los headers reales.
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
    """
    global _CACHE_MAP

    if _CACHE_MAP is not None and not force_reload:
        return _CACHE_MAP

    if not os.path.exists(RADICADOS_LOCAL_PATH):
        raise FileNotFoundError(f"No existe radicados local: {RADICADOS_LOCAL_PATH}")

    required = [RAD_COL_ASUNTO, RAD_COL_RADICADO, RAD_COL_PROY]

    df_raw = pd.read_excel(
        RADICADOS_LOCAL_PATH,
        sheet_name=RADICADOS_SHEET_NAME,
        header=None,
        engine="openpyxl",
    )

    header_row = _find_header_row(df_raw, required_headers=required, max_scan_rows=120)
    if header_row == -1:
        preview = df_raw.head(20).fillna("").astype(str).values.tolist()
        raise ValueError(
            f"[RAD] No pude detectar fila de headers en '{RADICADOS_SHEET_NAME}'. "
            f"Busqué: {required}. Preview primeras 20 filas: {preview}"
        )

    headers = df_raw.iloc[header_row].tolist()
    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).replace("\u00A0", " ").strip() for c in headers]
    df = df.reset_index(drop=True)

    df_cols_norm = {_norm_text(c): c for c in df.columns}
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
        if not num_raw:
            continue

        rad_str = "" if pd.isna(radicado) else str(radicado).strip()
        proy_str = "" if pd.isna(proy) else str(proy).strip()

        if not (rad_str or proy_str):
            continue

        for k in _build_possible_keys(num_raw):
            key_norm = _norm_factura(k)
            if not key_norm:
                continue

            if key_norm not in mapa:
                mapa[key_norm] = (rad_str, proy_str)
            else:
                old_rad, old_proy = mapa[key_norm]
                if (not old_rad and rad_str) or (not old_proy and proy_str):
                    mapa[key_norm] = (rad_str or old_rad, proy_str or old_proy)

        # alias adicionales para claves ya limpias
        key_main = _norm_factura(num_raw)
        for alias in _alias_keys(key_main):
            if alias and alias not in mapa:
                mapa[alias] = (rad_str, proy_str)

    _CACHE_MAP = mapa
    return mapa


def buscar_radicado_y_proyecto(numero_factura: str, force_reload: bool = False) -> Tuple[str, str]:
    """
    Busca (radicado, proyecto) por numero_factura.
    Intenta varios alias automáticamente.
    """
    mapa = cargar_mapa_radicados(force_reload=force_reload)

    posibles = _build_possible_keys(numero_factura)
    for p in posibles:
        key = _norm_factura(p)
        if key in mapa:
            return mapa[key]

    key = _norm_factura(numero_factura)
    for alias in _alias_keys(key):
        if alias in mapa:
            return mapa[alias]

    # match flexible por contención
    key_clean = _norm_factura(numero_factura)
    if key_clean:
        for mk, val in mapa.items():
            if key_clean == mk:
                return val
            if len(key_clean) >= 5 and (key_clean in mk or mk in key_clean):
                return val

    return ("", "")