# services/excel_service.py
import os
import re
from typing import Any, Dict, List, Set

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

from config import ARCHIVO_EXCEL, HISTORIAL_EXCEL
from utils.safe_io import safe_save_pandas


def _read_facturas_df() -> pd.DataFrame:
    """
    Lee facturas.xlsx (hoja 'Facturas' si existe). Si falla, intenta lectura genérica.
    """
    if not os.path.exists(ARCHIVO_EXCEL):
        return pd.DataFrame()

    try:
        return pd.read_excel(ARCHIVO_EXCEL, sheet_name="Facturas", engine="openpyxl")
    except Exception:
        return pd.read_excel(ARCHIVO_EXCEL, engine="openpyxl")


def _rebuild_table_facturas() -> None:
    """
    Reconstruye la tabla de Excel (TblFacturas) para evitar corrupción,
    congela encabezado y aplica ajustes visuales básicos.
    """
    wb = load_workbook(ARCHIVO_EXCEL)
    ws = wb["Facturas"] if "Facturas" in wb.sheetnames else wb.active

    max_row = ws.max_row
    max_col = ws.max_column
    last_col = get_column_letter(max_col)
    table_ref = f"A1:{last_col}{max_row}"

    # Eliminar tablas existentes
    if hasattr(ws, "_tables") and ws._tables:
        ws._tables = []

    # Crear tabla nueva
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

    # ===== Ajuste visual =====
    DEFAULT_ROW_HEIGHT = 15

    # Fijar altura de filas de datos
    for r in range(2, max_row + 1):
        ws.row_dimensions[r].height = DEFAULT_ROW_HEIGHT

    # Ubicar columna DESCRIPCIÓN por encabezado
    header_map = {}
    for c in range(1, max_col + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None:
            header_map[str(v).strip().upper()] = c

    desc_col = header_map.get("DESCRIPCIÓN")
    if desc_col:
        col_letter = get_column_letter(desc_col)
        ws.column_dimensions[col_letter].width = 45

        align = Alignment(wrap_text=False, vertical="top")
        for r in range(2, max_row + 1):
            ws.cell(row=r, column=desc_col).alignment = align

    wb.save(ARCHIVO_EXCEL)


def _radicado_sort_key(v: Any) -> int:
    if pd.isna(v):
        return 10**12
    s = str(v).strip()
    if not s:
        return 10**12
    try:
        return int(s)
    except Exception:
        return 10**12


def obtener_cufes_existentes() -> Set[str]:
    if not os.path.exists(ARCHIVO_EXCEL):
        return set()

    try:
        df = _read_facturas_df()
    except Exception as e:
        print(f"[Excel] No se pudo leer facturas.xlsx para índice de CUFEs: {e}")
        return set()

    if df.empty or "CUFE" not in df.columns:
        return set()

    cufes: Set[str] = set()
    for v in df["CUFE"]:
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s:
            cufes.add(s)

    return cufes


def _limpiar_descripcion(s: Any) -> str:
    """
    Quita saltos de línea reales que hacen que Excel Online agrande la fila.
    Mantiene el estilo: separado por '; ' en la misma celda.
    """
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    txt = str(s)

    txt = txt.replace("\r\n", "\n").replace("\r", "\n")
    txt = txt.replace("\n", "; ")

    txt = re.sub(r"[ \t]+", " ", txt).strip()
    txt = re.sub(r"(;\s*){2,}", "; ", txt).strip(" ;")

    return txt


def _descripcion_por_concepto(d: Dict[str, Any], concepto: str) -> str:
    """
    ✅ Mejora supermercado/D1:
    - IVA 19%: solo items 19%
    - IVA 5%:  solo items 5%
    - resto: descripción completa
    """
    full = d.get("DescripcionLineas", "") or ""

    if concepto == "IVA 19%":
        return d.get("DescripcionIVA19") or full
    if concepto == "IVA 5%":
        return d.get("DescripcionIVA5") or full

    return full


def guardar_en_excel(datos: List[Dict[str, Any]]) -> int:
    """
    Guarda datos (formato largo): una fila por concepto (Subtotal/IVA/Ret/Total).
    Dedupe: Archivo + Concepto
    Rebuild de tabla: TblFacturas
    """
    columnas_fijas = [
        "Archivo", "Empresa emisora", "CUFE",
        "Ciudad emisora", "Código ciudad", "NIT",
        "Cliente", "Número de factura", "Año", "Mes", "Día",
        "Tipo de contribuyente", "Actividad económica",
        "DESCRIPCIÓN", "Concepto", "VALOR"
    ]

    registros_transformados: List[Dict[str, Any]] = []

    for d in datos:
        for medida in [
            "Subtotal", "IVA 5%", "IVA 19%",
            "Retención de IVA", "Retención de ICA",
            "Retención en la fuente", "Total"
        ]:
            base = {
                "Archivo":               d.get("Archivo", ""),
                "Empresa emisora":       d.get("Empresa emisora", ""),
                "CUFE":                  d.get("CUFE", ""),
                "Ciudad emisora":        d.get("Ciudad emisora", ""),
                "Código ciudad":         d.get("Código ciudad", ""),
                "NIT":                   d.get("NIT", ""),
                "Cliente":               d.get("Cliente", ""),
                "Número de factura":     d.get("Número de factura", ""),
                "Año":                   d.get("Año", ""),
                "Mes":                   d.get("Mes", ""),
                "Día":                   d.get("Día", ""),
                "Tipo de contribuyente": d.get("Tipo de contribuyente", ""),
                "Actividad económica":   d.get("Actividad económica", ""),
                "DESCRIPCIÓN":           _limpiar_descripcion(_descripcion_por_concepto(d, medida)),
            }

            fila = base.copy()
            fila["Concepto"] = medida
            fila["VALOR"] = d.get(medida, 0)
            registros_transformados.append(fila)

    df_nuevo = pd.DataFrame(registros_transformados, columns=columnas_fijas)
    nuevos = 0

    if os.path.exists(ARCHIVO_EXCEL):
        antiguo = _read_facturas_df()
        combinado = pd.concat([antiguo, df_nuevo], ignore_index=True)
        combinado = combinado.drop_duplicates(subset=["Archivo", "Concepto"], keep="last")

        nuevos = len(combinado) - len(antiguo)
        final_df = combinado
    else:
        nuevos = len(df_nuevo)
        final_df = df_nuevo

    # Prioridad de columnas si existen
    prioridad = [c for c in ["Radicado", "ProyectoProceso"] if c in final_df.columns]
    if prioridad:
        resto = [c for c in final_df.columns if c not in prioridad]
        final_df = final_df[prioridad + resto]

    # Orden por radicado si existe
    if "Radicado" in final_df.columns:
        final_df["__rad_sort__"] = final_df["Radicado"].apply(_radicado_sort_key)

        sort_cols = ["__rad_sort__"]
        if "Número de factura" in final_df.columns:
            sort_cols.append("Número de factura")
        if "Concepto" in final_df.columns:
            sort_cols.append("Concepto")

        final_df = (
            final_df.sort_values(sort_cols, kind="mergesort")
                    .drop(columns="__rad_sort__")
                    .reset_index(drop=True)
        )

    safe_save_pandas(
        final_df,
        ARCHIVO_EXCEL,
        sheet_name="Facturas",
        header=True,
        index=False,
    )

    _rebuild_table_facturas()

    print(f"✅ Excel formateado y actualizado: {ARCHIVO_EXCEL}")
    return nuevos


def registrar_historial_por_zip(filas: List[Dict[str, Any]]) -> None:
    df_h = pd.DataFrame(filas)

    if os.path.exists(HISTORIAL_EXCEL):
        try:
            antiguo = pd.read_excel(HISTORIAL_EXCEL, engine="openpyxl")
            unido = pd.concat([antiguo, df_h], ignore_index=True)
        except Exception:
            unido = df_h
    else:
        unido = df_h

    safe_save_pandas(
        unido,
        HISTORIAL_EXCEL,
        sheet_name="Historial",
        header=True,
        index=False,
    )

    print(f"📁 Historial actualizado: {HISTORIAL_EXCEL}")


def obtener_filas_por_archivos(archivos: Set[str]) -> List[Dict[str, Any]]:
    """
    Devuelve filas (dicts) desde facturas.xlsx filtradas por columna 'Archivo'.
    Se usa para mandar SOLO lo nuevo a SharePoint Workbook API.
    """
    if not archivos:
        return []

    df = _read_facturas_df()
    if df.empty or "Archivo" not in df.columns:
        return []

    archivos_norm = {str(a).strip() for a in archivos if str(a).strip()}
    if not archivos_norm:
        return []

    df2 = df[df["Archivo"].astype(str).str.strip().isin(archivos_norm)].copy()
    if df2.empty:
        return []

    df2 = df2.fillna("")
    return df2.to_dict(orient="records")
