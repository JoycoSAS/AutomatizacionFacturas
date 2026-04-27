# services/excel_service.py
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Set

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from config import ARCHIVO_EXCEL, HISTORIAL_EXCEL
from utils.safe_io import safe_save_pandas


CONCEPTOS_BASE_FIJOS = [
    "Subtotal",
    "IVA 5%",
    "IVA 19%",
    "Retención de IVA",
    "Retención de ICA",
    "Retención en la fuente",
    "Total",
]

COLUMNAS_BASE_LARGO = [
    "Archivo",
    "Empresa emisora",
    "CUFE",
    "Ciudad emisora",
    "Código ciudad",
    "NIT",
    "Cliente",
    "Número de factura",
    "Año",
    "Mes",
    "Día",
    "Tipo de contribuyente",
    "Actividad económica",
    "DESCRIPCIÓN",
    "Concepto",
    "VALOR",
]

COLUMNAS_EXTRA_PERMITIDAS = [
    "Radicado",
    "ProyectoProceso",
]

COLUMNAS_VALIDAS_FINALES = COLUMNAS_EXTRA_PERMITIDAS + COLUMNAS_BASE_LARGO
CONCEPTO_ORDEN = {nombre: idx for idx, nombre in enumerate(CONCEPTOS_BASE_FIJOS, start=1)}


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

    try:
        for table_name in list(ws.tables.keys()):
            del ws.tables[table_name]
    except Exception:
        pass

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

    DEFAULT_ROW_HEIGHT = 15
    for r in range(2, max_row + 1):
        ws.row_dimensions[r].height = DEFAULT_ROW_HEIGHT

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
    wb.close()


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


def _concepto_sort_key(v: Any) -> int:
    s = str(v or "").strip()
    return CONCEPTO_ORDEN.get(s, 999)


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
    Mejora supermercado/D1:
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


def _forzar_texto_numero_factura(valor: Any) -> str:
    """
    Fuerza el número de factura a texto para evitar que Excel/pandas
    lo convierta a notación científica o float.
    """
    if valor is None:
        return ""

    if isinstance(valor, str):
        s = valor.strip()
    else:
        s = str(valor).strip()

    if not s:
        return ""

    s_upper = s.upper().replace(",", ".")

    if "E+" in s_upper or "E-" in s_upper:
        try:
            num = float(s_upper)
            return "{:.0f}".format(num)
        except Exception:
            return s

    if re.fullmatch(r"\d+\.0", s):
        try:
            return str(int(float(s)))
        except Exception:
            return s

    return s


def _normalizar_valor_excel(v: Any) -> Any:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return v


def _limpiar_dataframe_a_formato_largo(df: pd.DataFrame) -> pd.DataFrame:
    """
    Fuerza que el dataframe final SOLO tenga las columnas válidas del formato largo.
    Elimina contaminación previa del Excel local (Subtotal, IVA 19%, etc. como columnas).
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=COLUMNAS_VALIDAS_FINALES)

    work = df.copy()

    for col in COLUMNAS_VALIDAS_FINALES:
        if col not in work.columns:
            work[col] = ""

    work = work[[c for c in COLUMNAS_VALIDAS_FINALES if c in work.columns]].copy()

    if "Número de factura" in work.columns:
        work["Número de factura"] = work["Número de factura"].apply(_forzar_texto_numero_factura)

    if "DESCRIPCIÓN" in work.columns:
        work["DESCRIPCIÓN"] = work["DESCRIPCIÓN"].apply(_limpiar_descripcion)

    if "Concepto" in work.columns:
        work["Concepto"] = work["Concepto"].astype(str).str.strip()

    if "VALOR" in work.columns:
        work["VALOR"] = work["VALOR"].apply(_normalizar_valor_excel)

    for col in COLUMNAS_EXTRA_PERMITIDAS:
        if col in work.columns:
            work[col] = work[col].apply(_normalizar_valor_excel)

    return work


def guardar_en_excel(datos: List[Dict[str, Any]]) -> int:
    """
    Guarda datos SIEMPRE en formato largo:
    una fila por concepto, con Concepto + VALOR.

    FIX PASO A:
    - NO preserva columnas anchas contaminantes como 'Subtotal', 'IVA 19%', etc.
    - limpia el Excel local para que vuelva a quedar en formato largo único.

    MEJORA EXTRA:
    - fuerza orden fijo de conceptos:
      Subtotal, IVA 5%, IVA 19%, Retención de IVA, Retención de ICA,
      Retención en la fuente, Total
    """
    registros_transformados: List[Dict[str, Any]] = []

    for d in (datos or []):
        numero_factura_texto = _forzar_texto_numero_factura(d.get("Número de factura", ""))

        extras = {
            "Radicado": d.get("Radicado", ""),
            "ProyectoProceso": d.get("ProyectoProceso", ""),
        }

        for medida in CONCEPTOS_BASE_FIJOS:
            base = {
                "Archivo":               d.get("Archivo", ""),
                "Empresa emisora":       d.get("Empresa emisora", ""),
                "CUFE":                  d.get("CUFE", ""),
                "Ciudad emisora":        d.get("Ciudad emisora", ""),
                "Código ciudad":         d.get("Código ciudad", ""),
                "NIT":                   d.get("NIT", ""),
                "Cliente":               d.get("Cliente", ""),
                "Número de factura":     numero_factura_texto,
                "Año":                   d.get("Año", ""),
                "Mes":                   d.get("Mes", ""),
                "Día":                   d.get("Día", ""),
                "Tipo de contribuyente": d.get("Tipo de contribuyente", ""),
                "Actividad económica":   d.get("Actividad económica", ""),
                "DESCRIPCIÓN":           _limpiar_descripcion(_descripcion_por_concepto(d, medida)),
                "Concepto":              medida,
                "VALOR":                 d.get(medida, 0),
                **extras,
            }
            registros_transformados.append(base)

    df_nuevo = pd.DataFrame(registros_transformados)
    df_nuevo = _limpiar_dataframe_a_formato_largo(df_nuevo)

    nuevos = 0

    if os.path.exists(ARCHIVO_EXCEL):
        antiguo = _read_facturas_df()
        antiguo = _limpiar_dataframe_a_formato_largo(antiguo)

        combinado = pd.concat([antiguo, df_nuevo], ignore_index=True)

        # DEDUPE CORREGIDO:
        # Antes se usaba solo ["Archivo", "Concepto"], lo cual hacía que facturas distintas
        # con nombres genéricos iguales se pisaran entre sí:
        #   - SIN.img
        #   - Representacion grafica.pdf
        #   - JOYCO S.A.S. BIC.pdf
        #   - Tu Factura ETB Marzo de 2026.pdf
        #
        # Ahora la llave mínima local es:
        #   Radicado + Archivo + Concepto
        #
        # Esto permite que dos facturas distintas con el mismo nombre de archivo
        # convivan correctamente si tienen distinto radicado.
        dedupe_cols = []
        for c in ["Radicado", "Archivo", "Concepto"]:
            if c in combinado.columns:
                dedupe_cols.append(c)

        if len(dedupe_cols) == 3:
            combinado = combinado.drop_duplicates(subset=dedupe_cols, keep="last")
        elif "Archivo" in combinado.columns and "Concepto" in combinado.columns:
            # Fallback conservador por compatibilidad si el Excel antiguo no trae Radicado.
            combinado = combinado.drop_duplicates(subset=["Archivo", "Concepto"], keep="last")

        nuevos = len(combinado) - len(antiguo)
        final_df = combinado
    else:
        nuevos = len(df_nuevo)
        final_df = df_nuevo

    final_df = _limpiar_dataframe_a_formato_largo(final_df)

    # Seguridad local: no conservar filas fantasma sin Concepto.
    # El registro mínimo válido siempre genera los 7 conceptos, así que esto no borra datos buenos.
    if "Concepto" in final_df.columns:
        final_df = final_df[final_df["Concepto"].astype(str).str.strip() != ""].copy()

    columnas_finales_presentes = [c for c in COLUMNAS_VALIDAS_FINALES if c in final_df.columns]
    final_df = final_df[columnas_finales_presentes]

    final_df["__concepto_sort__"] = final_df["Concepto"].apply(_concepto_sort_key)

    sort_cols = []
    if "Radicado" in final_df.columns:
        final_df["__rad_sort__"] = final_df["Radicado"].apply(_radicado_sort_key)
        sort_cols.append("__rad_sort__")

    if "Número de factura" in final_df.columns:
        sort_cols.append("Número de factura")

    sort_cols.append("__concepto_sort__")

    final_df = (
        final_df.sort_values(sort_cols, kind="mergesort")
        .drop(columns=[c for c in ["__rad_sort__", "__concepto_sort__"] if c in final_df.columns])
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


def _normalizar_ref_archivo(s: Any) -> str:
    """
    Normaliza referencias para comparar:
    - basename
    - trim
    - lower
    """
    if s is None:
        return ""
    txt = str(s).strip().replace("\\", "/")
    if not txt:
        return ""
    return os.path.basename(txt).strip().lower()


def _stem_archivo(s: Any) -> str:
    norm = _normalizar_ref_archivo(s)
    if not norm:
        return ""
    return Path(norm).stem.lower().strip()


def obtener_filas_por_archivos(archivos: Set[str]) -> List[Dict[str, Any]]:
    """
    Devuelve filas (dicts) desde facturas.xlsx filtradas por referencias amplias.

    FIX:
    Antes filtraba solo por coincidencia exacta de 'Archivo', lo que hacía perder filas
    en SharePoint cuando local guardaba una referencia distinta (XML/PDF/ZIP o stem).
    """
    if not archivos:
        return []

    df = _read_facturas_df()
    df = _limpiar_dataframe_a_formato_largo(df)

    if df.empty or "Archivo" not in df.columns:
        return []

    archivos_norm = {_normalizar_ref_archivo(a) for a in archivos if _normalizar_ref_archivo(a)}
    if not archivos_norm:
        return []

    stems_ref = {_stem_archivo(a) for a in archivos_norm if _stem_archivo(a)}
    if not stems_ref:
        stems_ref = set()

    df = df.copy()
    df["__archivo_norm__"] = df["Archivo"].apply(_normalizar_ref_archivo)
    df["__archivo_stem__"] = df["Archivo"].apply(_stem_archivo)

    mask_exacta = df["__archivo_norm__"].isin(archivos_norm)
    mask_stem = df["__archivo_stem__"].isin(stems_ref)

    df2 = df[mask_exacta | mask_stem].copy()
    if df2.empty:
        return []

    if "Número de factura" in df2.columns:
        df2["Número de factura"] = df2["Número de factura"].apply(_forzar_texto_numero_factura)

    df2 = df2.drop(columns=["__archivo_norm__", "__archivo_stem__"], errors="ignore")
    df2 = df2.fillna("")
    return df2.to_dict(orient="records")
