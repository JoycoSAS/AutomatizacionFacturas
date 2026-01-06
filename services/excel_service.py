# services/excel_service.py

import os
from typing import Any, Dict, List, Set

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from config import ARCHIVO_EXCEL, HISTORIAL_EXCEL
from utils.safe_io import safe_save_pandas


def _read_facturas_df() -> pd.DataFrame:
    """
    Lee facturas.xlsx (hoja 'Facturas' si existe). Si falla, intenta lectura genérica.
    """
    if not os.path.exists(ARCHIVO_EXCEL):
        return pd.DataFrame()

    try:
        # Preferimos hoja "Facturas" porque tu archivo la usa.
        return pd.read_excel(ARCHIVO_EXCEL, sheet_name="Facturas", engine="openpyxl")
    except Exception:
        # Fallback: por si la hoja cambia o el archivo viene sin nombre fijo.
        return pd.read_excel(ARCHIVO_EXCEL, engine="openpyxl")


def _rebuild_table_facturas() -> None:
    """
    Reconstruye la tabla de Excel (TblFacturas) para evitar corrupción, y congela encabezado.
    """
    wb = load_workbook(ARCHIVO_EXCEL)
    ws = wb["Facturas"] if "Facturas" in wb.sheetnames else wb.active

    # Rango completo
    max_row = ws.max_row
    max_col = ws.max_column
    last_col = get_column_letter(max_col)
    table_ref = f"A1:{last_col}{max_row}"

    # Eliminar cualquier tabla existente (evita corrupción)
    if hasattr(ws, "_tables") and ws._tables:
        ws._tables = []

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


def _radicado_sort_key(v: Any) -> int:
    """
    Ordena Radicado como número. Vacíos/no numéricos van al final.
    """
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
    """
    Devuelve un set con todos los CUFEs ya registrados en facturas.xlsx.
    Si el archivo no existe o no tiene la columna, devuelve set().

    Se usa para evitar reprocesar facturas ya registradas desde el flujo
    de 'Facturas aprobadas'.
    """
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


def guardar_en_excel(datos: List[Dict[str, Any]]) -> int:
    """
    Guarda los datos en formato largo:
      - DESCRIPCIÓN = texto de líneas
      - Concepto = (Subtotal, IVA 5%, IVA 19%, etc.)
      - VALOR = valor de cada concepto

    Luego convierte la hoja en una tabla con filtros/estilo.

    ⚠️ Para evitar que el archivo se corrompa, en cada guardado se eliminan
    las tablas existentes de la hoja y se crea una nueva tabla TblFacturas.

    MEJORA:
    - Si existen las columnas 'Radicado' y/o 'ProyectoProceso' (ya sea porque
      el archivo venía así o porque se sincronizan luego), se conservan.
    - Se reordenan al inicio: Radicado, ProyectoProceso, ...
    - Se ordenan todas las filas por 'Radicado' ascendente (numérico cuando
      se puede y vacíos al final) manteniendo un orden estable.
    """
    # Columnas mínimas que generamos desde XML (formato largo)
    columnas_fijas = [
        "Archivo", "Empresa emisora", "CUFE",
        "Ciudad emisora", "Código ciudad", "NIT",
        "Cliente", "Número de factura", "Año", "Mes", "Día",
        "Tipo de contribuyente", "Actividad económica",
        "DESCRIPCIÓN", "Concepto", "VALOR"
    ]

    registros_transformados: List[Dict[str, Any]] = []

    for d in datos:
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
            "DESCRIPCIÓN":           d.get("DescripcionLineas", ""),
        }

        for medida in [
            "Subtotal", "IVA 5%", "IVA 19%",
            "Retención de IVA", "Retención de ICA",
            "Retención en la fuente", "Total"
        ]:
            fila = base.copy()
            fila["Concepto"] = medida
            fila["VALOR"] = d.get(medida, 0)
            registros_transformados.append(fila)

    df_nuevo = pd.DataFrame(registros_transformados, columns=columnas_fijas)
    nuevos = 0

    # 1) Cargar existente y combinar (conservando columnas extra como Radicado/ProyectoProceso)
    if os.path.exists(ARCHIVO_EXCEL):
        antiguo = _read_facturas_df()

        # concat conserva todas las columnas presentes en cualquiera de los dos
        combinado = pd.concat([antiguo, df_nuevo], ignore_index=True)

        # Dedupe igual que antes: Archivo + Concepto
        # (cada factura genera varias filas, una por concepto)
        combinado = combinado.drop_duplicates(subset=["Archivo", "Concepto"], keep="last")

        nuevos = len(combinado) - len(antiguo)
        final_df = combinado
    else:
        nuevos = len(df_nuevo)
        final_df = df_nuevo

    # 2) Reordenar columnas: Radicado, ProyectoProceso primero si existen
    prioridad = [c for c in ["Radicado", "ProyectoProceso"] if c in final_df.columns]
    if prioridad:
        resto = [c for c in final_df.columns if c not in prioridad]
        final_df = final_df[prioridad + resto]

    # 3) Ordenar por Radicado (estable). Si no existe, no tocamos el orden.
    if "Radicado" in final_df.columns:
        final_df["__rad_sort__"] = final_df["Radicado"].apply(_radicado_sort_key)

        # orden estable para no “revolver” las filas dentro del mismo radicado
        sort_cols = ["__rad_sort__"]

        # si existe número de factura, ayuda a estabilidad
        if "Número de factura" in final_df.columns:
            sort_cols.append("Número de factura")

        # si existe Concepto, mantiene el “bloque” (Subtotal/IVA/Ret/Total) consistente
        if "Concepto" in final_df.columns:
            sort_cols.append("Concepto")

        final_df = (
            final_df.sort_values(sort_cols, kind="mergesort")
                    .drop(columns="__rad_sort__")
                    .reset_index(drop=True)
        )

    # 4) Guardar de forma segura (temp -> rename)
    safe_save_pandas(
        final_df,
        ARCHIVO_EXCEL,
        sheet_name="Facturas",
        header=True,
        index=False,
    )

    # 5) Reconstruir tabla (evita corrupción y mantiene filtros)
    _rebuild_table_facturas()

    print(f"✅ Excel formateado y actualizado: {ARCHIVO_EXCEL}")
    return nuevos


def registrar_historial_por_zip(filas: List[Dict[str, Any]]) -> None:
    """
    Guarda/actualiza el historial de ejecuciones en otro Excel.
    """
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
