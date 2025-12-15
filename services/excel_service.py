# services/excel_service.py

import os
import pandas as pd
from config import ARCHIVO_EXCEL, HISTORIAL_EXCEL

# Guardado seguro en .xlsx (temporal -> rename atómico)
from utils.safe_io import safe_save_pandas

# Formato de tabla en Excel
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


def obtener_cufes_existentes() -> set:
    """
    Devuelve un set con todos los CUFEs ya registrados en facturas.xlsx.
    Si el archivo no existe o no tiene la columna, devuelve set().

    Se usa para evitar reprocesar facturas ya registradas desde el flujo
    de 'Facturas aprobadas'.
    """
    if not os.path.exists(ARCHIVO_EXCEL):
        return set()

    try:
        df = pd.read_excel(ARCHIVO_EXCEL, engine="openpyxl")
    except Exception as e:
        print(f"[Excel] No se pudo leer facturas.xlsx para índice de CUFEs: {e}")
        return set()

    if "CUFE" not in df.columns:
        return set()

    cufes: set[str] = set()
    for v in df["CUFE"]:
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s:
            cufes.add(s)

    return cufes


def guardar_en_excel(datos):
    """
    Guarda los datos en formato largo:
      - DESCRIPCIÓN = texto de líneas
      - Concepto = (Subtotal, IVA 5%, IVA 19%, etc.)
      - VALOR = valor de cada concepto

    Luego convierte la hoja en una tabla con filtros/estilo.

    ⚠️ Para evitar que el archivo se corrompa, en cada guardado se eliminan
    las tablas existentes de la hoja y se crea una nueva tabla TblFacturas.

    Además:
    - Si existe la columna 'Radicado', se fuerza a que sea la PRIMERA columna.
    - Y se ordenan todas las filas por 'Radicado' ascendente (numérico cuando
      se puede, y los vacíos al final), y luego por 'Número de factura'.
    """
    columnas_fijas = [
        "Archivo", "Empresa emisora", "CUFE",
        "Ciudad emisora", "Código ciudad", "NIT",
        "Cliente", "Número de factura", "Año", "Mes", "Día",
        "Tipo de contribuyente", "Actividad económica",
        "DESCRIPCIÓN", "Concepto", "VALOR"
    ]
    registros_transformados = []

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
            "DESCRIPCIÓN":           d.get("DescripcionLineas", "")
        }
        for medida in [
            "Subtotal", "IVA 5%", "IVA 19%",
            "Retención de IVA", "Retención de ICA",
            "Retención en la fuente", "Total"
        ]:
            fila = base.copy()
            fila["Concepto"] = medida
            fila["VALOR"]   = d.get(medida, 0)
            registros_transformados.append(fila)

    df = pd.DataFrame(registros_transformados, columns=columnas_fijas)
    nuevos = 0

    # 1) Volcado al Excel (crear / actualizar) con guardado seguro
    if os.path.exists(ARCHIVO_EXCEL):
        antiguo = pd.read_excel(ARCHIVO_EXCEL, engine="openpyxl")

        # Unimos manteniendo TODAS las columnas que ya existan (incluyendo Radicado/ProyectoProceso)
        combinado = pd.concat([antiguo, df], ignore_index=True)

        # Eliminamos duplicados por Archivo+Concepto como antes
        combinado = combinado.drop_duplicates(subset=["Archivo", "Concepto"], keep="last")
        nuevos    = len(combinado) - len(antiguo)
        final_df  = combinado
    else:
        nuevos   = len(df)
        final_df = df

    # --- NUEVO: mover Radicado a primera columna y ordenar por Radicado ---
    if "Radicado" in final_df.columns:
        # 1) Reordenar columnas: Radicado primero
        cols = list(final_df.columns)
        cols = ["Radicado"] + [c for c in cols if c != "Radicado"]
        final_df = final_df[cols]

        # 2) Ordenar por Radicado (numérico cuando se puede) y luego por Número de factura
        def _radicado_sort_key(v):
            if pd.isna(v):
                return float("inf")
            s = str(v).strip()
            if not s:
                return float("inf")
            try:
                return int(s)
            except Exception:
                # Si no es numérico, lo mandamos después de los números pero manteniendo orden estable
                return float("inf")

        # Columna auxiliar para ordenar
        final_df["__rad_sort__"] = final_df["Radicado"].apply(_radicado_sort_key)

        # Si no existe la columna de Número de factura, sólo ordenamos por radicado
        sort_cols = ["__rad_sort__"]
        if "Número de factura" in final_df.columns:
            sort_cols.append("Número de factura")

        final_df = final_df.sort_values(sort_cols, kind="mergesort").drop(columns="__rad_sort__")
        final_df = final_df.reset_index(drop=True)

    # Escribe el archivo con temporal .xlsx y rename atómico
    safe_save_pandas(
        final_df,
        ARCHIVO_EXCEL,
        sheet_name="Facturas",
        header=True,
        index=False,
    )

    # 2) Formatear la hoja como tabla de Excel (reconstruyendo la tabla)
    wb = load_workbook(ARCHIVO_EXCEL)
    ws = wb["Facturas"]

    max_row = ws.max_row
    max_col = ws.max_column
    last_col = get_column_letter(max_col)
    table_ref = f"A1:{last_col}{max_row}"

    # Eliminar cualquier tabla existente (evita corrupción)
    if hasattr(ws, "_tables") and ws._tables:
        ws._tables = []

    # Crear nueva tabla con el rango completo
    tbl = Table(displayName="TblFacturas", ref=table_ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tbl)

    # Congelar encabezados
    ws.freeze_panes = "A2"

    wb.save(ARCHIVO_EXCEL)

    print(f"✅ Excel formateado y actualizado: {ARCHIVO_EXCEL}")
    return nuevos


def registrar_historial_por_zip(filas):
    """
    Guarda/actualiza el historial de ejecuciones en otro Excel.
    """
    df_h = pd.DataFrame(filas)
    if os.path.exists(HISTORIAL_EXCEL):
        antiguo = pd.read_excel(HISTORIAL_EXCEL, engine="openpyxl")
        unido   = pd.concat([antiguo, df_h], ignore_index=True)
    else:
        unido = df_h

    # Guardado seguro para el historial también
    safe_save_pandas(
        unido,
        HISTORIAL_EXCEL,
        sheet_name="Historial",
        header=True,
        index=False,
    )

    print(f"📁 Historial actualizado: {HISTORIAL_EXCEL}")
