# services/excel_service.py

import os
import pandas as pd
from config import ARCHIVO_EXCEL, HISTORIAL_EXCEL

# Para formatear la tabla en Excel
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


def guardar_en_excel(datos):
    """
    Guarda los datos en “formato largo”: DESCRIPCIÓN = texto de líneas,
    Concepto = tipo de medida (Subtotal, IVA 5%, …), y luego convierte
    la hoja en una tabla de Excel con filtros y estilo.
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

    # 1) Volcar al Excel (sobrescribe o crea)
    if os.path.exists(ARCHIVO_EXCEL):
        antiguo   = pd.read_excel(ARCHIVO_EXCEL, engine='openpyxl')
        combinado = pd.concat([antiguo, df], ignore_index=True)
        combinado = combinado.drop_duplicates(subset=['Archivo', 'Concepto'], keep='last')
        nuevos    = len(combinado) - len(antiguo)
        final     = combinado
    else:
        nuevos = len(df)
        final  = df

    final.to_excel(ARCHIVO_EXCEL, index=False, sheet_name="Facturas")

    # 2) Formatear la hoja como tabla de Excel
    wb = load_workbook(ARCHIVO_EXCEL)
    ws = wb["Facturas"]

    # Determinar rango completo de datos
    max_row = ws.max_row
    max_col = ws.max_column
    last_col = get_column_letter(max_col)
    table_ref = f"A1:{last_col}{max_row}"

    # Crear la tabla
    tabla = Table(displayName="TblFacturas", ref=table_ref)
    # Asignar estilo y filas rayadas
    tabla.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    ws.add_table(tabla)

    # 3) Congelar la primera fila (encabezados)
    ws.freeze_panes = "A2"

    # Guardar cambios
    wb.save(ARCHIVO_EXCEL)

    print(f"✅ Excel formateado y actualizado: {ARCHIVO_EXCEL}")
    return nuevos


def registrar_historial_por_zip(filas):
    """
    Guarda un historial de ejecuciones en otro Excel.
    """
    df_h = pd.DataFrame(filas)
    if os.path.exists(HISTORIAL_EXCEL):
        antiguo = pd.read_excel(HISTORIAL_EXCEL, engine='openpyxl')
        unido   = pd.concat([antiguo, df_h], ignore_index=True)
    else:
        unido = df_h

    unido.to_excel(HISTORIAL_EXCEL, index=False)
    print(f"📁 Historial actualizado: {HISTORIAL_EXCEL}")
