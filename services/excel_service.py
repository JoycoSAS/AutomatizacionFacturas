import os
import pandas as pd
from config import ARCHIVO_EXCEL, HISTORIAL_EXCEL


def guardar_en_excel(datos):
    columnas = [
        "Empresa emisora", "NIT", "Cliente", "Número de factura", "Concepto", "Subtotal",
        "IVA 5%", "IVA 19%", "Retención de IVA", "Retención de ICA", "Total",
        "Año", "Mes", "Día", "Tipo de contribuyente",
        "Responsable de IVA", "Actividad económica", "Archivo"
    ]
    df = pd.DataFrame(datos)[columnas]
    nuevos = 0

    if os.path.exists(ARCHIVO_EXCEL):
        viejo = pd.read_excel(ARCHIVO_EXCEL, engine='openpyxl')
        combo = pd.concat([viejo, df], ignore_index=True)
        combo = combo.drop_duplicates(subset=['Número de factura'], keep='last')
        nuevos = len(combo) - len(viejo)
        final = combo
    else:
        nuevos = len(df)
        final = df

    final.to_excel(ARCHIVO_EXCEL, index=False)
    print(f"✅ Excel actualizado: {ARCHIVO_EXCEL}")
    return nuevos


def registrar_historial_por_zip(filas):
    dfh = pd.DataFrame(filas)
    if os.path.exists(HISTORIAL_EXCEL):
        viejo = pd.read_excel(HISTORIAL_EXCEL, engine='openpyxl')
        out = pd.concat([viejo, dfh], ignore_index=True)
    else:
        out = dfh

    out.to_excel(HISTORIAL_EXCEL, index=False)
    print(f"📑 Historial actualizado: {HISTORIAL_EXCEL}")
