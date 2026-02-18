# tests/test_radicados_apply_to_facturas.py
import pandas as pd
from config import ARCHIVO_EXCEL, FACT_COL_NUMERO, FACT_COL_RAD, FACT_COL_PROY
from services.radicados_service import buscar_radicado_y_proyecto

def main():
    print("FACTURAS:", ARCHIVO_EXCEL)
    df = pd.read_excel(ARCHIVO_EXCEL, engine="openpyxl")

    print("Columnas:", list(df.columns))

    # toma las últimas 15 filas para probar rápido
    tail = df.tail(15).copy()

    cambios = 0
    for idx, row in tail.iterrows():
        num = str(row.get(FACT_COL_NUMERO, "")).strip()
        if not num:
            continue

        rad, proy = buscar_radicado_y_proyecto(num)
        if rad or proy:
            before_rad = str(row.get(FACT_COL_RAD, "")).strip()
            before_proy = str(row.get(FACT_COL_PROY, "")).strip()

            if (not before_rad) or (not before_proy):
                df.at[idx, FACT_COL_RAD] = rad
                df.at[idx, FACT_COL_PROY] = proy
                cambios += 1
                print(f"✅ {num} -> RAD={rad} | PROY={proy}")

    if cambios:
        df.to_excel(ARCHIVO_EXCEL, index=False)
        print(f"✅ Guardado facturas.xlsx con {cambios} cambio(s)")
    else:
        print("ℹ️ No hubo cambios (quizá ya estaban llenos o esas filas no están en radicados)")

if __name__ == "__main__":
    main()
