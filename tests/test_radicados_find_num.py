import re
import pandas as pd
from config import RADICADOS_LOCAL_PATH, RADICADOS_SHEET_NAME, RAD_COL_ASUNTO, RAD_COL_RADICADO, RAD_COL_PROY

def norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).upper().strip()
    # deja solo letras y números
    return re.sub(r"[^A-Z0-9]+", "", s)

def main():
    objetivo = "N0761734"   # cambia aquí cuando quieras
    obj_norm = norm(objetivo)

    print("LOCAL:", RADICADOS_LOCAL_PATH)
    print("SHEET:", RADICADOS_SHEET_NAME)
    print("OBJETIVO:", objetivo, "->", obj_norm)

    df = pd.read_excel(RADICADOS_LOCAL_PATH, sheet_name=RADICADOS_SHEET_NAME, header=None)

    # Buscar fila real de headers: donde aparezcan "Asunto" y "Consecutivo"
    header_row = None
    for i in range(min(80, len(df))):
        row = [str(x).strip() if x is not None else "" for x in df.iloc[i].tolist()]
        row_norm = [norm(x) for x in row]
        if "ASUNTO" in row_norm and "CONSECUTIVODEENTRADA" in row_norm:
            header_row = i
            break

    if header_row is None:
        print("❌ No encontré fila de headers en las primeras 80 filas.")
        return

    # Releer ya con headers reales
    df2 = pd.read_excel(RADICADOS_LOCAL_PATH, sheet_name=RADICADOS_SHEET_NAME, header=header_row)
    df2.columns = [str(c).strip() for c in df2.columns]

    if RAD_COL_ASUNTO not in df2.columns:
        print("❌ No existe columna Asunto. Columnas:", list(df2.columns))
        return

    encontrados = 0
    for _, r in df2.iterrows():
        asunto = r.get(RAD_COL_ASUNTO)
        if obj_norm in norm(asunto):
            rad = r.get(RAD_COL_RADICADO)
            proy = r.get(RAD_COL_PROY)
            print("✅ MATCH")
            print("ASUNTO:", asunto)
            print("RAD:", rad)
            print("PROY:", proy)
            print("-" * 60)
            encontrados += 1
            if encontrados >= 10:
                break

    print("TOTAL MATCHES:", encontrados)

if __name__ == "__main__":
    main()
