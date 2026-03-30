import os
import math
import pandas as pd

RUTA_AUDIT = r"data/audit/audit_detalle_2026-03-27.csv"
RUTA_EXCEL = r"data/facturas.xlsx"
HOJA = "Facturas"


def norm_text(v):
    if pd.isna(v):
        return ""
    return str(v).strip().lower()


def norm_num(v):
    if pd.isna(v):
        return ""
    s = str(v).strip().lower()
    s = s.replace("–", "-").replace("—", "-").replace("_", "-")
    s = "".join(ch for ch in s if ch.isalnum())
    return s


def cargar_audit(ruta):
    df = pd.read_csv(ruta, dtype=str).fillna("")
    return df


def cargar_excel(ruta_excel, hoja):
    df = pd.read_excel(ruta_excel, sheet_name=hoja, dtype=str).fillna("")
    return df


def construir_indices_excel(df_excel):
    idx_archivo = {}
    idx_numero = {}
    idx_cufe = {}

    for i, row in df_excel.iterrows():
        archivo = norm_text(row.get("Archivo", ""))
        numero = norm_num(row.get("Número de factura", ""))
        cufe = norm_text(row.get("CUFE", ""))

        if archivo:
            idx_archivo.setdefault(archivo, []).append(i)

        if numero:
            idx_numero.setdefault(numero, []).append(i)

        if cufe:
            idx_cufe.setdefault(cufe, []).append(i)

    return idx_archivo, idx_numero, idx_cufe


def buscar_filas_excel(row_audit, df_excel, idx_archivo, idx_numero, idx_cufe):
    candidatos_archivo = set()
    candidatos_numero = set()
    candidatos_cufe = set()

    pdf_elegido = norm_text(row_audit.get("pdf_elegido", ""))
    zip_match = norm_text(row_audit.get("zip_match", ""))
    numero = norm_num(row_audit.get("numero", ""))
    cufe = norm_text(row_audit.get("cufe", ""))

    if pdf_elegido:
        candidatos_archivo.add(pdf_elegido)
    if zip_match and not zip_match.startswith("("):
        candidatos_archivo.add(zip_match)

    filas = set()

    for a in candidatos_archivo:
        filas.update(idx_archivo.get(a, []))

    if not filas and numero:
        filas.update(idx_numero.get(numero, []))

    if not filas and cufe:
        filas.update(idx_cufe.get(cufe, []))

    return sorted(filas)


def main():
    if not os.path.exists(RUTA_AUDIT):
        print(f"No existe audit detalle: {RUTA_AUDIT}")
        return

    if not os.path.exists(RUTA_EXCEL):
        print(f"No existe excel local: {RUTA_EXCEL}")
        return

    df_audit = cargar_audit(RUTA_AUDIT)
    df_excel = cargar_excel(RUTA_EXCEL, HOJA)

    idx_archivo, idx_numero, idx_cufe = construir_indices_excel(df_excel)

    # Estados que sí cuentan como éxito real del sistema
    estados_exitosos = {
        "ok",
        "ok_dian_pdf_only",
        "ok_dian_zip",
        "ok_pdf_aprobadas_fallback",
    }

    df_ok = df_audit[df_audit["estado"].isin(estados_exitosos)].copy()

    print("===== RESUMEN DEBUG ENRIQUECIMIENTO =====")
    print(f"Total filas audit detalle: {len(df_audit)}")
    print(f"Total facturas exitosas: {len(df_ok)}")
    print()

    con_7 = []
    con_otras = []
    sin_filas = []

    for _, row in df_ok.iterrows():
        filas = buscar_filas_excel(row, df_excel, idx_archivo, idx_numero, idx_cufe)
        n = len(filas)

        item = {
            "pdf_elegido": row.get("pdf_elegido", ""),
            "estado": row.get("estado", ""),
            "numero": row.get("numero", ""),
            "cufe": row.get("cufe", ""),
            "zip_match": row.get("zip_match", ""),
            "filas_excel": n,
        }

        if n == 7:
            con_7.append(item)
        elif n == 0:
            sin_filas.append(item)
        else:
            con_otras.append(item)

    print(f"Con 7 filas correctas: {len(con_7)}")
    print(f"Con cantidad distinta de 7: {len(con_otras)}")
    print(f"Sin filas en Excel: {len(sin_filas)}")
    print()

    if con_otras:
        print("===== CANTIDAD DISTINTA DE 7 =====")
        for x in con_otras:
            print(
                f"- {x['pdf_elegido']} | estado={x['estado']} | "
                f"filas_excel={x['filas_excel']} | numero={x['numero']} | zip={x['zip_match']}"
            )
        print()

    if sin_filas:
        print("===== SIN FILAS EN EXCEL =====")
        for x in sin_filas:
            print(
                f"- {x['pdf_elegido']} | estado={x['estado']} | "
                f"numero={x['numero']} | cufe={x['cufe']} | zip={x['zip_match']}"
            )
        print()


if __name__ == "__main__":
    main()