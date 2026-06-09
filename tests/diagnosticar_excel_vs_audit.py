import re
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
EXCEL = ROOT / "data" / "facturas.xlsx"
AUDIT_DIR = ROOT / "data" / "audit"


def normalizar_radicado(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def main():
    print("===== DIAGNÓSTICO EXCEL VS AUDIT =====")
    print(f"Excel: {EXCEL}")

    if not EXCEL.exists():
        print("❌ No existe data/facturas.xlsx")
        return

    audits = sorted(AUDIT_DIR.glob("audit_detalle_*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)

    if not audits:
        print("❌ No encontré audit_detalle en data/audit")
        return

    audit_path = audits[0]
    print(f"Audit usado: {audit_path}")

    df = pd.read_excel(EXCEL, sheet_name="Facturas", engine="openpyxl")
    ad = pd.read_csv(audit_path, encoding="utf-8-sig")

    print("\n--- Excel ---")
    print(f"Filas Excel: {len(df)}")
    print(f"Facturas Excel por filas/7: {len(df) / 7}")

    if "Concepto" in df.columns:
        print("\nConceptos en Excel:")
        print(df["Concepto"].value_counts(dropna=False).to_string())

    grupos_excel = (
        df.groupby(["Radicado", "Archivo", "Número de factura"], dropna=False)
        .size()
        .reset_index(name="filas")
    )

    incompletas = grupos_excel[grupos_excel["filas"] != 7]

    print(f"\nGrupos/facturas Excel: {len(grupos_excel)}")
    print(f"Grupos con cantidad distinta de 7: {len(incompletas)}")

    if not incompletas.empty:
        print("\n⚠️ Facturas incompletas:")
        print(incompletas.to_string(index=False))

    print("\n--- Audit ---")
    print(f"Filas audit detalle: {len(ad)}")
    print(f"Suma nuevos audit: {ad['nuevos'].sum() if 'nuevos' in ad.columns else 'SIN COLUMNA'}")
    print(f"Suma filas_generadas audit: {ad['filas_generadas'].sum() if 'filas_generadas' in ad.columns else 'SIN COLUMNA'}")

    ad["RadicadoAudit"] = ad["subject"].astype(str).str.extract(r"Radicado\s+(\d+)", flags=re.IGNORECASE)

    excel_radicados = set(df["Radicado"].apply(normalizar_radicado).astype(str))
    audit_radicados = set(ad["RadicadoAudit"].dropna().astype(str))

    faltan_en_excel = sorted(audit_radicados - excel_radicados)
    sobran_en_excel = sorted(excel_radicados - audit_radicados)

    print("\n--- Comparación Radicados ---")
    print(f"Radicados en audit y NO en Excel: {len(faltan_en_excel)}")
    print(faltan_en_excel[:80])

    print(f"\nRadicados en Excel y NO en audit: {len(sobran_en_excel)}")
    print(sobran_en_excel[:80])

    if faltan_en_excel:
        print("\n--- Detalle de faltantes en audit ---")
        cols = ["RadicadoAudit", "pdf_elegido", "numero", "estado", "zip_match", "nuevos", "fuente"]
        cols = [c for c in cols if c in ad.columns]
        print(ad[ad["RadicadoAudit"].isin(faltan_en_excel)][cols].to_string(index=False))

    print("\n===== FIN DIAGNÓSTICO =====")


if __name__ == "__main__":
    main()