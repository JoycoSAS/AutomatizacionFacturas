import pandas as pd

ruta = "data/audit/audit_detalle_2026-03-25.csv"

df = pd.read_csv(ruta)

# última corrida
run_id = df["run_id"].iloc[-1]
df_run = df[df["run_id"] == run_id]

# =========================
# SIN MATCH
# =========================
sin_match = df_run[df_run["estado"].str.contains("sin_match", na=False)]

print("\n===== SIN MATCH =====")
for _, row in sin_match.iterrows():
    print(f"- {row['pdf_elegido']} | estado={row['estado']}")

# =========================
# SIN PDF
# =========================
sin_pdf = df_run[df_run["estado"] == "sin_pdf"]

print("\n===== SIN PDF =====")
for _, row in sin_pdf.iterrows():
    print(f"- msg_id={row['msg_id']} | asunto={row['subject']}")