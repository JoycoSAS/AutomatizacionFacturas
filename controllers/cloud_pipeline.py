# controllers/cloud_pipeline.py
import os
import datetime

from config import DATA_DIR, ARCHIVO_EXCEL, HISTORIAL_EXCEL, TMP_DIR
from utils.fs_utils import borrar_pdfs_en_arbol

# 1) Correo (Graph) -> descarga a temp y mueve sólo ZIPs con XML
from services.m365.mail_graph import descargar_zips_validos

# 2) ZIP -> extracción (tu servicio existente)
from services.zip_service import extraer_por_zip

# 3) XML -> parseo + Excel local (tus servicios existentes)
from services.factura_service import procesar_xml_en_carpeta
from services.excel_service import guardar_en_excel, registrar_historial_por_zip

# 4) SharePoint (Graph) -> subir directorios/archivos
from services.m365.sp_graph import (
    upload_directory,
    upload_small_file,
    ensure_folder,
    SP_FOLDER as BASE_SP
)

# --- carpetas locales del flujo híbrido ---
ADJ_HOY = os.path.join(DATA_DIR, "adjuntos", "hoy")
EXT_HOY = os.path.join(DATA_DIR, "extraidos", "hoy")

# --- opciones de organización en SharePoint ---
USE_DATE_SUBFOLDERS = False
UPLOAD_MODE = "skip"   # "replace" para sobrescribir; "skip" para no reemplazar


def run_hibrido(read_all: bool = False, max_messages: int = 200, since_days: int | None = None):
    """
    Flujo híbrido:
      1) Lee correo ONLINE (Graph) y descarga a temp_check; sólo mueve ZIPs válidos a adjuntos/hoy
      2) Extrae ZIPs -> extraidos/hoy/<carpeta>
      3) Parsea XMLs -> actualiza Excel local
      4) Sube a SharePoint: ZIPs, extraídos y Excels
      5) Limpia PDFs temporales (temp_check) al final
    """
    # Default de ventana de búsqueda si no te pasan since_days
    if since_days is None:
        since_days = 5  # ← ajustable cuando llames a run_hibrido(...)

    # Asegurar carpetas locales
    os.makedirs(ADJ_HOY, exist_ok=True)
    os.makedirs(TMP_DIR, exist_ok=True)  # /data/temp_check
    os.makedirs(EXT_HOY, exist_ok=True)

    # 1) Descarga selectiva de adjuntos
    print("🔗 Conectando al correo online y descargando ZIPs (peek en temp_check)…")
    zips = descargar_zips_validos(
        temp_check_dir=TMP_DIR,
        destino_dir=ADJ_HOY,
        read_all=read_all,
        max_messages=max_messages,
        since_days=since_days,
    )
    print(f"📥 Descargados {len(zips)} ZIP(s) válidos a {ADJ_HOY}")

    if not zips:
        # Limpieza por si quedaron PDFs de ‘peek’ previos
        try:
            n = borrar_pdfs_en_arbol(TMP_DIR)
            print(f"🧹 Limpieza temp_check: borrados {n} PDF(s).")
        except Exception:
            print("⚠️ Limpieza temp_check: no se pudo completar (continuo).")
        print("ℹ️ No hay ZIPs válidos nuevos. Fin.")
        return

    # 2) Extraer por ZIP (uno por carpeta)
    print("🗜️  Extrayendo ZIPs por carpeta…")
    resultados = extraer_por_zip(ADJ_HOY, EXT_HOY)  # -> [(zip_name, carpeta_destino), …]

    # 3) Procesar XMLs de cada carpeta extraída
    print("🧾 Procesando XMLs…")
    historial_rows = []
    ahora = datetime.datetime.now()
    fecha = ahora.strftime("%Y-%m-%d")
    hora = ahora.strftime("%H:%M:%S")

    total_nuevos = 0
    for zip_name, carpeta in resultados:
        ruta = os.path.join(EXT_HOY, carpeta)

        # ⚠️ Skip si esta carpeta ya fue procesada en una corrida anterior
        done_marker = os.path.join(ruta, ".done")
        if os.path.exists(done_marker):
            continue

        regs, errores_zip = procesar_xml_en_carpeta(ruta)

        nuevos = guardar_en_excel(regs) if regs else 0
        total_nuevos += nuevos

        if nuevos > 0 or errores_zip > 0:
            historial_rows.append({
                "Fecha": fecha,
                "Hora": hora,
                "Archivo ZIP": zip_name,
                "Nuevos XML guardados": nuevos,
                "Errores encontrados": errores_zip,
            })

    print(f"✅ Excel local actualizado ({total_nuevos} registros nuevos): {ARCHIVO_EXCEL}")
    if historial_rows:
        registrar_historial_por_zip(historial_rows)
        print(f"📁 Historial actualizado: {HISTORIAL_EXCEL}")

    # 4) Subir a SharePoint
    print("☁️  Subiendo a SharePoint…")
    print(f"[DEBUG] SP_FOLDER efectivo: {BASE_SP!r}")

    if USE_DATE_SUBFOLDERS:
        sp_adj = f"{BASE_SP}/adjuntos/{fecha}"
        sp_ext = f"{BASE_SP}/extraidos/{fecha}"
    else:
        sp_adj = f"{BASE_SP}/adjuntos"
        sp_ext = f"{BASE_SP}/extraidos"

    sp_excel = f"{BASE_SP}/excel"

    ensure_folder(sp_adj)
    ensure_folder(sp_ext)
    ensure_folder(sp_excel)

    print("   ⬆️  ZIPs…")
    upload_directory(ADJ_HOY, sp_adj, mode=UPLOAD_MODE)

    print("   ⬆️  Extraídos…")
    upload_directory(EXT_HOY, sp_ext, mode=UPLOAD_MODE)

    print("   ⬆️  Excels…")
    upload_small_file(ARCHIVO_EXCEL, f"{sp_excel}/facturas.xlsx", mode="replace")
    if os.path.exists(HISTORIAL_EXCEL):
        upload_small_file(HISTORIAL_EXCEL, f"{sp_excel}/historial_ejecuciones.xlsx", mode="replace")

    # ✅ Crear marcadores .done para todas las carpetas extraídas que participaron en esta corrida
    for zip_name, carpeta in resultados:
        ruta = os.path.join(EXT_HOY, carpeta)
        if not os.path.isdir(ruta):
            continue
        marker = os.path.join(ruta, ".done")
        try:
            with open(marker, "w", encoding="utf-8") as f:
                f.write("ok")
        except Exception:
            pass

    print("🎉 Flujo híbrido finalizado.")

    # 5) Limpieza de PDFs temporales (AL FINAL DEL PROCESO)
    try:
        n = borrar_pdfs_en_arbol(TMP_DIR)
        print(f"🧹 Limpieza temp_check: borrados {n} PDF(s).")
    except Exception:
        print("⚠️ Limpieza temp_check: no se pudo completar (continuo).")
