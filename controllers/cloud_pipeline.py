# controllers/cloud_pipeline.py
import os
import datetime

from config import DATA_DIR, ARCHIVO_EXCEL, HISTORIAL_EXCEL, TMP_DIR
from utils.fs_utils import borrar_pdfs_en_arbol

from services.m365.mail_graph import descargar_zips_validos
from services.zip_service import extraer_por_zip
from services.factura_service import procesar_xml_en_carpeta
from services.excel_service import guardar_en_excel, registrar_historial_por_zip, obtener_filas_por_archivos

from services.m365.sp_graph import (
    upload_directory,
    upload_small_file,
    ensure_folder,
    SP_FOLDER as BASE_SP
)

from services.m365.excel_workbook_graph import ExcelWorkbookGraph

ADJ_HOY = os.path.join(DATA_DIR, "adjuntos", "hoy")
EXT_HOY = os.path.join(DATA_DIR, "extraidos", "hoy")

USE_DATE_SUBFOLDERS = False
UPLOAD_MODE = "skip"


def run_hibrido(read_all: bool = False, max_messages: int = 200, since_days: int | None = None):
    if since_days is None:
        since_days = 5

    os.makedirs(ADJ_HOY, exist_ok=True)
    os.makedirs(TMP_DIR, exist_ok=True)
    os.makedirs(EXT_HOY, exist_ok=True)

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
        try:
            n = borrar_pdfs_en_arbol(TMP_DIR)
            print(f"🧹 Limpieza temp_check: borrados {n} PDF(s).")
        except Exception:
            print("⚠️ Limpieza temp_check: no se pudo completar (continuo).")
        print("ℹ️ No hay ZIPs válidos nuevos. Fin.")
        return

    print("🗜️  Extrayendo ZIPs por carpeta…")
    resultados = extraer_por_zip(ADJ_HOY, EXT_HOY)

    print("🧾 Procesando XMLs…")
    historial_rows = []
    ahora = datetime.datetime.now()
    fecha = ahora.strftime("%Y-%m-%d")
    hora = ahora.strftime("%H:%M:%S")

    total_nuevos = 0

    # Para Workbook API: acumulamos XML procesados en esta corrida
    archivos_xml_corrida = set()

    for zip_name, carpeta in resultados:
        ruta = os.path.join(EXT_HOY, carpeta)

        done_marker = os.path.join(ruta, ".done")
        if os.path.exists(done_marker):
            continue

        regs, errores_zip = procesar_xml_en_carpeta(ruta)
        nuevos = guardar_en_excel(regs) if regs else 0
        total_nuevos += nuevos

        # registrar nombres xml del folder (más fiable que depender de regs)
        try:
            if os.path.isdir(ruta):
                for fn in os.listdir(ruta):
                    if fn.lower().endswith(".xml"):
                        archivos_xml_corrida.add(fn)
        except Exception:
            pass

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

    # ✅ Workbook API para facturas.xlsx (sin reemplazar)
    print("   ⬆️  facturas.xlsx (Workbook API, sin reemplazar)…")
    if total_nuevos > 0 and archivos_xml_corrida:
        try:
            filas = obtener_filas_por_archivos(archivos_xml_corrida)
            sp_facturas_path = f"{sp_excel}/facturas.xlsx".strip("/")

            xl = ExcelWorkbookGraph(sp_facturas_path)
            insertadas = xl.append_rows_dedup(
                table_name="TblFacturas",
                rows_dicts=filas,
                key_cols=("Archivo", "Concepto"),
            )
            print(f"✅ SharePoint facturas.xlsx actualizado: +{insertadas} fila(s) nuevas.")
        except Exception as e:
            print(f"⚠️ Workbook API falló (no se cae el flujo): {e}")
    else:
        print("ℹ️ No hay nuevos registros o no se detectaron XMLs; no se actualiza tabla en nube.")

    # Historial se sube normal
    if os.path.exists(HISTORIAL_EXCEL):
        upload_small_file(HISTORIAL_EXCEL, f"{sp_excel}/historial_ejecuciones.xlsx", mode="replace")

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

    try:
        n = borrar_pdfs_en_arbol(TMP_DIR)
        print(f"🧹 Limpieza temp_check: borrados {n} PDF(s).")
    except Exception:
        print("⚠️ Limpieza temp_check: no se pudo completar (continuo).")
