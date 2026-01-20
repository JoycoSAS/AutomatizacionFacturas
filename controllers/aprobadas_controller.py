# controllers/aprobadas_controller.py
import os
import io
import re
import zipfile
import datetime
import time
from pathlib import Path
from typing import List, Dict, Tuple

from utils.fs_utils import borrar_pdfs_en_arbol

from config import (
    DATA_DIR, ARCHIVO_EXCEL, HISTORIAL_EXCEL,
    APROB_FOLDER_NAME, APROB_SEARCH_SINCE_DAYS, MATCH_PRIORIDAD,
    TMP_DIR,
    AUTO_STOP_MIN_PROCESADOS, AUTO_STOP_SIN_MATCH_CONSEC, AUTO_STOP_SIN_NUEVOS_CONSEC,
)

from services.excel_service import (
    guardar_en_excel,
    registrar_historial_por_zip,
    obtener_cufes_existentes,
    obtener_filas_por_archivos,
)
from services.factura_service import procesar_xml_en_carpeta
from services.zip_service import extraer_por_zip

from services.m365.sp_graph import (
    upload_directory, upload_small_file, ensure_folder, SP_FOLDER as BASE_SP
)

from services.m365.mail_graph import (
    get_folder_id_by_name, find_folder_id_anywhere,
    listar_mensajes_en_carpeta, listar_adjuntos_pdf,
    listar_mensajes_zip_inbox, listar_adjuntos_zip,
    descargar_adjunto_por_id,
    marcar_mensaje_como_leido,
)

from utils.pdf_utils import extraer_texto_pdf, parse_identificadores_pdf, normalizar_fecha
from utils.normalizacion_facturas import claves_normalizadas_factura

from services.aprobaciones_service import sincronizar_aprobaciones_en_facturas

# ✅ Workbook API
from services.m365.excel_workbook_graph import ExcelWorkbookGraph

ADJ_HOY = os.path.join(DATA_DIR, "adjuntos", "hoy")
EXT_HOY = os.path.join(DATA_DIR, "extraidos", "hoy")

USE_DATE_SUBFOLDERS = False
UPLOAD_MODE = "skip"


def __re(pattern: str, text: str):
    import re as _re
    return _re.search(pattern, text, flags=_re.IGNORECASE)


def _norm_cufe(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"[^0-9a-f]", "", s)
    return s


def _parse_ident_from_xml_bytes(xml_bytes: bytes) -> Dict[str, str]:
    text = xml_bytes.decode("utf-8", errors="ignore")
    ident: Dict[str, str] = {}

    m = __re(r"<(?:cbc:|)UUID[^>]*>([^<]{20,})</", text)
    if m:
        ident["CUFE"] = _norm_cufe(m.group(1).strip())

    m = __re(r"<(?:cbc:|)ID[^>]*>([^<]{3,})</", text)
    if m:
        ident["NUMERO"] = m.group(1).strip()

    m = __re(r"<(?:cbc:|)IssueDate[^>]*>([^<]+)</", text)
    if m:
        ident["FECHA"] = normalizar_fecha(m.group(1).strip()) or m.group(1).strip()

    return ident


def _peek_ident_xml_from_zip_bytes(zip_bytes: bytes) -> List[Dict[str, str]]:
    out: List[Dict[str, str]] = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
        for m in zf.infolist():
            if not m.filename.lower().endswith(".xml"):
                continue
            try:
                xml_data = zf.read(m)
                ident = _parse_ident_from_xml_bytes(xml_data)
                ident["xml_name"] = Path(m.filename).name
                out.append(ident)
            except Exception as e:
                print(f"[ZIP] No se pudo leer {m.filename}: {e}")
    return out


def _build_zip_index(
    since_days: int,
    max_zip_buscar: int
) -> Tuple[Dict[str, Tuple[str, bytes]], Dict[Tuple[str, str], Tuple[str, bytes]]]:
    idx_cufe: Dict[str, Tuple[str, bytes]] = {}
    idx_nf: Dict[Tuple[str, str], Tuple[str, bytes]] = {}

    inbox_msgs = listar_mensajes_zip_inbox(top=max_zip_buscar, since_days=since_days)
    limite_utc = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=since_days)

    candidatos = []
    for imsg in inbox_msgs:
        rdt = imsg.get("receivedDateTime")
        if rdt:
            try:
                rdt_dt = datetime.datetime.fromisoformat(rdt.replace("Z", "+00:00"))
                if rdt_dt < limite_utc:
                    continue
            except Exception:
                pass
        candidatos.append(imsg)

    print(f"📦 Prefetch ZIPs: {len(candidatos)} mensajes con adjuntos (ventana {since_days} día(s))")

    for imsg in candidatos:
        zips = listar_adjuntos_zip(imsg["id"])
        if not zips:
            continue

        for z in zips:
            zname = z.get("name") or f"{z['id']}.zip"
            tmp_zip = os.path.join(TMP_DIR, f"prefetch_{zname}")
            if not descargar_adjunto_por_id(imsg["id"], z["id"], tmp_zip):
                continue

            try:
                with open(tmp_zip, "rb") as f:
                    zip_bytes = f.read()
            finally:
                try:
                    os.remove(tmp_zip)
                except Exception:
                    pass

            idents_xml = _peek_ident_xml_from_zip_bytes(zip_bytes)
            for ident_xml in idents_xml:
                cufe = _norm_cufe(ident_xml.get("CUFE") or "")
                num  = (ident_xml.get("NUMERO") or "").strip()
                fec  = (ident_xml.get("FECHA") or "").strip()

                if fec:
                    fec = normalizar_fecha(fec) or fec

                if cufe and cufe not in idx_cufe:
                    idx_cufe[cufe] = (zname, zip_bytes)

                if num and fec:
                    for k in claves_normalizadas_factura(num):
                        key = (k, fec)
                        if key not in idx_nf:
                            idx_nf[key] = (zname, zip_bytes)

    print(f"✅ Índice listo: {len(idx_cufe)} por CUFE, {len(idx_nf)} por NUMERO+FECHA (multi-clave)")
    return idx_cufe, idx_nf


def run_desde_aprobadas(
    max_aprobados: int = 50,
    max_zip_buscar: int = 150,
    since_days: int | None = None
):
    if since_days is None:
        since_days = APROB_SEARCH_SINCE_DAYS

    os.makedirs(ADJ_HOY, exist_ok=True)
    os.makedirs(TMP_DIR, exist_ok=True)
    os.makedirs(EXT_HOY, exist_ok=True)

    folder_id = get_folder_id_by_name("Inbox", APROB_FOLDER_NAME) or find_folder_id_anywhere(APROB_FOLDER_NAME)
    if not folder_id:
        print(f"[APROB] No se encontró la carpeta: {APROB_FOLDER_NAME!r}")
        return

    print(f"📬 Leyendo carpeta de aprobadas (solo NO leídos): {APROB_FOLDER_NAME}")
    msgs = listar_mensajes_en_carpeta(folder_id, top=max_aprobados)
    if not msgs:
        print("ℹ️ No hay mensajes no leídos con aprobaciones recientes.")
        return

    idx_cufe, idx_nf = _build_zip_index(since_days=since_days, max_zip_buscar=max_zip_buscar)

    cufes_existentes = obtener_cufes_existentes()
    print(f"ℹ️ CUFEs ya registrados en facturas.xlsx: {len(cufes_existentes)}")

    fecha_local = datetime.datetime.now().strftime("%Y-%m-%d")
    hora_local  = datetime.datetime.now().strftime("%H:%M:%S")

    t0_total = time.time()
    resumen: List[Tuple[str, float, str, int]] = []

    procesados = 0
    sin_match_consec = 0
    sin_nuevos_consec = 0

    for msg in msgs:
        t0 = time.time()
        msg_id = msg["id"]
        subj   = msg.get("subject") or ""

        pdf_atts = listar_adjuntos_pdf(msg_id)
        if not pdf_atts:
            continue

        pdf = pdf_atts[0]
        pdf_name = pdf.get("name") or f"{pdf['id']}.pdf"
        pdf_tmp  = os.path.join(TMP_DIR, pdf_name)

        if not descargar_adjunto_por_id(msg_id, pdf["id"], pdf_tmp):
            print(f"[APROB] No pude descargar PDF {pdf_name}")
            continue

        texto     = extraer_texto_pdf(pdf_tmp)
        ident_pdf = parse_identificadores_pdf(texto)

        if not ident_pdf.get("NUMERO"):
            ident_pdf.setdefault("NUMERO", _numero_from_subject(subj))
        if not ident_pdf.get("FECHA"):
            ident_pdf.setdefault("FECHA", _fecha_from_subject(subj))

        cufe_pdf = _norm_cufe(ident_pdf.get("CUFE") or "")
        fecha_pdf = (ident_pdf.get("FECHA") or "").strip()
        if fecha_pdf:
            fecha_pdf = normalizar_fecha(fecha_pdf) or fecha_pdf

        if cufe_pdf and cufe_pdf in { _norm_cufe(x) for x in cufes_existentes }:
            print(f"🔁 Factura ya registrada (CUFE en Excel). Se omite búsqueda de ZIP para {pdf_name}.")
            resumen.append((pdf_name, time.time() - t0, "ya registrado", 0))

            sin_match_consec = 0
            sin_nuevos_consec += 1
            procesados += 1

            try:
                marcar_mensaje_como_leido(msg_id)
            except Exception as e:
                print(f"[APROB] No se pudo marcar como leído el mensaje: {e}")

            if (procesados >= AUTO_STOP_MIN_PROCESADOS) and (sin_nuevos_consec >= AUTO_STOP_SIN_NUEVOS_CONSEC):
                print("🛑 Deteniendo flujo: varios PDFs ya registrados/sin nuevos (optimización de tiempo).")
                break
            continue

        found_match = False
        found_zip_name = None
        found_zip_bytes = None

        if cufe_pdf and cufe_pdf in idx_cufe:
            found_zip_name, found_zip_bytes = idx_cufe[cufe_pdf]
            found_match = True
        else:
            num_pdf = (ident_pdf.get("NUMERO") or "").strip()
            if num_pdf and fecha_pdf:
                claves_pdf = claves_normalizadas_factura(num_pdf)
                for k in claves_pdf:
                    key = (k, fecha_pdf)
                    if key in idx_nf:
                        found_zip_name, found_zip_bytes = idx_nf[key]
                        found_match = True
                        break

        if not found_match:
            pdf_base = Path(pdf_name).stem.lower()
            pdf_clean = re.sub(r"[^a-z0-9]", "", pdf_base)

            vistos = set()
            for zn, zbytes in list(idx_cufe.values()) + list(idx_nf.values()):
                if zn in vistos:
                    continue
                vistos.add(zn)

                zbase = Path(zn).stem.lower()
                zclean = re.sub(r"[^a-z0-9]", "", zbase)
                if pdf_clean == zclean or pdf_clean in zclean or zclean in pdf_clean:
                    found_zip_name, found_zip_bytes = zn, zbytes
                    found_match = True
                    print(f"🔄 Emparejado por nombre: {pdf_name} ↔ {zn}")
                    break

        if not found_match or not found_zip_name or not found_zip_bytes:
            print(f"❌ No se encontró ZIP que coincida para PDF {pdf_name}.")
            resumen.append((pdf_name, time.time() - t0, "sin match", 0))

            sin_match_consec += 1
            sin_nuevos_consec = 0
            procesados += 1

            if (procesados >= AUTO_STOP_MIN_PROCESADOS) and (sin_match_consec >= AUTO_STOP_SIN_MATCH_CONSEC):
                print("🛑 Deteniendo flujo: varios PDFs consecutivos sin match (optimización de tiempo).")
                break
            continue

        zip_local_path = Path(ADJ_HOY) / found_zip_name
        with open(zip_local_path, "wb") as f:
            f.write(found_zip_bytes)

        print(f"🗜️  Extrayendo {found_zip_name} ...")
        resultados = extraer_por_zip(ADJ_HOY, EXT_HOY)
        print("🧾 Procesando XMLs...")

        historial_rows = []
        total_nuevos = 0
        carpeta_obj = None
        ruta_obj = None

        for zip_name, carpeta in resultados:
            if zip_name != found_zip_name:
                continue

            ruta = os.path.join(EXT_HOY, carpeta)
            done_marker = os.path.join(ruta, ".done")
            carpeta_obj = carpeta
            ruta_obj = ruta

            if os.path.exists(done_marker):
                continue

            regs, errores_zip = procesar_xml_en_carpeta(ruta)
            nuevos = guardar_en_excel(regs) if regs else 0
            total_nuevos += nuevos

            if nuevos > 0 or errores_zip > 0:
                historial_rows.append({
                    "Fecha": fecha_local,
                    "Hora":  hora_local,
                    "Archivo ZIP": zip_name,
                    "Nuevos XML guardados": nuevos,
                    "Errores encontrados": errores_zip
                })

        print(f"✅ Excel local actualizado (+{total_nuevos}): {ARCHIVO_EXCEL}")
        if historial_rows:
            registrar_historial_por_zip(historial_rows)

        enriquecidas = 0
        try:
            enriquecidas = sincronizar_aprobaciones_en_facturas()
            if enriquecidas > 0:
                print(f"🔗 Enriquecidas {enriquecidas} fila(s) con Radicado/Proyecto desde aprobaciones.")
        except Exception as e:
            print(f"[APROB] Error al sincronizar aprobaciones: {e}")

        print("☁️  Subiendo a SharePoint (desde aprobadas)...")
        if USE_DATE_SUBFOLDERS:
            sp_adj_root = f"{BASE_SP}/adjuntos/{fecha_local}"
            sp_ext_root = f"{BASE_SP}/extraidos/{fecha_local}"
        else:
            sp_adj_root = f"{BASE_SP}/adjuntos"
            sp_ext_root = f"{BASE_SP}/extraidos"
        sp_excel = f"{BASE_SP}/excel"

        ensure_folder(sp_adj_root)
        ensure_folder(sp_ext_root)
        ensure_folder(sp_excel)

        upload_small_file(str(zip_local_path), f"{sp_adj_root}/{found_zip_name}", mode="skip")

        if carpeta_obj and ruta_obj and os.path.exists(ruta_obj):
            upload_directory(ruta_obj, f"{sp_ext_root}/{carpeta_obj}", mode="skip")
        else:
            upload_directory(EXT_HOY, sp_ext_root, mode="skip")

        # ✅ Workbook API: append sin reemplazar el archivo
        hubo_cambios_excel = (total_nuevos > 0) or (enriquecidas > 0)
        if hubo_cambios_excel and sp_excel:
            try:
                # Tomamos nombres XML reales de la carpeta extraída para saber qué filas mandar
                archivos_xml = set()
                if ruta_obj and os.path.isdir(ruta_obj):
                    for fn in os.listdir(ruta_obj):
                        if fn.lower().endswith(".xml"):
                            archivos_xml.add(fn)

                if archivos_xml:
                    filas = obtener_filas_por_archivos(archivos_xml)
                    sp_facturas_path = f"{sp_excel}/facturas.xlsx".strip("/")

                    xl = ExcelWorkbookGraph(sp_facturas_path)
                    insertadas = xl.append_rows_dedup(
                        table_name="TblFacturas",
                        rows_dicts=filas,
                        key_cols=("Archivo", "Concepto"),
                    )
                    print(f"✅ SharePoint facturas.xlsx actualizado (Workbook API): +{insertadas} fila(s) nuevas.")
                else:
                    print("ℹ️ No se detectaron XMLs en la carpeta extraída; no se actualiza tabla en nube.")
            except Exception as e:
                print(f"⚠️ Workbook API falló (no se cae el flujo): {e}")
        else:
            print("ℹ️ Excel sin cambios; no se actualiza facturas.xlsx en nube.")

        # Historial se sube normal (siempre es seguro)
        if historial_rows and os.path.exists(HISTORIAL_EXCEL):
            upload_small_file(HISTORIAL_EXCEL, f"{sp_excel}/historial_ejecuciones.xlsx", mode="replace")

        print("🎉 Proceso por aprobadas finalizado para:", found_zip_name)
        resumen.append((pdf_name, time.time() - t0, "match", total_nuevos))

        try:
            marcar_mensaje_como_leido(msg_id)
        except Exception as e:
            print(f"[APROB] No se pudo marcar como leído el mensaje: {e}")

        sin_match_consec = 0
        if total_nuevos == 0:
            sin_nuevos_consec += 1
        else:
            sin_nuevos_consec = 0
            if cufe_pdf:
                cufes_existentes.add(cufe_pdf)

        procesados += 1
        if (procesados >= AUTO_STOP_MIN_PROCESADOS) and (sin_nuevos_consec >= AUTO_STOP_SIN_NUEVOS_CONSEC):
            print("🛑 Deteniendo flujo: varios PDFs con match pero sin nuevos registros (optimización de tiempo).")
            break

    try:
        n = borrar_pdfs_en_arbol(TMP_DIR)
        print(f"🧹 Limpieza temp_check: borrados {n} PDF(s).")
    except Exception:
        print("⚠️ Limpieza temp_check: no se pudo completar (continuo).")

    total_secs = time.time() - t0_total
    print("\n===== ⏱️ Resumen de tiempos (aprobadas) =====")
    for name, secs, estado, nuevos in resumen:
        print(f"• {name} -> {secs:.2f}s | {estado} | nuevos={nuevos}")
    print(f"⏱️ Tiempo total de ejecución: {total_secs:.2f} s")
    print("=============================================")


def _numero_from_subject(subj: str) -> str | None:
    m = re.search(r"(?:Factura|#|N[o°\.]?)[^\d]*([A-Za-z0-9\-\/\.]{3,})", subj, flags=re.IGNORECASE)
    return m.group(1).strip() if m else None


def _fecha_from_subject(subj: str) -> str | None:
    for pat in [r"(\d{4}[-/]\d{2}[-/]\d{2})", r"(\d{2}[-/]\d{2}[-/]\d{4})"]:
        m = re.search(pat, subj)
        if m:
            s = m.group(1)
            return normalizar_fecha(s) or s
    return None
