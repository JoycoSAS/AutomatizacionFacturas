# controllers/aprobadas_controller.py
import os
import io
import re
import base64
import zipfile
import datetime
import time
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from urllib.parse import quote

from utils.fs_utils import borrar_pdfs_en_arbol

from config import (
    DATA_DIR, ARCHIVO_EXCEL, HISTORIAL_EXCEL,
    APROB_FOLDER_NAME, APROB_SEARCH_SINCE_DAYS, MATCH_PRIORIDAD,
    TMP_DIR,
)

# Servicios existentes
from services.excel_service import guardar_en_excel, registrar_historial_por_zip
from services.factura_service import procesar_xml_en_carpeta
from services.zip_service import extraer_por_zip

# SharePoint
from services.m365.sp_graph import (
    upload_directory, upload_small_file, ensure_folder, SP_FOLDER as BASE_SP
)

# Microsoft Graph (correo)
from services.m365.mail_graph import (
    get_folder_id_by_name, find_folder_id_anywhere,
    listar_mensajes_en_carpeta, listar_adjuntos_pdf,
    listar_mensajes_zip_inbox, listar_adjuntos_zip,
    descargar_adjunto_por_id
)

# PDF utils
from utils.pdf_utils import extraer_texto_pdf, parse_identificadores_pdf, normalizar_fecha


# Carpetas locales
ADJ_HOY = os.path.join(DATA_DIR, "adjuntos", "hoy")
EXT_HOY = os.path.join(DATA_DIR, "extraidos", "hoy")

USE_DATE_SUBFOLDERS = False
UPLOAD_MODE = "skip"


# ------------------------
# Matching PDF vs XML/ZIP
# ------------------------
def _match_campos(dict_pdf: Dict[str, str], dict_xml: Dict[str, str]) -> bool:
    pr = MATCH_PRIORIDAD or ["CUFE", "NUMERO_FECHA"]
    pdf_cufe = (dict_pdf.get("CUFE") or "").strip()
    xml_cufe = (dict_xml.get("CUFE") or "").strip()
    pdf_num  = (dict_pdf.get("NUMERO") or "").strip()
    xml_num  = (dict_xml.get("NUMERO") or "").strip()
    pdf_fec  = dict_pdf.get("FECHA")
    xml_fec  = dict_xml.get("FECHA")

    for regla in pr:
        if regla == "CUFE":
            if pdf_cufe and xml_cufe and (pdf_cufe == xml_cufe):
                return True
        elif regla == "NUMERO_FECHA":
            if pdf_num and xml_num and pdf_fec and xml_fec:
                if (pdf_num == xml_num) and (pdf_fec == xml_fec):
                    return True
    return False


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


def _parse_ident_from_xml_bytes(xml_bytes: bytes) -> Dict[str, str]:
    text = xml_bytes.decode("utf-8", errors="ignore")
    ident: Dict[str, str] = {}

    m = __re(r"<(?:cbc:|)UUID[^>]*>([^<]{20,})</", text)
    if m:
        ident["CUFE"] = m.group(1).strip()

    m = __re(r"<(?:cbc:|)ID[^>]*>([^<]{3,})</", text)
    if m:
        ident["NUMERO"] = m.group(1).strip()

    m = __re(r"<(?:cbc:|)IssueDate[^>]*>([^<]+)</", text)
    if m:
        ident["FECHA"] = normalizar_fecha(m.group(1).strip()) or m.group(1).strip()

    return ident


def __re(pattern: str, text: str):
    import re as _re
    return _re.search(pattern, text, flags=_re.IGNORECASE)


# ----------------------------------------------------
# Prefetch/Índice de ZIPs: una sola vez por ejecución
# ----------------------------------------------------
def _build_zip_index(
    since_days: int,
    max_zip_buscar: int
) -> Tuple[Dict[str, Tuple[str, bytes]], Dict[Tuple[str, str], Tuple[str, bytes]]]:
    """
    Descarga una sola vez los ZIPs candidatos y construye dos índices:
      - por CUFE
      - por (NUMERO, FECHA)
    """
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
                cufe = (ident_xml.get("CUFE") or "").strip()
                num  = (ident_xml.get("NUMERO") or "").strip()
                fec  = (ident_xml.get("FECHA") or "").strip()

                if cufe and cufe not in idx_cufe:
                    idx_cufe[cufe] = (zname, zip_bytes)

                if num and fec and (num, fec) not in idx_nf:
                    idx_nf[(num, fec)] = (zname, zip_bytes)

    print(f"✅ Índice listo: {len(idx_cufe)} por CUFE, {len(idx_nf)} por NUMERO+FECHA")
    return idx_cufe, idx_nf


# -------------------------
# Flujo desde "Aprobadas"
# -------------------------
def run_desde_aprobadas(max_aprobados: int = 50, max_zip_buscar: int = 150, since_days: int | None = None):
    """
    Flujo principal: busca coincidencias por CUFE o NUMERO+FECHA,
    y si no hay match, usa fallback por nombre del archivo PDF/ZIP.
    """
    if since_days is None:
        since_days = APROB_SEARCH_SINCE_DAYS

    os.makedirs(ADJ_HOY, exist_ok=True)
    os.makedirs(TMP_DIR, exist_ok=True)
    os.makedirs(EXT_HOY, exist_ok=True)

    folder_id = get_folder_id_by_name("Inbox", APROB_FOLDER_NAME) or \
                find_folder_id_anywhere(APROB_FOLDER_NAME)
    if not folder_id:
        print(f"[APROB] No se encontró la carpeta: {APROB_FOLDER_NAME!r}")
        return

    print(f"📬 Leyendo carpeta de aprobadas: {APROB_FOLDER_NAME}")
    msgs = listar_mensajes_en_carpeta(folder_id, top=max_aprobados)
    if not msgs:
        print("ℹ️ No hay mensajes con aprobaciones recientes.")
        return

    idx_cufe, idx_nf = _build_zip_index(since_days=since_days, max_zip_buscar=max_zip_buscar)
    fecha_local = datetime.datetime.now().strftime("%Y-%m-%d")
    hora_local  = datetime.datetime.now().strftime("%H:%M:%S")

    t0_total = time.time()
    resumen: List[Tuple[str, float, str, int]] = []

    # -------- LOOP PRINCIPAL --------
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

        found_match = False
        found_zip_name = None
        found_zip_bytes = None

        # --- A) Por CUFE ---
        cufe = (ident_pdf.get("CUFE") or "").strip()
        if cufe and cufe in idx_cufe:
            found_zip_name, found_zip_bytes = idx_cufe[cufe]
            found_match = True
        else:
            # --- B) Por NUMERO+FECHA ---
            num = (ident_pdf.get("NUMERO") or "").strip()
            fec = (ident_pdf.get("FECHA") or "").strip()
            if num and fec and (num, fec) in idx_nf:
                found_zip_name, found_zip_bytes = idx_nf[(num, fec)]
                found_match = True

        # --- C) Fallback por nombre del archivo (corregido) ---
        if not found_match:
            pdf_base = Path(pdf_name).stem.lower()
            pdf_clean = re.sub(r"[^a-z0-9]", "", pdf_base)
            for zn, zbytes in list(idx_cufe.values()) + list(idx_nf.values()):
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
            continue

        # --- Guardar ZIP local ---
        zip_local_path = Path(ADJ_HOY) / found_zip_name
        with open(zip_local_path, "wb") as f:
            f.write(found_zip_bytes)

        # --- Procesamiento normal ---
        print(f"🗜️  Extrayendo {found_zip_name} ...")
        resultados = extraer_por_zip(ADJ_HOY, EXT_HOY)
        print("🧾 Procesando XMLs...")

        historial_rows = []
        total_nuevos = 0
        for zip_name, carpeta in resultados:
            if zip_name != found_zip_name:
                continue
            ruta = os.path.join(EXT_HOY, carpeta)
            done_marker = os.path.join(ruta, ".done")
            if os.path.exists(done_marker):
                continue

            regs, errores_zip = procesar_xml_en_carpeta(ruta)
            nuevos = guardar_en_excel(regs) if regs else 0
            total_nuevos += nuevos

            if nuevos > 0 or errores_zip > 0:
                historial_rows.append({
                    "Fecha": fecha_local,
                    "Hora": hora_local,
                    "Archivo ZIP": zip_name,
                    "Nuevos XML guardados": nuevos,
                    "Errores encontrados": errores_zip
                })

        print(f"✅ Excel local actualizado (+{total_nuevos}): {ARCHIVO_EXCEL}")
        if historial_rows:
            registrar_historial_por_zip(historial_rows)

        # --- Subida a SharePoint ---
        print("☁️  Subiendo a SharePoint (desde aprobadas)...")
        if USE_DATE_SUBFOLDERS:
            sp_adj = f"{BASE_SP}/adjuntos/{fecha_local}"
            sp_ext = f"{BASE_SP}/extraidos/{fecha_local}"
        else:
            sp_adj = f"{BASE_SP}/adjuntos"
            sp_ext = f"{BASE_SP}/extraidos"
        sp_excel = f"{BASE_SP}/excel"

        ensure_folder(sp_adj); ensure_folder(sp_ext); ensure_folder(sp_excel)
        upload_small_file(str(zip_local_path), f"{sp_adj}/{found_zip_name}", mode="skip")
        upload_directory(EXT_HOY, sp_ext, mode="skip")
        upload_small_file(ARCHIVO_EXCEL, f"{sp_excel}/facturas.xlsx", mode="replace")
        if os.path.exists(HISTORIAL_EXCEL):
            upload_small_file(HISTORIAL_EXCEL, f"{sp_excel}/historial_ejecuciones.xlsx", mode="replace")

        for zip_name, carpeta in resultados:
            if zip_name != found_zip_name:
                continue
            ruta = os.path.join(EXT_HOY, carpeta)
            with open(os.path.join(ruta, ".done"), "w", encoding="utf-8") as f:
                f.write("ok")

        print("🎉 Proceso por aprobadas finalizado para:", found_zip_name)
        resumen.append((pdf_name, time.time() - t0, "match", total_nuevos))

    # --- Limpieza final ---
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


# --------------------
# Helpers desde asunto
# --------------------
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
