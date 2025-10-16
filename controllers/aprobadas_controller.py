# controllers/aprobadas_controller.py
import os
import io
import re
import base64
import zipfile
import datetime
from pathlib import Path
from typing import List, Dict
from urllib.parse import quote

from utils.fs_utils import borrar_pdfs_en_arbol

from config import (
    DATA_DIR, ARCHIVO_EXCEL, HISTORIAL_EXCEL,
    APROB_FOLDER_NAME, APROB_SEARCH_SINCE_DAYS, MATCH_PRIORIDAD,
    TMP_DIR,  # ⬅️ importamos el alias unificado
)

# Reutilizamos tus servicios existentes
from services.excel_service import guardar_en_excel, registrar_historial_por_zip
from services.factura_service import procesar_xml_en_carpeta
from services.zip_service import extraer_por_zip

# SharePoint (usa versión recursiva corregida)
from services.m365.sp_graph import (
    upload_directory, upload_small_file, ensure_folder, SP_FOLDER as BASE_SP
)

# Microsoft Graph (correo)
from services.m365.mail_graph import (
    get_folder_id_by_name, find_folder_id_anywhere,
    listar_mensajes_en_carpeta, listar_adjuntos_pdf,
    listar_mensajes_zip_inbox, listar_adjuntos_zip,
    descargar_adjunto_por_id, MAILBOX, GRAPH, _SESSION, _h
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
    pdf_cufe = dict_pdf.get("CUFE", "")
    xml_cufe = dict_xml.get("CUFE", "")
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


# -------------------------
# Flujo desde "Aprobadas"
# -------------------------
def run_desde_aprobadas(max_aprobados: int = 50, max_zip_buscar: int = 300):
    """
    1) Lee PDFs desde carpeta de aprobadas.
    2) Extrae identificadores del PDF.
    3) Busca en Inbox el correo con ZIP cuyo XML haga match.
    4) Si hay match, ejecuta tu pipeline normal (XML -> Excel -> SharePoint)
    """

    # Asegurar carpetas locales SOLO aquí (para evitar borrar lo que no es)
    os.makedirs(ADJ_HOY, exist_ok=True)
    os.makedirs(TMP_DIR, exist_ok=True)
    os.makedirs(EXT_HOY, exist_ok=True)

    # 0) Ubicar carpeta de aprobadas
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

    ahora = datetime.datetime.now()
    fecha = ahora.strftime("%Y-%m-%d")
    hora  = ahora.strftime("%H:%M:%S")

    # 1) Por cada mensaje aprobado (PDF)
    for msg in msgs:
        msg_id = msg["id"]
        subj   = msg.get("subject") or ""

        pdf_atts = listar_adjuntos_pdf(msg_id)
        if not pdf_atts:
            continue

        pdf = pdf_atts[0]
        pdf_name = pdf.get("name") or f"{pdf['id']}.pdf"
        pdf_tmp  = os.path.join(TMP_DIR, pdf_name)

        # ✅ Descargar el PDF siempre vía GET /attachments/{id}
        if not descargar_adjunto_por_id(msg_id, pdf["id"], pdf_tmp):
            print(f"[APROB] No pude descargar PDF {pdf_name}")
            continue

        # 2) Identificadores desde el PDF aprobado (con respaldo desde el asunto)
        texto     = extraer_texto_pdf(pdf_tmp)
        ident_pdf = parse_identificadores_pdf(texto)
        if not ident_pdf.get("NUMERO"):
            ident_pdf.setdefault("NUMERO", _numero_from_subject(subj))
        if not ident_pdf.get("FECHA"):
            ident_pdf.setdefault("FECHA", _fecha_from_subject(subj))

        if not (ident_pdf.get("CUFE") or (ident_pdf.get("NUMERO") and ident_pdf.get("FECHA"))):
            print(f"[APROB] PDF {pdf_name}: no hallé CUFE/NUM+FECHA. Saltando.")
            continue

        # 3) Buscar correo original con ZIP y "peek" de XML hasta encontrar match
        print(f"🔎 Buscando ZIP que coincida con {ident_pdf} ...")
        found_match = False
        found_zip_name = None

        inbox_msgs = listar_mensajes_zip_inbox(top=max_zip_buscar)
        for imsg in inbox_msgs:
            zips = listar_adjuntos_zip(imsg["id"])
            if not zips:
                continue

            for z in zips:
                zname   = z.get("name") or f"{z['id']}.zip"

                # ✅ Descargar ZIP por id
                tmp_zip = os.path.join(TMP_DIR, f"peek_{zname}")
                if not descargar_adjunto_por_id(imsg["id"], z["id"], tmp_zip):
                    continue

                with open(tmp_zip, "rb") as f:
                    zip_bytes = f.read()

                # Identificar XMLs dentro
                idents_xml = _peek_ident_xml_from_zip_bytes(zip_bytes)
                os.remove(tmp_zip)

                for ident_xml in idents_xml:
                    if _match_campos(ident_pdf, ident_xml):
                        # ¡MATCH! Guardar ZIP definitivo para pipeline normal
                        found_match    = True
                        found_zip_name = zname
                        zip_local = Path(ADJ_HOY) / zname
                        with open(zip_local, "wb") as f:
                            f.write(zip_bytes)
                        break

                if found_match:
                    break
            if found_match:
                break

        if not found_match:
            print(f"❌ No se encontró ZIP que coincida para PDF {pdf_name}.")
            continue

        # 4) Con el ZIP local, ejecutar flujo normal (solo ese ZIP)
        print(f"🗜️  Extrayendo {found_zip_name} ...")
        resultados = extraer_por_zip(ADJ_HOY, EXT_HOY)  # [(zip_name, carpeta_destino), ...]
        print("🧾 Procesando XMLs...")

        historial_rows = []
        total_nuevos = 0

        for zip_name, carpeta in resultados:
            if zip_name != found_zip_name:
                continue

            ruta = os.path.join(EXT_HOY, carpeta)
            regs, errores_zip = procesar_xml_en_carpeta(ruta)
            nuevos = guardar_en_excel(regs) if regs else 0
            total_nuevos += nuevos

            if nuevos > 0 or errores_zip > 0:
                historial_rows.append({
                    "Fecha": fecha,
                    "Hora":  hora,
                    "Archivo ZIP": zip_name,
                    "Nuevos XML guardados": nuevos,
                    "Errores encontrados": errores_zip
                })

        print(f"✅ Excel local actualizado (+{total_nuevos}): {ARCHIVO_EXCEL}")
        if historial_rows:
            registrar_historial_por_zip(historial_rows)

        # 5) Subida a SharePoint (igual que tu cloud_pipeline)
        print("☁️  Subiendo a SharePoint (desde aprobadas)...")
        if USE_DATE_SUBFOLDERS:
            sp_adj = f"{BASE_SP}/adjuntos/{fecha}"
            sp_ext = f"{BASE_SP}/extraidos/{fecha}"
        else:
            sp_adj = f"{BASE_SP}/adjuntos"
            sp_ext = f"{BASE_SP}/extraidos"
        sp_excel = f"{BASE_SP}/excel"

        ensure_folder(sp_adj); ensure_folder(sp_ext); ensure_folder(sp_excel)

        # Solo el ZIP asociado
        zip_path = os.path.join(ADJ_HOY, found_zip_name)
        upload_small_file(zip_path, f"{sp_adj}/{found_zip_name}", mode="skip")

        # Los extraídos (recursivo) y los excels
        upload_directory(EXT_HOY, sp_ext, mode="skip")
        upload_small_file(ARCHIVO_EXCEL,   f"{sp_excel}/facturas.xlsx",             mode="replace")
        if os.path.exists(HISTORIAL_EXCEL):
            upload_small_file(HISTORIAL_EXCEL, f"{sp_excel}/historial_ejecuciones.xlsx", mode="replace")

        print("🎉 Proceso por aprobadas finalizado para:", found_zip_name)

    # --- Limpieza de PDFs temporales (AL FINAL DEL PROCESO) ---
    try:
        n = borrar_pdfs_en_arbol(TMP_DIR)
        print(f"🧹 Limpieza temp_check: borrados {n} PDF(s).")
    except Exception:
        print("⚠️ Limpieza temp_check: no se pudo completar (continuo).")


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
