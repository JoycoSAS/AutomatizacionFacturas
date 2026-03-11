# controllers/aprobadas_controller.py
import os
import io
import re
import zipfile
import datetime
import time
import shutil
import uuid
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import xml.etree.ElementTree as ET

from utils.fs_utils import borrar_pdfs_en_arbol
from utils.processed_store import ProcessedStore
from utils.text_normalizer import normalize_text
from utils.attachment_index_store import AttachmentIndexStore

from utils.audit_csv_logger import append_run_summary, append_detalle_rows
from utils.single_instance_lock import SingleInstanceLock

from config import (
    DATA_DIR, ARCHIVO_EXCEL, HISTORIAL_EXCEL,
    APROB_FOLDER_NAME, APROB_SEARCH_SINCE_DAYS,
    TMP_DIR,
    AUTO_STOP_MIN_PROCESADOS, AUTO_STOP_SIN_MATCH_CONSEC, AUTO_STOP_SIN_NUEVOS_CONSEC,
    PROCESSED_MESSAGES_PATH, PROCESSED_MESSAGES_TTL_DAYS,
    ATTACHMENT_INDEX_PATH, ATTACHMENT_INDEX_TTL_DAYS,

    # DIAN
    APROB_DIAN_KEYWORD,
    INBOX_DIAN_SUBJECT_CANDIDATES,
    REQUIRE_DIAN_IN_BODY_PREVIEW,

    # AUDIT
    AUDIT_DIR, AUDIT_RUNS_PREFIX, AUDIT_DETALLE_PREFIX, AUDIT_WRITE_ONLY_IF_ACTIVITY,

    # LOCK
    LOCK_FILE_APROBADAS, LOCK_TTL_SECONDS,
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
    buscar_mensajes_inbox_por_asunto,
)

from utils.pdf_utils import (
    extraer_texto_pdf,
    parse_identificadores_pdf,
    normalizar_fecha,
    extraer_totales_basicos_pdf,
    extraer_campos_basicos_pdf
)

from utils.normalizacion_facturas import claves_normalizadas_factura
from services.aprobaciones_service import sincronizar_aprobaciones_en_facturas
from services.m365.excel_workbook_graph import ExcelWorkbookGraph


ADJ_HOY = os.path.join(DATA_DIR, "adjuntos", "hoy")
EXT_HOY = os.path.join(DATA_DIR, "extraidos", "hoy")

USE_DATE_SUBFOLDERS = False


# ============================================================
# ✅ Helper: auditoría detalle (una fila por mensaje)
# ============================================================
def _push_detalle(
    detalle_rows: list,
    run_id: str,
    msg_id: str,
    subj: str,
    pdf_name: str = "",
    cufe: str = "",
    numero: str = "",
    fecha_factura: str = "",
    zip_match: str = "",
    estado: str = "",
    duracion_s: float = 0.0,
    nuevos: int = 0,
    enriquecidas: int = 0,
    fuente: str = "",
    error: str = ""
):
    detalle_rows.append({
        "run_id": run_id,
        "fecha_hora": datetime.datetime.now().isoformat(timespec="seconds"),
        "msg_id": msg_id,
        "subject": subj or "",
        "pdf_elegido": pdf_name or "",
        "cufe": cufe or "",
        "numero": numero or "",
        "fecha_factura": fecha_factura or "",
        "zip_match": zip_match or "",
        "estado": estado or "",
        "duracion_s": round(float(duracion_s or 0.0), 3),
        "nuevos": int(nuevos or 0),
        "enriquecidas": int(enriquecidas or 0),
        "fuente": fuente or "",
        "error": (error or "")[:500],
    })


def _limpiar_adj_hoy() -> int:
    borrados = 0
    try:
        if not os.path.isdir(ADJ_HOY):
            return 0
        for fn in os.listdir(ADJ_HOY):
            if fn.lower().endswith(".zip"):
                try:
                    os.remove(os.path.join(ADJ_HOY, fn))
                    borrados += 1
                except Exception:
                    pass
    except Exception:
        pass
    return borrados


def _limpiar_ext_hoy() -> int:
    borrados = 0
    try:
        if not os.path.isdir(EXT_HOY):
            return 0

        for name in os.listdir(EXT_HOY):
            p = os.path.join(EXT_HOY, name)
            try:
                if os.path.isdir(p):
                    shutil.rmtree(p, ignore_errors=True)
                else:
                    os.remove(p)
                borrados += 1
            except Exception:
                pass
    except Exception:
        pass
    return borrados


def __re(pattern: str, text: str):
    import re as _re
    return _re.search(pattern, text, flags=_re.IGNORECASE | _re.DOTALL)


def _norm_cufe(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"[^0-9a-f]", "", s)
    return s


def _cufe_is_valid(cufe: str) -> bool:
    c = _norm_cufe(cufe or "")
    return bool(c) and len(c) >= 40


def _is_acta_filename(name: str) -> bool:
    s = (name or "").lower()
    s_clean = re.sub(r"\s+", " ", s)
    bad_keys = [
        "acta", "constancia", "certificado", "aprobacion", "aprobación",
        "memorando", "oficio", "radicado", "soporte de radicacion", "soporte de radicación",
        "documento", "comunicado"
    ]
    return any(k in s_clean for k in bad_keys)


# ============================================================
# Detectar DIAN en asunto + bodyPreview (si existe)
# ============================================================
def _contains_dian(text: str) -> bool:
    return normalize_text(APROB_DIAN_KEYWORD) in normalize_text(text or "")


def _is_dian_trigger_message(msg: dict) -> bool:
    subj = msg.get("subject") or ""
    if not _contains_dian(subj):
        return False

    preview = (msg.get("bodyPreview") or msg.get("body_preview") or "")
    if REQUIRE_DIAN_IN_BODY_PREVIEW:
        if preview:
            return _contains_dian(preview)
        else:
            print("[DIAN] ⚠️ Mensaje no trae bodyPreview; se valida solo por asunto.")
            return True

    return True


_CTRL_REGEX = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
_AMP_FIX = re.compile(r"&(?!(?:[a-zA-Z]+|#\d+|#x[0-9A-Fa-f]+);)")


def _clean_xml_text(txt: str) -> str:
    txt = _CTRL_REGEX.sub("", txt)
    txt = _AMP_FIX.sub("&amp;", txt)
    return txt


def _extract_inner_invoice_text(xml_text: str) -> str | None:
    if not xml_text:
        return None

    m = re.search(
        r'(<\s*(?:Invoice|CreditNote|DebitNote)\b.*?</\s*(?:Invoice|CreditNote|DebitNote)\s*>)',
        xml_text,
        flags=re.IGNORECASE | re.DOTALL
    )
    if m:
        inner = m.group(1)
        return _clean_xml_text(inner)

    return None


def _parse_ident_from_xml_bytes(xml_bytes: bytes) -> Dict[str, str]:
    ident: Dict[str, str] = {}

    try:
        text = xml_bytes.decode("utf-8-sig", errors="replace")
    except Exception:
        text = xml_bytes.decode("utf-8", errors="ignore")

    text = _clean_xml_text(text)

    inner = _extract_inner_invoice_text(text)
    if inner:
        try:
            r = ET.fromstring(inner)
            id_el = r.find("./{*}ID")
            uuid_el = r.find(".//{*}UUID")
            issue_el = r.find("./{*}IssueDate")

            if uuid_el is not None and uuid_el.text:
                ident["CUFE"] = _norm_cufe(uuid_el.text.strip())

            if id_el is not None and id_el.text:
                ident["NUMERO"] = id_el.text.strip()

            if issue_el is not None and issue_el.text:
                ident["FECHA"] = normalizar_fecha(issue_el.text.strip()) or issue_el.text.strip()

            return ident
        except Exception:
            pass

    try:
        root = ET.fromstring(text)
    except Exception:
        m = __re(r"<(?:cbc:|)UUID[^>]*>([^<]{20,})</", text)
        if m:
            ident["CUFE"] = _norm_cufe(m.group(1).strip())
        m = __re(r"<(?:cbc:|)IssueDate[^>]*>([^<]+)</", text)
        if m:
            ident["FECHA"] = normalizar_fecha(m.group(1).strip()) or m.group(1).strip()
        m = __re(r"<(?:cbc:|)ParentDocumentID[^>]*>([^<]{3,})</", text)
        if m:
            ident["NUMERO"] = m.group(1).strip()
        return ident

    local = root.tag.split("}")[-1] if "}" in root.tag else root.tag

    if local.lower() == "attacheddocument":
        pd = root.find(".//{*}ParentDocumentID")
        if pd is not None and pd.text:
            ident["NUMERO"] = pd.text.strip()

        uuid_el = root.find(".//{*}UUID")
        if uuid_el is not None and uuid_el.text:
            ident["CUFE"] = _norm_cufe(uuid_el.text.strip())

        issue_el = root.find(".//{*}IssueDate")
        if issue_el is not None and issue_el.text:
            ident["FECHA"] = normalizar_fecha(issue_el.text.strip()) or issue_el.text.strip()

        return ident

    id_el = root.find("./{*}ID")
    uuid_el = root.find(".//{*}UUID")
    issue_el = root.find("./{*}IssueDate")

    if uuid_el is not None and uuid_el.text:
        ident["CUFE"] = _norm_cufe(uuid_el.text.strip())
    if id_el is not None and id_el.text:
        ident["NUMERO"] = id_el.text.strip()
    if issue_el is not None and issue_el.text:
        ident["FECHA"] = normalizar_fecha(issue_el.text.strip()) or issue_el.text.strip()

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
    max_zip_buscar: int,
    aidx: AttachmentIndexStore
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
        mid = imsg.get("id")
        if not mid:
            continue

        zips = listar_adjuntos_zip(mid)
        if not zips:
            continue

        for z in zips:
            zid = z.get("id")
            if not zid:
                continue

            zname = z.get("name") or f"{zid}.zip"
            tmp_zip = os.path.join(TMP_DIR, f"prefetch_{zname}")
            if not descargar_adjunto_por_id(mid, zid, tmp_zip):
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
                num = (ident_xml.get("NUMERO") or "").strip()
                fec = (ident_xml.get("FECHA") or "").strip()
                if fec:
                    fec = normalizar_fecha(fec) or fec

                try:
                    aidx.upsert_zip(
                        cufe=cufe,
                        numero=num,
                        fecha=fec,
                        msg_id=mid,
                        att_id=zid,
                        att_name=zname,
                        received_dt_iso=imsg.get("receivedDateTime", "") or "",
                    )
                except Exception as e:
                    print(f"[AIDX] No pude upsert ZIP index: {e}")

                if cufe and cufe not in idx_cufe:
                    idx_cufe[cufe] = (zname, zip_bytes)

                if num and fec:
                    for k in claves_normalizadas_factura(num):
                        key = (k, fec)
                        if key not in idx_nf:
                            idx_nf[key] = (zname, zip_bytes)

    print(f"✅ Índice listo: {len(idx_cufe)} por CUFE, {len(idx_nf)} por NUMERO+FECHA (multi-clave)")
    return idx_cufe, idx_nf


_NON_INVOICE_PREFIXES = {"DDI", "RAD", "RDI", "RDC", "REC", "RCP", "DOC", "REF"}


def _is_generic_or_non_invoice_numero(n: str) -> bool:
    n = (n or "").strip().upper()
    if not n:
        return True

    m = re.match(r"^([A-Z]{1,10})[-–—]?\s*(\d{3,})$", n)
    if m:
        pref = m.group(1).upper()
        if pref in _NON_INVOICE_PREFIXES:
            return True

    if re.match(r"^[A-Z]{1,4}-\d{1,3}$", n):
        return True

    return False


def _prefer_subject_numero(pdf_num: str | None, subj_num: str | None) -> str | None:
    pdf_num = (pdf_num or "").strip()
    subj_num = (subj_num or "").strip()
    if not subj_num:
        return pdf_num or None
    if not pdf_num:
        return subj_num

    if _is_generic_or_non_invoice_numero(pdf_num) and not _is_generic_or_non_invoice_numero(subj_num):
        return subj_num

    if len(re.findall(r"\d", subj_num)) > len(re.findall(r"\d", pdf_num)):
        return subj_num

    return pdf_num


def _clean_name(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def _match_pdf_candidate(
    target_ident: Dict[str, str],
    target_pdf_name: str,
    candidate_ident: Dict[str, str],
    candidate_name: str
) -> bool:
    t_cufe = _norm_cufe(target_ident.get("CUFE") or "")
    t_num = (target_ident.get("NUMERO") or "").strip()
    t_ap = (target_ident.get("NUMERO_APROB") or "").strip()

    c_cufe = _norm_cufe(candidate_ident.get("CUFE") or "")
    c_num = (candidate_ident.get("NUMERO") or "").strip()
    c_ap = (candidate_ident.get("NUMERO_APROB") or "").strip()

    if t_cufe and c_cufe and t_cufe == c_cufe:
        return True

    nums_target = {x for x in [t_num, t_ap] if x}
    nums_cand = {x for x in [c_num, c_ap] if x}
    if nums_target and nums_cand and (nums_target & nums_cand):
        return True

    if _clean_name(Path(target_pdf_name).stem) == _clean_name(Path(candidate_name).stem):
        return True

    base = (candidate_name or "").upper()
    for n in nums_target:
        if n and n.upper() in base:
            return True

    return False


# ============================================================
# buscar correo contenedor "Validación(s) DIAN"
# ============================================================
def _is_validacion_dian_subject(subj: str) -> bool:
    s = normalize_text(subj or "")
    for cand in INBOX_DIAN_SUBJECT_CANDIDATES:
        if normalize_text(cand) in s:
            return True
    return False


def _buscar_pdf_en_correo_validaciones_dian(
    target_ident: Dict[str, str],
    target_pdf_name: str,
    since_days: int,
    top_msgs: int = 80
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    os.makedirs(os.path.join(TMP_DIR, "dian_pdf_only"), exist_ok=True)

    msgs = buscar_mensajes_inbox_por_asunto(
        asunto_contiene="DIAN",
        top=top_msgs,
        since_days=since_days
    ) or []

    if not msgs:
        print("[DIAN] Graph search no devolvió nada por 'DIAN'. Fallback: inbox sin filtro...")
        msgs = buscar_mensajes_inbox_por_asunto(
            asunto_contiene="",
            top=max(top_msgs, 120),
            since_days=since_days
        ) or []

    if not msgs:
        print("[DIAN] ❌ No pude listar Inbox para buscar 'Validación(es) DIAN'.")
        return None, None, None

    contenedores = [m for m in msgs if _is_validacion_dian_subject(m.get("subject") or "")]
    if not contenedores:
        print("[DIAN] ❌ No encontré correos contenedores con asunto tipo 'Validación(es) DIAN'.")
        for x in msgs[:10]:
            print(f"   - subj: {x.get('subject','')!r}")
        return None, None, None

    for m in contenedores:
        mid = m.get("id")
        if not mid:
            continue

        pdfs = listar_adjuntos_pdf(mid)
        if not pdfs:
            continue

        for att in pdfs:
            aname = att.get("name") or f"{att.get('id')}.pdf"
            aid = att.get("id")
            local = os.path.join(TMP_DIR, "dian_pdf_only", aname)

            ok = descargar_adjunto_por_id(mid, aid, local)
            if not ok:
                continue

            try:
                txt = extraer_texto_pdf(local)
                ident = parse_identificadores_pdf(txt)
            except Exception:
                ident = {}

            if _match_pdf_candidate(target_ident, target_pdf_name, ident, aname):
                print(f"[DIAN] ✅ Match PDF dentro de correo contenedor: {aname}")
                return local, mid, aid

            try:
                os.remove(local)
            except Exception:
                pass

    print("[DIAN] ❌ No encontré PDF coincidente dentro de los correos contenedores.")
    return None, None, None


def _extraer_descripciones_items_pdf(texto: str) -> str:
    t = (texto or "").replace("\u00a0", " ")
    lines = [ln.strip() for ln in t.splitlines() if ln.strip()]

    def find_idx(pat: str, start: int = 0) -> int:
        for i in range(start, len(lines)):
            if re.search(pat, lines[i], flags=re.IGNORECASE):
                return i
        return -1

    idx_desc = find_idx(r"^Descripci[oó]n$")
    if idx_desc < 0:
        idx_desc = find_idx(r"Descripci[oó]n")
    if idx_desc < 0:
        return ""

    idx_end = find_idx(r"Datos\s+Totales|Notas\s+Finales", start=idx_desc + 1)
    if idx_end < 0:
        idx_end = min(len(lines), idx_desc + 120)

    seg = lines[idx_desc + 1: idx_end]

    header_stop = {
        "nro.", "nro", "código", "codigo", "u/m", "cantidad", "precio unitario",
        "subtotal", "iva %", "iva%", "valor total", "valor total item", "moneda",
        "tasa de cambio", "impuestos", "total"
    }

    def looks_numeric(l: str) -> bool:
        if re.fullmatch(r"\d{1,15}", l or ""):
            return True
        if re.fullmatch(r"\d{1,3}(\.\d{3})*(,\d{2})?", l or ""):
            return True
        if re.fullmatch(r"\d+(,\d+)?", l or ""):
            return True
        return False

    def looks_money(l: str) -> bool:
        s = (l or "").upper()
        return ("$" in s) or ("COP" in s) or bool(re.search(r"\d{1,3}(\.\d{3})*(,\d{2})", s))

    descs: List[str] = []
    i = 0
    while i < len(seg):
        ln = seg[i]
        if re.fullmatch(r"\d{1,3}", ln):
            i += 1
            if i < len(seg) and re.fullmatch(r"\d{1,20}", seg[i]):
                i += 1

            parts: List[str] = []
            while i < len(seg):
                cur = seg[i]
                cur_norm = cur.strip().lower()

                if cur_norm in header_stop:
                    break
                if looks_money(cur) or looks_numeric(cur):
                    break
                if re.fullmatch(r"(UN|UND|KG|LT|GL|NIU|EA|H87|94|ZZ|\d{1,4})", cur.strip().upper()):
                    break

                parts.append(cur.strip())
                i += 1

            desc = re.sub(r"\s{2,}", " ", " ".join(parts)).strip()
            if desc:
                descs.append(desc)
        else:
            i += 1

    seen = set()
    out = []
    for d in descs:
        key = d.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(d)

    return "; ".join(out).strip()


def _generar_registro_pdf_only(pdf_local_path: str, pdf_name: str) -> Dict[str, object]:
    """
    Genera 1 registro estilo "XML" usando solo el PDF.
    IMPORTANTE: el campo Archivo debe ser el que realmente quieras ver y filtrar en Excel/SharePoint.
    """
    texto = extraer_texto_pdf(pdf_local_path)
    ident = parse_identificadores_pdf(texto)

    campos = extraer_campos_basicos_pdf(texto)
    tot = extraer_totales_basicos_pdf(texto)

    desc_items = _extraer_descripciones_items_pdf(texto)
    if desc_items:
        campos["DescripcionLineas"] = desc_items

    fecha = (ident.get("FECHA") or "").strip()
    y = fecha[:4] if len(fecha) >= 4 else ""
    mo = fecha[5:7] if len(fecha) >= 7 else ""
    d = fecha[8:10] if len(fecha) >= 10 else ""

    return {
        "Archivo": os.path.basename(pdf_name),
        "Empresa emisora": campos.get("Empresa emisora", ""),
        "CUFE": ident.get("CUFE", ""),
        "Ciudad emisora": campos.get("Ciudad emisora", ""),
        "Código ciudad": campos.get("Código ciudad", ""),
        "NIT": campos.get("NIT", ""),
        "Cliente": campos.get("Cliente", ""),
        "Número de factura": ident.get("NUMERO", "") or ident.get("NUMERO_APROB", ""),
        "Año": y,
        "Mes": mo,
        "Día": d,
        "Tipo de contribuyente": campos.get("Tipo de contribuyente", ""),
        "Actividad económica": campos.get("Actividad económica", ""),
        "DescripcionLineas": campos.get("DescripcionLineas", ""),
        "Subtotal": float(tot.get("Subtotal", 0.0) or 0.0),
        "IVA 5%": float(tot.get("IVA 5%", 0.0) or 0.0),
        "IVA 19%": float(tot.get("IVA 19%", 0.0) or 0.0),
        "Retención de IVA": 0.0,
        "Retención de ICA": 0.0,
        "Retención en la fuente": 0.0,
        "Total": float(tot.get("Total", 0.0) or 0.0),
    }


def _score_pdf_candidate(pdf_name: str, ident: Dict[str, str]) -> int:
    score = 0
    name = (pdf_name or "")

    if _is_acta_filename(name):
        score -= 200

    low = name.lower()
    if "factura" in low:
        score += 15
    if "representacion" in low or "representación" in low:
        score += 5
    if "dian" in low:
        score += 3

    cufe = ident.get("CUFE") or ""
    if _cufe_is_valid(cufe):
        score += 150
    elif cufe:
        score += 10

    num = (ident.get("NUMERO") or "").strip()
    if num:
        if _is_generic_or_non_invoice_numero(num):
            score -= 15
        else:
            score += 25

    fec = (ident.get("FECHA") or "").strip()
    if fec:
        score += 10

    if re.search(r"\bproyecto\b", (num or "").lower()):
        score -= 50

    return score


def _seleccionar_mejor_pdf(msg_id: str, subj: str, pdf_atts: List[dict]) -> Tuple[Optional[dict], Optional[str], Optional[dict]]:
    os.makedirs(TMP_DIR, exist_ok=True)

    subj_num = _numero_from_subject(subj)
    best_att = None
    best_path = None
    best_ident = None
    best_score = -10**9

    for i, att in enumerate(pdf_atts, start=1):
        aid = att.get("id")
        if not aid:
            continue

        aname = att.get("name") or f"{aid}.pdf"
        safe_name = re.sub(r"[^A-Za-z0-9_.\- ]", "_", aname)
        tmp_path = os.path.join(TMP_DIR, f"{msg_id[:8]}_{i}_{safe_name}")

        ok = descargar_adjunto_por_id(msg_id, aid, tmp_path)
        if not ok or not os.path.exists(tmp_path):
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass
            continue

        try:
            texto = extraer_texto_pdf(tmp_path)
            ident = parse_identificadores_pdf(texto)

            best_num = _prefer_subject_numero(ident.get("NUMERO"), subj_num)
            if best_num:
                ident["NUMERO"] = best_num

            if not ident.get("FECHA"):
                ident.setdefault("FECHA", _fecha_from_subject(subj))

            sc = _score_pdf_candidate(aname, ident)
        except Exception:
            ident = {}
            sc = -10**6

        if sc > best_score:
            if best_path and best_path != tmp_path:
                try:
                    os.remove(best_path)
                except Exception:
                    pass

            best_score = sc
            best_att = att
            best_path = tmp_path
            best_ident = ident
        else:
            try:
                os.remove(tmp_path)
            except Exception:
                pass

    return best_att, best_path, best_ident


# ============================================================
# ✅ Helper: subir por replace con reintentos (solo para archivos NO bloqueados)
# ============================================================
def _upload_replace_with_retries(local_path: str, sp_path: str, retries: int = 2, sleep_s: float = 1.5) -> bool:
    last = None
    for i in range(retries + 1):
        try:
            upload_small_file(local_path, sp_path, mode="replace")
            return True
        except Exception as e:
            last = e
            msg = str(e)
            if "423" in msg or "resourceLocked" in msg or "locked" in msg.lower():
                if i < retries:
                    time.sleep(sleep_s * (2 ** i))
                    continue
                print(f"⚠️ Recurso LOCKED (423). No se pudo reemplazar: {sp_path}")
                return False

            if i < retries:
                time.sleep(sleep_s * (2 ** i))
                continue
            print(f"⚠️ No se pudo reemplazar {sp_path}: {last}")
            return False
    return False


def _subir_excels_a_sharepoint(sp_excel_root: str, hubo_cambios_excel: bool, historial_actualizado: bool):
    """
    ✅ FIX DEFINITIVO 423:
    - facturas.xlsx NO se reemplaza jamás por /content (se actualiza por Workbook API).
    - Solo se sube historial_ejecuciones.xlsx por replace.
    """
    try:
        ensure_folder(sp_excel_root)
    except Exception as e:
        print(f"⚠️ No pude ensure_folder excel en SP: {e}")
        return

    if hubo_cambios_excel:
        print("ℹ️ facturas.xlsx NO se sube por replace (se actualiza por Workbook API).")

    if historial_actualizado:
        if os.path.exists(HISTORIAL_EXCEL):
            ok = _upload_replace_with_retries(
                HISTORIAL_EXCEL,
                f"{sp_excel_root}/historial_ejecuciones.xlsx",
                retries=2
            )
            if ok:
                print("✅ Subido historial_ejecuciones.xlsx a SharePoint (replace).")
        else:
            print(f"⚠️ No existe HISTORIAL_EXCEL local: {HISTORIAL_EXCEL}")


def _expand_archivos_ref(archivos_ref: set[str]) -> set[str]:
    """
    Expande posibles variantes para que obtener_filas_por_archivos encuentre filas
    aunque el 'Archivo' guardado difiera (PDF vs XML, mayúsculas/minúsculas, stem, etc.).
    """
    out = set()
    for a in (archivos_ref or set()):
        if not a:
            continue
        a = str(a).strip()
        if not a:
            continue

        base = os.path.basename(a)
        out.add(base)
        out.add(base.lower())
        out.add(base.upper())

        stem = Path(base).stem
        if stem:
            out.add(stem)
            out.add(stem.lower())
            out.add(stem.upper())

            # Variantes típicas
            out.add(f"{stem}.xml")
            out.add(f"{stem}.pdf")
            out.add(f"{stem}.zip")
            out.add(f"{stem}.XML")
            out.add(f"{stem}.PDF")
            out.add(f"{stem}.ZIP")
    return out


def _try_workbook_append(sp_excel_root: str, archivos_ref: set[str], table_name: str = "TblFacturas") -> int:
    """
    Inserta filas nuevas en TblFacturas usando Workbook API (dedup).
    """
    if not archivos_ref:
        return 0

    archivos_ref = _expand_archivos_ref(set(archivos_ref))

    filas = obtener_filas_por_archivos(archivos_ref)
    if not filas:
        return 0

    sp_facturas_path = f"{sp_excel_root}/facturas.xlsx".strip("/")

    xl = ExcelWorkbookGraph(sp_facturas_path)
    insertadas = xl.append_rows_dedup(
        table_name=table_name,
        rows_dicts=filas,
        key_cols=("Archivo", "Concepto"),
        require_table=True,
    )
    return int(insertadas or 0)


def run_desde_aprobadas(
    max_aprobados: int = 50,
    max_zip_buscar: int = 150,
    since_days: int | None = None
):
    if since_days is None:
        since_days = APROB_SEARCH_SINCE_DAYS

    lock = SingleInstanceLock(LOCK_FILE_APROBADAS, ttl_seconds=LOCK_TTL_SECONDS)
    if not lock.acquire():
        print("🧷 Otra instancia está corriendo. Salgo para evitar duplicados/interrupciones.")
        return

    run_id = datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:8]
    inicio_dt = datetime.datetime.now().isoformat(timespec="seconds")
    t0_total = time.perf_counter()

    os.makedirs(ADJ_HOY, exist_ok=True)
    os.makedirs(TMP_DIR, exist_ok=True)
    os.makedirs(EXT_HOY, exist_ok=True)

    store = ProcessedStore(PROCESSED_MESSAGES_PATH, ttl_days=PROCESSED_MESSAGES_TTL_DAYS)

    aidx = AttachmentIndexStore(ATTACHMENT_INDEX_PATH, ttl_days=ATTACHMENT_INDEX_TTL_DAYS)
    try:
        purged = aidx.purge()
        if purged:
            print(f"🧹 [AIDX] Purge index: removidas {purged} entradas viejas.")
    except Exception:
        pass

    folder_id = get_folder_id_by_name("Inbox", APROB_FOLDER_NAME) or find_folder_id_anywhere(APROB_FOLDER_NAME)
    if not folder_id:
        print(f"[APROB] No se encontró la carpeta: {APROB_FOLDER_NAME!r}")
        try:
            lock.release()
        except Exception:
            pass
        return

    print(f"📬 Leyendo carpeta de aprobadas (solo NO leídos): {APROB_FOLDER_NAME}")
    msgs = listar_mensajes_en_carpeta(folder_id, top=max_aprobados)
    msgs_leidos = len(msgs) if msgs else 0

    if not msgs:
        print("ℹ️ No hay mensajes no leídos con aprobaciones recientes.")
        total_secs = time.perf_counter() - t0_total
        print(f"⏱️ Tiempo total real: {total_secs:.2f} s")
        try:
            lock.release()
        except Exception:
            pass
        return

    # Performance: filtrar por store ANTES de prefetch ZIPs
    msgs_pendientes = []
    for m in msgs:
        mid = m.get("id")
        if mid and (not store.is_processed(mid)):
            msgs_pendientes.append(m)

    msgs_pendientes_count = len(msgs_pendientes)

    if not msgs_pendientes:
        print("✅ No hay mensajes nuevos para procesar (todos ya estaban en ProcessedStore). Salgo sin prefetch.")
        try:
            n = borrar_pdfs_en_arbol(TMP_DIR)
            print(f"🧹 Limpieza temp_check: borrados {n} PDF(s).")
        except Exception:
            pass

        total_secs = time.perf_counter() - t0_total
        print(f"⏱️ Tiempo total real: {total_secs:.2f} s")
        try:
            lock.release()
        except Exception:
            pass
        return

    msgs = msgs_pendientes

    idx_cufe, idx_nf = _build_zip_index(since_days=since_days, max_zip_buscar=max_zip_buscar, aidx=aidx)

    cufes_existentes = obtener_cufes_existentes()
    norm_cufes_existentes = {_norm_cufe(x) for x in cufes_existentes}
    print(f"ℹ️ CUFEs ya registrados en facturas.xlsx: {len(cufes_existentes)}")

    fecha_local = datetime.datetime.now().strftime("%Y-%m-%d")
    hora_local = datetime.datetime.now().strftime("%H:%M:%S")

    resumen: List[Tuple[str, float, str, int]] = []

    msgs_procesados = 0
    cnt_ok = 0
    cnt_sin_match = 0
    cnt_ya_reg = 0
    cnt_sin_pdf = 0
    cnt_err = 0
    cnt_dian = 0
    nuevos_total = 0
    enriq_total = 0

    detalle_rows: List[Dict[str, object]] = []

    procesados = 0
    sin_match_consec = 0
    sin_nuevos_consec = 0

    for msg in msgs:
        t0 = time.perf_counter()
        msg_id = msg["id"]
        subj = msg.get("subject") or ""

        if store.is_processed(msg_id):
            print(f"⏭️  Mensaje ya procesado (store). Se omite. id={msg_id}")
            continue

        pdf_atts = listar_adjuntos_pdf(msg_id)
        if not pdf_atts:
            store.mark_processed(msg_id, {"status": "sin_pdf"})
            cnt_sin_pdf += 1
            msgs_procesados += 1

            _push_detalle(
                detalle_rows, run_id, msg_id, subj,
                estado="sin_pdf",
                duracion_s=(time.perf_counter() - t0)
            )
            continue

        # Elegir mejor PDF si hay múltiples
        pdf = None
        pdf_tmp = None
        ident_pdf = {}

        if len(pdf_atts) == 1:
            pdf = pdf_atts[0]
            pdf_name = pdf.get("name") or f"{pdf['id']}.pdf"
            pdf_tmp = os.path.join(TMP_DIR, pdf_name)

            if not descargar_adjunto_por_id(msg_id, pdf["id"], pdf_tmp):
                print(f"[APROB] No pude descargar PDF {pdf_name}")
                store.mark_processed(msg_id, {"status": "error_descarga_pdf", "pdf": pdf_name})
                cnt_err += 1
                msgs_procesados += 1

                _push_detalle(
                    detalle_rows, run_id, msg_id, subj,
                    pdf_name=pdf_name,
                    estado="error_descarga_pdf",
                    duracion_s=(time.perf_counter() - t0),
                    error="No se pudo descargar PDF"
                )
                continue

            texto = extraer_texto_pdf(pdf_tmp)
            ident_pdf = parse_identificadores_pdf(texto)
        else:
            pdf, pdf_tmp, ident_pdf = _seleccionar_mejor_pdf(msg_id, subj, pdf_atts)
            if not pdf or not pdf_tmp:
                pdf = pdf_atts[0]
                pdf_name = pdf.get("name") or f"{pdf['id']}.pdf"
                pdf_tmp = os.path.join(TMP_DIR, pdf_name)
                if not descargar_adjunto_por_id(msg_id, pdf["id"], pdf_tmp):
                    print(f"[APROB] No pude descargar PDF (fallback) {pdf_name}")
                    store.mark_processed(msg_id, {"status": "error_descarga_pdf", "pdf": pdf_name})
                    cnt_err += 1
                    msgs_procesados += 1

                    _push_detalle(
                        detalle_rows, run_id, msg_id, subj,
                        pdf_name=pdf_name,
                        estado="error_descarga_pdf",
                        duracion_s=(time.perf_counter() - t0),
                        error="No se pudo descargar PDF (fallback)"
                    )
                    continue
                texto = extraer_texto_pdf(pdf_tmp)
                ident_pdf = parse_identificadores_pdf(texto)

        pdf_name = pdf.get("name") or f"{pdf['id']}.pdf"

        subj_num = _numero_from_subject(subj)
        best_num = _prefer_subject_numero(ident_pdf.get("NUMERO"), subj_num)
        if best_num:
            ident_pdf["NUMERO"] = best_num

        numero_aprob = (ident_pdf.get("NUMERO_APROB") or "").strip()
        if not numero_aprob:
            if subj_num and subj_num.strip() and subj_num.strip() != (ident_pdf.get("NUMERO") or "").strip():
                numero_aprob = subj_num.strip()

        if not ident_pdf.get("FECHA"):
            ident_pdf.setdefault("FECHA", _fecha_from_subject(subj))

        cufe_pdf = _norm_cufe(ident_pdf.get("CUFE") or "")
        fecha_pdf = (ident_pdf.get("FECHA") or "").strip()
        if fecha_pdf:
            fecha_pdf = normalizar_fecha(fecha_pdf) or fecha_pdf

        print("\n===== DEBUG PDF PARSE =====")
        print(f"→ PDF elegido: {pdf_name}")
        print(f"→ CUFE detectado: {ident_pdf.get('CUFE')}")
        print(f"→ NUMERO detectado: {ident_pdf.get('NUMERO')}")
        print(f"→ NUMERO_APROB detectado: {numero_aprob or None}")
        print(f"→ FECHA detectada: {ident_pdf.get('FECHA')}")
        print("===========================\n")

        # ============================================================
        # FLUJO ESPECIAL DIAN
        # ============================================================
        if _is_dian_trigger_message(msg):
            print(f"[DIAN] Detectado mensaje DIAN en aprobadas (asunto+cuerpo): {subj!r}")

            pdf_real_path, mid_src, aid_src = _buscar_pdf_en_correo_validaciones_dian(
                target_ident=ident_pdf,
                target_pdf_name=pdf_name,
                since_days=since_days,
                top_msgs=80
            )

            if not pdf_real_path:
                secs = time.perf_counter() - t0
                resumen.append((pdf_name, secs, "sin match dian", 0))
                store.mark_processed(msg_id, {"status": "sin_match_dian", "pdf": pdf_name, "cufe": cufe_pdf})
                cnt_sin_match += 1
                msgs_procesados += 1
                sin_match_consec += 1
                sin_nuevos_consec = 0
                procesados += 1

                _push_detalle(
                    detalle_rows, run_id, msg_id, subj,
                    pdf_name=pdf_name,
                    cufe=cufe_pdf,
                    numero=ident_pdf.get("NUMERO") or "",
                    fecha_factura=fecha_pdf,
                    estado="sin_match_dian",
                    duracion_s=(time.perf_counter() - t0),
                    fuente="DIAN"
                )
                continue

            # ✅ IMPORTANTE: Archivo debe coincidir con lo que realmente subes a SP y lo que quieres filtrar
            pdf_real_name = os.path.basename(pdf_real_path)
            reg = _generar_registro_pdf_only(pdf_real_path, pdf_real_name)
            if numero_aprob and len(numero_aprob) >= 5:
                reg["Número de factura"] = numero_aprob

            total_nuevos = guardar_en_excel([reg])

            historial_actualizado = False
            if total_nuevos > 0:
                registrar_historial_por_zip([{
                    "Fecha": fecha_local,
                    "Hora": hora_local,
                    "Archivo ZIP": "(PDF-ONLY) VALIDACIONES DIAN",
                    "Nuevos XML guardados": total_nuevos,
                    "Errores encontrados": 0
                }])
                historial_actualizado = True

            enriquecidas = 0
            try:
                enriquecidas = sincronizar_aprobaciones_en_facturas()
            except Exception as e:
                print(f"[APROB] Error al sincronizar aprobaciones: {e}")

            print("☁️  Subiendo a SharePoint (DIAN / PDF-only)...")
            sp_ext_root = f"{BASE_SP}/extraidos/dian_pdf_only"
            sp_excel = f"{BASE_SP}/excel"
            ensure_folder(sp_ext_root)
            ensure_folder(sp_excel)

            try:
                upload_small_file(pdf_real_path, f"{sp_ext_root}/{pdf_real_name}", mode="skip")
            except Exception as e:
                print(f"[DIAN] No pude subir PDF real: {e}")

            hubo_cambios_excel = (total_nuevos > 0) or (enriquecidas > 0)

            insertadas = 0
            if hubo_cambios_excel:
                try:
                    # ✅ usar exactamente el "Archivo" real que quedó en Excel
                    archivos_ref = {str(reg.get("Archivo") or pdf_real_name)}
                    insertadas = _try_workbook_append(sp_excel, archivos_ref, table_name="TblFacturas")
                    print(f"✅ Workbook API (DIAN): +{insertadas} fila(s) nuevas en TblFacturas.")
                except Exception as e:
                    print(f"⚠️ Workbook API falló (DIAN): {e}")

            _subir_excels_a_sharepoint(sp_excel, hubo_cambios_excel, historial_actualizado)

            store.mark_processed(msg_id, {
                "status": "ok_dian_pdf_only",
                "pdf": pdf_name,
                "nuevos": int(total_nuevos),
                "enriquecidas": int(enriquecidas),
                "cufe": cufe_pdf,
                "src_msg": mid_src,
                "src_att": aid_src,
            })

            try:
                marcar_mensaje_como_leido(msg_id)
            except Exception as e:
                print(f"⚠️ No se pudo marcar como leído: {e}")

            secs = time.perf_counter() - t0
            resumen.append((pdf_name, secs, "match dian", total_nuevos))

            _push_detalle(
                detalle_rows, run_id, msg_id, subj,
                pdf_name=pdf_name,
                cufe=cufe_pdf,
                numero=reg.get("Número de factura") or "",
                fecha_factura=ident_pdf.get("FECHA") or fecha_pdf,
                zip_match="(PDF-ONLY) VALIDACIONES DIAN",
                estado="ok_dian_pdf_only",
                duracion_s=(time.perf_counter() - t0),
                nuevos=int(total_nuevos or 0),
                enriquecidas=int(enriquecidas or 0),
                fuente="DIAN"
            )

            cnt_dian += 1
            msgs_procesados += 1
            nuevos_total += int(total_nuevos or 0)
            enriq_total += int(enriquecidas or 0)

            sin_match_consec = 0
            sin_nuevos_consec = 0 if total_nuevos > 0 else (sin_nuevos_consec + 1)
            procesados += 1
            continue

        # ------------------------------------------------------------
        # FLUJO NORMAL (ZIP/XML)
        # ------------------------------------------------------------
        if cufe_pdf and cufe_pdf in norm_cufes_existentes:
            print(f"🔁 Factura ya registrada (CUFE en Excel). Se omite búsqueda de ZIP para {pdf_name}.")
            secs = time.perf_counter() - t0
            resumen.append((pdf_name, secs, "ya registrado", 0))
            store.mark_processed(msg_id, {"status": "ya_registrado", "pdf": pdf_name, "cufe": cufe_pdf})

            _push_detalle(
                detalle_rows, run_id, msg_id, subj,
                pdf_name=pdf_name,
                cufe=cufe_pdf,
                numero=ident_pdf.get("NUMERO") or "",
                fecha_factura=fecha_pdf,
                estado="ya_registrado",
                duracion_s=(time.perf_counter() - t0)
            )

            cnt_ya_reg += 1
            msgs_procesados += 1

            sin_match_consec = 0
            sin_nuevos_consec += 1
            procesados += 1

            try:
                marcar_mensaje_como_leido(msg_id)
            except Exception as e:
                print(f"⚠️ No se pudo marcar como leído: {e}")
            if (procesados >= AUTO_STOP_MIN_PROCESADOS) and (sin_nuevos_consec >= AUTO_STOP_SIN_NUEVOS_CONSEC):
                print("🛑 Deteniendo flujo: varios PDFs ya registrados/sin nuevos (optimización de tiempo).")
                break
            continue

        found_match = False
        found_zip_name = None
        found_zip_bytes = None
        fuente_match = "normal"

        if cufe_pdf and cufe_pdf in idx_cufe:
            found_zip_name, found_zip_bytes = idx_cufe[cufe_pdf]
            found_match = True
        else:
            num_pdf = (ident_pdf.get("NUMERO") or "").strip()
            if num_pdf and fecha_pdf:
                for k in claves_normalizadas_factura(num_pdf):
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
            entry = None

            if cufe_pdf:
                entry = aidx.find_zip_by_cufe(cufe_pdf)

            if not entry:
                num_pdf = (ident_pdf.get("NUMERO") or "").strip()
                if num_pdf and fecha_pdf:
                    entry = aidx.find_zip_by_num_fecha(num_pdf, fecha_pdf)

            if entry:
                try:
                    zname = entry.get("att_name") or "factura.zip"
                    mid = entry.get("msg_id")
                    aid = entry.get("att_id")

                    if mid and aid:
                        print(f"🧠 [AIDX] Encontré ZIP histórico: {zname} (descargando directo por IDs)...")
                        tmp_zip = os.path.join(TMP_DIR, f"aidx_{zname}")
                        ok = descargar_adjunto_por_id(mid, aid, tmp_zip)
                        if ok and os.path.exists(tmp_zip):
                            with open(tmp_zip, "rb") as f:
                                found_zip_bytes = f.read()
                            try:
                                os.remove(tmp_zip)
                            except Exception:
                                pass

                            found_zip_name = zname
                            found_match = True
                            fuente_match = "AIDX"
                            print(f"✅ [AIDX] ZIP histórico listo en memoria: {found_zip_name}")
                except Exception as e:
                    print(f"⚠️ [AIDX] Falló descarga ZIP histórico: {e}")

        if not found_match or not found_zip_name or not found_zip_bytes:
            print(f"❌ No se encontró ZIP que coincida para PDF {pdf_name}.")
            secs = time.perf_counter() - t0
            resumen.append((pdf_name, secs, "sin match", 0))
            store.mark_processed(msg_id, {"status": "sin_match", "pdf": pdf_name, "cufe": cufe_pdf})

            _push_detalle(
                detalle_rows, run_id, msg_id, subj,
                pdf_name=pdf_name,
                cufe=cufe_pdf,
                numero=ident_pdf.get("NUMERO") or "",
                fecha_factura=fecha_pdf,
                estado="sin_match",
                duracion_s=(time.perf_counter() - t0),
                fuente=fuente_match
            )

            cnt_sin_match += 1
            msgs_procesados += 1

            sin_match_consec += 1
            sin_nuevos_consec = 0
            procesados += 1

            if (procesados >= AUTO_STOP_MIN_PROCESADOS) and (sin_match_consec >= AUTO_STOP_SIN_MATCH_CONSEC):
                print("🛑 Deteniendo flujo: varios PDFs consecutivos sin match (optimización de tiempo).")
                break
            continue

        b1 = _limpiar_adj_hoy()
        if b1:
            print(f"🧹 Limpieza ADJ_HOY: borrados {b1} ZIP(s) viejos.")

        b2 = _limpiar_ext_hoy()
        if b2:
            print(f"🧹 Limpieza EXT_HOY: borrados {b2} elemento(s) viejos.")

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

        # ✅ Aquí guardamos los "Archivo" reales que quedaron en regs (para subir a Workbook sí o sí)
        archivos_realmente_guardados: set[str] = set()

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

            if regs and numero_aprob:
                for dct in regs:
                    old = str(dct.get("Número de factura", "")).strip()
                    if old != numero_aprob and len(numero_aprob) >= 5:
                        dct["Número de factura"] = numero_aprob

            # Capturar "Archivo" antes de guardar (si existe)
            if regs:
                for dct in regs:
                    av = dct.get("Archivo")
                    if av:
                        archivos_realmente_guardados.add(str(av).strip())

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

        historial_actualizado = False
        if historial_rows:
            registrar_historial_por_zip(historial_rows)
            historial_actualizado = True

        enriquecidas = 0
        try:
            enriquecidas = sincronizar_aprobaciones_en_facturas()
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

        try:
            upload_small_file(str(zip_local_path), f"{sp_adj_root}/{found_zip_name}", mode="skip")
        except Exception as e:
            print(f"⚠️ Error subiendo ZIP a SharePoint: {e}")

        try:
            if carpeta_obj and ruta_obj and os.path.exists(ruta_obj):
                upload_directory(ruta_obj, f"{sp_ext_root}/{carpeta_obj}", mode="skip")
            else:
                upload_directory(EXT_HOY, sp_ext_root, mode="skip")
        except Exception as e:
            print(f"⚠️ Error subiendo extraídos a SharePoint: {e}")
        hubo_cambios_excel = (total_nuevos > 0) or (enriquecidas > 0)

        insertadas = 0
        if hubo_cambios_excel:
            try:
                archivos_xml = set()
                if ruta_obj and os.path.isdir(ruta_obj):
                    for fn in os.listdir(ruta_obj):
                        if fn.lower().endswith(".xml"):
                            archivos_xml.add(fn)

                # ✅ mezcla robusta: (1) lo real guardado, (2) xmls detectados, (3) pdf y zip
                archivos_ref = set(archivos_realmente_guardados)
                archivos_ref |= set(archivos_xml)
                archivos_ref.add(os.path.basename(pdf_name))
                if found_zip_name:
                    archivos_ref.add(os.path.basename(found_zip_name))

                insertadas = _try_workbook_append(sp_excel, archivos_ref, table_name="TblFacturas")
                print(f"✅ Workbook API: +{insertadas} fila(s) nuevas en TblFacturas.")
            except Exception as e:
                print(f"⚠️ Workbook API falló: {e}")

        _subir_excels_a_sharepoint(sp_excel, hubo_cambios_excel, historial_actualizado)

        print("🎉 Proceso por aprobadas finalizado para:", found_zip_name)
        secs = time.perf_counter() - t0
        resumen.append((pdf_name, secs, "match", total_nuevos))

        store.mark_processed(msg_id, {
            "status": "ok",
            "pdf": pdf_name,
            "zip": found_zip_name,
            "nuevos": int(total_nuevos),
            "enriquecidas": int(enriquecidas),
            "cufe": cufe_pdf,
        })

        try:
            marcar_mensaje_como_leido(msg_id)
        except Exception as e:
            print(f"⚠️ No se pudo marcar como leído: {e}")

        _push_detalle(
            detalle_rows, run_id, msg_id, subj,
            pdf_name=pdf_name,
            cufe=cufe_pdf,
            numero=ident_pdf.get("NUMERO") or "",
            fecha_factura=fecha_pdf,
            zip_match=found_zip_name,
            estado="ok",
            duracion_s=(time.perf_counter() - t0),
            nuevos=int(total_nuevos or 0),
            enriquecidas=int(enriquecidas or 0),
            fuente=fuente_match
        )

        cnt_ok += 1
        msgs_procesados += 1
        nuevos_total += int(total_nuevos or 0)
        enriq_total += int(enriquecidas or 0)

        sin_match_consec = 0
        if total_nuevos == 0:
            sin_nuevos_consec += 1
        else:
            sin_nuevos_consec = 0
            if cufe_pdf:
                cufes_existentes.add(cufe_pdf)
                norm_cufes_existentes.add(cufe_pdf)

        procesados += 1
        if (procesados >= AUTO_STOP_MIN_PROCESADOS) and (sin_nuevos_consec >= AUTO_STOP_SIN_NUEVOS_CONSEC):
            print("🛑 Deteniendo flujo: varios PDFs con match pero sin nuevos registros (optimización de tiempo).")
            break

    try:
        n = borrar_pdfs_en_arbol(TMP_DIR)
        print(f"🧹 Limpieza temp_check: borrados {n} PDF(s).")
    except Exception:
        print("⚠️ Limpieza temp_check: no se pudo completar (continuo).")

    total_secs = time.perf_counter() - t0_total
    fin_dt = datetime.datetime.now().isoformat(timespec="seconds")

    hubo_actividad = (msgs_procesados > 0) or (cnt_err > 0) or (nuevos_total > 0) or (cnt_dian > 0)
    if (not AUDIT_WRITE_ONLY_IF_ACTIVITY) or hubo_actividad:
        try:
            append_detalle_rows(AUDIT_DIR, AUDIT_DETALLE_PREFIX, detalle_rows)
        except Exception as e:
            print(f"⚠️ No pude escribir audit detalle CSV: {e}")

        try:
            append_run_summary(AUDIT_DIR, AUDIT_RUNS_PREFIX, {
                "run_id": run_id,
                "inicio": inicio_dt,
                "fin": fin_dt,
                "duracion_s": round(total_secs, 3),
                "carpeta": APROB_FOLDER_NAME,
                "since_days": since_days,
                "max_aprobados": max_aprobados,
                "max_zip_buscar": max_zip_buscar,
                "msgs_leidos": msgs_leidos,
                "msgs_pendientes": msgs_pendientes_count,
                "msgs_procesados": msgs_procesados,
                "ok": cnt_ok,
                "sin_match": cnt_sin_match,
                "ya_registrado": cnt_ya_reg,
                "sin_pdf": cnt_sin_pdf,
                "errores": cnt_err,
                "dian_pdf_only": cnt_dian,
                "nuevos_total": nuevos_total,
                "enriquecidas_total": enriq_total,
                "nota": ""
            })
        except Exception as e:
            print(f"⚠️ No pude escribir audit runs CSV: {e}")

    print("\n===== ⏱️ Resumen de tiempos (aprobadas) =====")
    for name, secs, estado, nuevos in resumen:
        print(f"• {name} -> {secs:.2f}s | {estado} | nuevos={nuevos}")
    print(f"⏱️ Tiempo total real de ejecución: {total_secs:.2f} s")
    print("=============================================")

    try:
        lock.release()
    except Exception:
        pass


def _numero_from_subject(subj: str) -> str | None:
    m = re.search(
        r"(?:Factura|#|N[o°\.]?)[^\w]*([A-Za-z]{1,6}\s*[-–—]?\s*\d{2,20}|[A-Za-z0-9\-\/\.]{3,})",
        subj or "",
        flags=re.IGNORECASE
    )
    if not m:
        return None
    s = (m.group(1) or "").strip()
    s = s.replace("–", "-").replace("—", "-")
    s = re.sub(r"\s*-\s*", "-", s)
    s = re.sub(r"\s+", "", s) if re.match(r"^[A-Za-z]{2,6}\s*\d{2,20}$", s) else s
    return s.strip() or None


def _fecha_from_subject(subj: str) -> str | None:
    for pat in [r"(\d{4}[-/]\d{2}[-/]\d{2})", r"(\d{2}[-/]\d{2}[-/]\d{4})"]:
        m = re.search(pat, subj or "")
        if m:
            s = m.group(1)
            return normalizar_fecha(s) or s
    return None