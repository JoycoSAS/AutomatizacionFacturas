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
    APROB_DIAN_KEYWORD,
    INBOX_DIAN_SUBJECT_CANDIDATES,
    REQUIRE_DIAN_IN_BODY_PREVIEW,
    AUDIT_DIR, AUDIT_RUNS_PREFIX, AUDIT_DETALLE_PREFIX, AUDIT_WRITE_ONLY_IF_ACTIVITY,
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

from services.aprobaciones_service import sincronizar_aprobaciones_en_facturas
from services.m365.excel_workbook_graph import ExcelWorkbookGraph


ADJ_HOY = os.path.join(DATA_DIR, "adjuntos", "hoy")
EXT_HOY = os.path.join(DATA_DIR, "extraidos", "hoy")

USE_DATE_SUBFOLDERS = False

_CTRL_REGEX = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
_AMP_FIX = re.compile(r"&(?!(?:[a-zA-Z]+|#\d+|#x[0-9A-Fa-f]+);)")
_NON_INVOICE_PREFIXES = {"DDI", "RAD", "RDI", "RDC", "REC", "RCP", "DOC", "REF"}


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


def _norm_numero(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip().upper()
    s = s.replace("–", "-").replace("—", "-").replace("_", "-")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^A-Z0-9\-]", "", s)
    s = re.sub(r"-{2,}", "-", s).strip("-")
    return s


def _solo_alnum(s: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", (s or "").upper())


def _normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def _clean_name(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def _is_hex_like_token(s: str) -> bool:
    s = _solo_alnum(s or "")
    if not s:
        return False
    if len(s) < 4:
        return True
    if re.fullmatch(r"[0-9A-F]+", s) and len(s) <= 12:
        return True
    return False


def _is_uuid_like_name(name: str) -> bool:
    stem = Path(name or "").stem.strip()
    if not stem:
        return False

    if re.fullmatch(
        r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}",
        stem
    ):
        return True

    if re.fullmatch(r"[0-9a-fA-F]{24,64}", stem):
        return True

    return False


def _token_es_util_para_match(token: str) -> bool:
    t = _solo_alnum(token or "")
    if not t:
        return False

    if re.fullmatch(r"\d{4,20}", t):
        return True

    if len(t) < 5:
        return False

    if _is_hex_like_token(t):
        return False

    if re.fullmatch(r"[A-Z]{1,3}\d{1,2}", t):
        return False

    return True


def _numero_variantes(numero: str) -> List[str]:
    n = _norm_numero(numero)
    if not n:
        return []

    variants = []

    def add(v: str):
        v = (v or "").strip()
        if v and v not in variants:
            variants.append(v)

    add(n)
    add(n.replace("-", ""))
    add(n.replace("-", " "))
    add(_solo_alnum(n))

    m = re.match(r"^([A-Z]+)(\d+)$", _solo_alnum(n))
    if m:
        pref, dig = m.groups()
        add(f"{pref}-{dig}")
        add(f"{pref} {dig}")
        add(f"{pref}{dig}")

    m2 = re.match(r"^([A-Z]+)-(\d+)$", n)
    if m2:
        pref, dig = m2.groups()
        add(f"{pref}{dig}")
        add(f"{pref} {dig}")

    m3 = re.match(r"^([A-Z]+)-(\d+)-(\d+)$", n)
    if m3:
        a, b, c = m3.groups()
        add(f"{a}{b}{c}")
        add(f"{a}-{b}{c}")
        add(f"{a} {b}{c}")

    return variants


def _normalizar_numero_match(s: str) -> str:
    if not s:
        return ""
    s = str(s).upper().strip()
    s = s.replace("–", "-").replace("—", "-").replace("_", "-")
    s = re.sub(r"\bFACTURA\b", "", s)
    s = re.sub(r"\bFACT\b", "", s)
    s = re.sub(r"\bNO\.\b", "", s)
    s = re.sub(r"\bNUMERO\b", "", s)
    s = re.sub(r"\bN[ÚU]MERO\b", "", s)
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s.strip()


def _numero_parece_valido(n: str) -> bool:
    n = (n or "").strip().upper()
    if not n:
        return False

    invalid_prefixes = (
        "NIT", "CUFE", "UUID", "DOC", "RAD", "RDI", "RDC", "REC", "RCP",
        "RADICADO", "ID"
    )
    if n.startswith(invalid_prefixes):
        return False

    solo = _solo_alnum(n)

    if solo.startswith(("RADICADO", "NIT", "ID", "CUFE", "UUID")):
        return False

    if len(solo) < 4:
        return False

    if re.fullmatch(r"[A-Z]+", n):
        return False

    return True


def _tokens_match_from_text(texto: str) -> List[str]:
    texto = (texto or "").upper()
    if not texto:
        return []

    patrones = [
        r"[A-Z]{1,10}\s*[-]?\s*\d{2,20}",
        r"[A-Z]{1,10}\s*[-]?\s*\d{1,10}\s*[-]?\s*\d{2,20}",
        r"\d{2,20}\s*[-]?\s*[A-Z]{1,6}",
        r"[A-Z]{2,20}\d{2,20}",
        r"\b\d{4,20}\b",
    ]

    out: List[str] = []
    seen = set()

    for pat in patrones:
        for m in re.finditer(pat, texto):
            raw = (m.group(0) or "").strip()
            k = _normalizar_numero_match(raw)
            if not k:
                continue

            if re.fullmatch(r"\d{1,3}", k):
                continue

            if re.fullmatch(r"\d{4,20}", k):
                if k not in seen:
                    seen.add(k)
                    out.append(k)
                continue

            if not _token_es_util_para_match(k):
                continue

            if k not in seen:
                seen.add(k)
                out.append(k)

    return out


def _variantes_match_numero(numero: str) -> List[str]:
    out = []
    seen = set()

    candidatos = []
    candidatos.extend(_numero_variantes(numero))
    candidatos.append(numero)

    for c in candidatos:
        k = _normalizar_numero_match(c)
        if not k:
            continue
        if not _token_es_util_para_match(k):
            continue
        if k not in seen:
            seen.add(k)
            out.append(k)

    return out


def _elegir_numero_principal(ident_pdf: Dict[str, str], subj: str, pdf_name: str) -> str:
    candidatos = [
        (ident_pdf.get("NUMERO_APROB") or "").strip(),
        (_numero_from_subject(subj) or "").strip(),
        (ident_pdf.get("NUMERO") or "").strip(),
    ]

    for c in candidatos:
        if _numero_parece_valido(c):
            return c

    if not _is_uuid_like_name(pdf_name):
        toks = _tokens_match_from_text(Path(pdf_name).stem)
        for t in toks:
            if _numero_parece_valido(t):
                return t

    for c in candidatos:
        if c:
            return c

    return ""


def _es_probable_factura_electronica(subj: str, pdf_name: str, ident: Dict[str, str]) -> bool:
    texto = f"{subj} {pdf_name}".upper()

    bloqueados_fuertes = [
        "CUENTA DE COBRO",
        "ACTA",
        "MEMORANDO",
        "OFICIO",
        "COMUNICADO",
        "CERTIFICADO",
        "CONSTANCIA",
        "SOLICITUD ANTICIPO",
    ]

    if any(x in texto for x in bloqueados_fuertes):
        return False

    if _cufe_is_valid(ident.get("CUFE") or ""):
        return True

    num = ident.get("NUMERO_APROB") or ident.get("NUMERO") or ""
    if _numero_parece_valido(num):
        return True

    return False


def _is_acta_filename(name: str) -> bool:
    s = (name or "").lower()
    s_clean = re.sub(r"\s+", " ", s)
    bad_keys = [
        "acta", "constancia", "certificado", "aprobacion", "aprobación",
        "memorando", "oficio", "radicado", "soporte de radicacion", "soporte de radicación",
        "documento", "comunicado"
    ]
    return any(k in s_clean for k in bad_keys)


def _contains_dian(text: str) -> bool:
    return normalize_text(APROB_DIAN_KEYWORD) in normalize_text(text or "")


def _contains_validacion(text: str) -> bool:
    s = normalize_text(text or "")
    return ("validacion" in s) or ("validaciones" in s)


def _is_validation_like_subject(subj: str) -> bool:
    s = normalize_text(subj or "")
    if not s:
        return False

    reglas_directas = [
        "02-validacion dian",
        "02 validacion dian",
        "02-validaciones dian",
        "02 validaciones dian",
        "validacion dian",
        "validaciones dian",
        "validacion dian joyco",
        "validaciones dian joyco",
        "dian vial",
        "correo validacion dian",
        "correo validaciones dian",
        "validacion joyco",
        "validaciones joyco",
        "validacion joyco sas",
        "validaciones joyco sas",
        "validacion joyco s a s",
        "validaciones joyco s a s",
        "02-validacion joyco",
        "02 validacion joyco",
        "02-validaciones joyco",
        "02 validaciones joyco",
        "correo validacion joyco",
        "correo validaciones joyco",
    ]

    if any(x in s for x in reglas_directas):
        return True

    tiene_dian = "dian" in s
    tiene_joyco = "joyco" in s
    tiene_validacion = _contains_validacion(s)

    if tiene_validacion and (tiene_dian or tiene_joyco):
        return True

    try:
        if any(normalize_text(c) in s for c in INBOX_DIAN_SUBJECT_CANDIDATES):
            return True
    except Exception:
        pass

    return False


def _is_inbox_dian_container_subject(subj: str) -> bool:
    s = normalize_text(subj or "")
    if not s:
        return False

    reglas_directas = [
        "02-validacion dian",
        "02 validacion dian",
        "02-validaciones dian",
        "02 validaciones dian",
        "validacion dian",
        "validaciones dian",
        "validacion dian joyco",
        "validaciones dian joyco",
        "dian vial",
        "correo validacion dian",
        "correo validaciones dian",
        "validacion joyco",
        "validaciones joyco",
        "validacion joyco sas",
        "validaciones joyco sas",
        "validacion joyco s a s",
        "validaciones joyco s a s",
    ]

    if any(x in s for x in reglas_directas):
        return True

    tiene_dian = "dian" in s
    tiene_validacion = ("validacion" in s) or ("validaciones" in s)
    tiene_joyco = "joyco" in s

    if tiene_validacion and (tiene_dian or tiene_joyco):
        return True

    try:
        if any(normalize_text(c) in s for c in INBOX_DIAN_SUBJECT_CANDIDATES):
            return True
    except Exception:
        pass

    return False


def _is_dian_trigger_message(msg: dict) -> bool:
    subj = msg.get("subject") or ""
    subj_norm = normalize_text(subj)

    preview = (msg.get("bodyPreview") or msg.get("body_preview") or "")
    preview_norm = normalize_text(preview)

    # 1) Casos claramente DIAN / validación
    if _is_validation_like_subject(subj):
        return True

    # 2) Si el asunto dice DIAN pero también es un correo típico de aprobadas,
    # NO lo mandamos automáticamente por rama DIAN
    asunto_aprobado = (
        ("aprobado" in subj_norm)
        and ("factura" in subj_norm)
        and ("radicado" in subj_norm)
    )

    if "dian" in subj_norm and not asunto_aprobado:
        return True

    # 3) Revisión por cuerpo / preview
    if REQUIRE_DIAN_IN_BODY_PREVIEW:
        if preview:
            if "dian" in preview_norm and not asunto_aprobado:
                return True
            if _contains_validacion(preview_norm) and "joyco" in preview_norm:
                return True
            return False

        return _is_validation_like_subject(subj)

    if "dian" in preview_norm and not asunto_aprobado:
        return True

    if _contains_validacion(preview_norm) and "joyco" in preview_norm:
        return True

    return False


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
        return _clean_xml_text(m.group(1))

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
            root = ET.fromstring(inner)
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

        m = __re(r"<(?:cbc:|)ID[^>]*>([^<]{3,})</", text)
        if m and not ident.get("NUMERO"):
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

    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
            for member in zf.infolist():
                if not member.filename.lower().endswith(".xml"):
                    continue
                try:
                    xml_data = zf.read(member)
                    ident = _parse_ident_from_xml_bytes(xml_data)
                    ident["xml_name"] = Path(member.filename).name
                    out.append(ident)
                except Exception as e:
                    print(f"[ZIP] No se pudo leer {member.filename}: {e}")
    except Exception as e:
        print(f"[ZIP] Error abriendo ZIP en memoria: {e}")

    return out


def _is_generic_or_non_invoice_numero(n: str) -> bool:
    n = (n or "").strip().upper()
    if not n:
        return True

    if n.startswith(("RADICADO", "NIT", "ID", "CUFE", "UUID")):
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

    def invalido(n: str) -> bool:
        n2 = (n or "").strip().upper()
        if not n2:
            return True
        if n2.startswith(("RADICADO", "NIT", "ID", "CUFE", "UUID")):
            return True
        if re.fullmatch(r"RADICADO\d+", _solo_alnum(n2)):
            return True
        if re.fullmatch(r"NIT\d+", _solo_alnum(n2)):
            return True
        if re.fullmatch(r"ID\d+", _solo_alnum(n2)):
            return True
        return False

    if subj_num and not invalido(subj_num):
        if not pdf_num or invalido(pdf_num):
            return subj_num

    if pdf_num and not invalido(pdf_num):
        return pdf_num

    if subj_num:
        return subj_num
    return pdf_num or None


def _build_num_set(ident: Dict[str, str]) -> set[str]:
    out = set()
    for x in [
        ident.get("NUMERO"),
        ident.get("NUMERO_APROB"),
    ]:
        if not x:
            continue

        x_norm = _norm_numero(x)
        x_alnum = _solo_alnum(x_norm)

        if x_norm:
            out.add(x_norm)
        if x_alnum:
            out.add(x_alnum)

        for v in _numero_variantes(x):
            if v:
                out.add(v)
                out.add(_solo_alnum(v))

        for v in _variantes_match_numero(x):
            if v:
                out.add(v)
                out.add(_solo_alnum(v))

    return {z for z in out if z}


def _safe_pdf_ident(local_pdf: str) -> Dict[str, str]:
    try:
        txt = extraer_texto_pdf(local_pdf)
        ident = parse_identificadores_pdf(txt) or {}
        return ident
    except Exception as e:
        print(f"[DIAN] ⚠️ Error leyendo PDF para ident: {local_pdf} | {e}")
        return {}


def _match_pdf_candidate(
    target_ident: Dict[str, str],
    target_pdf_name: str,
    candidate_ident: Dict[str, str],
    candidate_name: str
) -> bool:
    t_cufe = _norm_cufe(target_ident.get("CUFE") or "")
    c_cufe = _norm_cufe(candidate_ident.get("CUFE") or "")

    if t_cufe and c_cufe and t_cufe == c_cufe:
        print(f"[DIAN MATCH] ✅ Match por CUFE | {candidate_name}")
        return True

    nums_target = _build_num_set(target_ident)
    nums_cand = _build_num_set(candidate_ident)

    inter = nums_target & nums_cand
    if inter:
        print(f"[DIAN MATCH] ✅ Match por número | intersección={sorted(list(inter))[:5]} | {candidate_name}")
        return True

    target_name_clean = _solo_alnum(Path(target_pdf_name).stem)
    cand_name_clean = _solo_alnum(Path(candidate_name).stem)

    if target_name_clean and cand_name_clean:
        if target_name_clean == cand_name_clean:
            print(f"[DIAN MATCH] ✅ Match por nombre exacto limpio | {candidate_name}")
            return True

        if target_name_clean in cand_name_clean or cand_name_clean in target_name_clean:
            print(f"[DIAN MATCH] ✅ Match por nombre contenido | {candidate_name}")
            return True

    for n in nums_target:
        n_alnum = _solo_alnum(n)
        if not n_alnum:
            continue
        if len(n_alnum) < 5:
            continue
        if n_alnum in cand_name_clean:
            print(f"[DIAN MATCH] ✅ Match por número en nombre candidato | num={n_alnum} | {candidate_name}")
            return True

    return False


def _numero_from_subject(subj: str) -> str | None:
    if not subj:
        return None

    s = subj.strip()

    patrones = [
        r"Aprobado-\s*Factura(?:\s+de\s+servicio\s+p[uú]blico)?\s*-\s*([A-Za-z]{1,10}(?:\s*[-–—]?\s*\d+){1,3})\s*-\s*Radicado",
        r"Factura(?:\s+de\s+servicio\s+p[uú]blico)?\s*-\s*([A-Za-z]{1,10}(?:\s*[-–—]?\s*\d+){1,3})\s*-\s*Radicado",
        r"Factura(?:\s+de\s+servicio\s+p[uú]blico)?\s*-\s*([A-Za-z]{1,10}(?:\s*[-–—]?\s*\d+){1,3})",
        r"Aprobado-\s*Factura\s*-\s*(\d{3,20})\s*-\s*Radicado",
    ]

    for pat in patrones:
        m = re.search(pat, s, flags=re.IGNORECASE)
        if not m:
            continue

        raw = (m.group(1) or "").strip()
        raw = raw.replace("–", "-").replace("—", "-")
        raw = re.sub(r"\s*-\s*", "-", raw)
        raw = re.sub(r"\s+", "", raw)

        test = _solo_alnum(raw)
        if test.startswith(("RADICADO", "NIT", "ID", "CUFE", "UUID")):
            continue

        return raw.strip()

    return None


def _fecha_from_subject(subj: str) -> str | None:
    for pat in [r"(\d{4}[-/]\d{2}[-/]\d{2})", r"(\d{2}[-/]\d{2}[-/]\d{4})"]:
        m = re.search(pat, subj or "")
        if m:
            s = m.group(1)
            return normalizar_fecha(s) or s
    return None


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

    if ident.get("FECHA"):
        score += 5

    if re.search(r"\bproyecto\b", (num or "").lower()):
        score -= 50

    return score


def _seleccionar_mejor_pdf(
    msg_id: str,
    subj: str,
    pdf_atts: List[dict]
) -> Tuple[Optional[dict], Optional[str], Optional[dict]]:
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
        safe_name = re.sub(r"[^A-Za-z0-9_. -]", "_", aname)
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
            ident = parse_identificadores_pdf(texto) or {}

            best_num = _prefer_subject_numero(ident.get("NUMERO"), subj_num)
            if best_num:
                ident["NUMERO"] = best_num

            if subj_num and subj_num != (ident.get("NUMERO") or "").strip():
                ident["NUMERO_APROB"] = subj_num

            if not ident.get("FECHA"):
                ident.setdefault("FECHA", _fecha_from_subject(subj))

            sc = _score_pdf_candidate(aname, ident)

            if _is_uuid_like_name(aname):
                sc -= 5

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


def _debug_top_similares_idx(idx_num: Dict[str, Tuple[str, bytes]], numero: str, limite: int = 15):
    try:
        objetivo = _solo_alnum(numero or "")
        if not objetivo:
            print("[DEBUG TOP] sin objetivo para comparar")
            return

        candidatos = []
        for k, (zname, _zbytes) in idx_num.items():
            ka = _solo_alnum(k)
            if not ka:
                continue

            score = 0
            if objetivo == ka:
                score += 100
            if objetivo in ka or ka in objetivo:
                score += 40

            common = 0
            for ch in set(objetivo):
                if ch in ka:
                    common += 1
            score += common

            if score > 0:
                candidatos.append((score, k, zname))

        candidatos.sort(reverse=True, key=lambda x: x[0])

        print(f"[DEBUG TOP] similares para numero={numero} / objetivo={objetivo}")
        for score, k, zname in candidatos[:limite]:
            print(f"   score={score:03d} | idx_num={k} | zip={zname}")

        if not candidatos:
            print("[DEBUG TOP] no encontré similares en idx_num")
    except Exception as e:
        print(f"[DEBUG TOP] error viendo similares: {e}")


def _tokens_dian_objetivo(target_ident: Dict[str, str], target_pdf_name: str) -> List[str]:
    out = []
    seen = set()

    candidatos = [
        target_ident.get("NUMERO_APROB") or "",
        target_ident.get("NUMERO") or "",
    ]

    if not _is_uuid_like_name(target_pdf_name):
        candidatos.append(Path(target_pdf_name).stem)

    for c in candidatos:
        if not c:
            continue

        for v in _numero_variantes(c):
            va = _solo_alnum(v)
            if va and _token_es_util_para_match(va) and va not in seen:
                seen.add(va)
                out.append(va)

        for v in _variantes_match_numero(c):
            va = _solo_alnum(v)
            if va and _token_es_util_para_match(va) and va not in seen:
                seen.add(va)
                out.append(va)

        toks = _tokens_match_from_text(c)
        for t in toks:
            ta = _solo_alnum(t)
            if ta and _token_es_util_para_match(ta) and ta not in seen:
                seen.add(ta)
                out.append(ta)

    return out


def _build_zip_index(
    since_days: int,
    max_zip_buscar: int,
    aidx: AttachmentIndexStore
) -> Tuple[
    Dict[str, Tuple[str, bytes]],
    Dict[str, Tuple[str, bytes]],
    Dict[str, Tuple[str, bytes]]
]:
    idx_cufe: Dict[str, Tuple[str, bytes]] = {}
    idx_num: Dict[str, Tuple[str, bytes]] = {}
    idx_num_match: Dict[str, Tuple[str, bytes]] = {}

    mensajes_fuente = []

    try:
        inbox_msgs = listar_mensajes_zip_inbox(top=max_zip_buscar, since_days=since_days) or []
        mensajes_fuente.extend(inbox_msgs)
    except Exception as e:
        print(f"[ZIP INDEX] Error en listar_mensajes_zip_inbox: {e}")

    busquedas_fallback = [
        "DIAN",
        "Factura",
        "factura",
        "",
    ]

    for termino in busquedas_fallback:
        try:
            extra = buscar_mensajes_inbox_por_asunto(
                asunto_contiene=termino,
                top=max(200, min(max_zip_buscar, 1500)),
                since_days=since_days
            ) or []
            mensajes_fuente.extend(extra)
        except Exception as e:
            print(f"[ZIP INDEX] Error fallback asunto={termino!r}: {e}")

    dedup = []
    seen_msg = set()
    for m in mensajes_fuente:
        mid = m.get("id")
        if not mid:
            continue
        if mid in seen_msg:
            continue
        seen_msg.add(mid)
        dedup.append(m)

    limite_utc = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=since_days)

    candidatos = []
    for imsg in dedup:
        rdt = imsg.get("receivedDateTime")
        if rdt:
            try:
                rdt_dt = datetime.datetime.fromisoformat(rdt.replace("Z", "+00:00"))
                if rdt_dt < limite_utc:
                    continue
            except Exception:
                pass
        candidatos.append(imsg)

    print(f"📦 Prefetch ZIPs: {len(candidatos)} mensajes candidatos (ventana {since_days} día(s))")

    vistos_zip_ids = set()

    for i_msg, imsg in enumerate(candidatos, start=1):
        mid = imsg.get("id")
        if not mid:
            continue

        try:
            zips = listar_adjuntos_zip(mid) or []
        except Exception as e:
            print(f"[AIDX DEBUG] No pude listar adjuntos ZIP del mensaje {mid}: {e}")
            zips = []

        try:
            print(
                f"[AIDX DEBUG] mensaje={mid} | "
                f"asunto={(imsg.get('subject') or '')[:120]} | "
                f"zip_count={len(zips)} | "
                f"progreso={i_msg}/{len(candidatos)}"
            )
        except Exception:
            pass

        if not zips:
            continue

        asunto_zip = imsg.get("subject") or ""

        for z in zips:
            zid = z.get("id")
            if not zid:
                continue

            if zid in vistos_zip_ids:
                continue
            vistos_zip_ids.add(zid)

            zname = z.get("name") or f"{zid}.zip"

            try:
                print(f"[AIDX DEBUG] ZIP detectado: {zname} | id={zid}")
            except Exception:
                pass

            tmp_zip = os.path.join(TMP_DIR, f"prefetch_{uuid.uuid4().hex}_{re.sub(r'[^A-Za-z0-9_. -]', '_', zname)}")
            if not descargar_adjunto_por_id(mid, zid, tmp_zip):
                print(f"[AIDX DEBUG] No se pudo descargar ZIP: {zname}")
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

            if not idents_xml:
                print(f"[AIDX DEBUG] ZIP sin XMLs útiles o no legibles: {zname}")

            for tk in _tokens_match_from_text(asunto_zip):
                if _token_es_util_para_match(tk) and tk not in idx_num_match:
                    idx_num_match[tk] = (zname, zip_bytes)

            if not _is_uuid_like_name(zname):
                for tk in _tokens_match_from_text(Path(zname).stem):
                    if _token_es_util_para_match(tk) and tk not in idx_num_match:
                        idx_num_match[tk] = (zname, zip_bytes)

                zn_clean = _solo_alnum(Path(zname).stem)
                if _token_es_util_para_match(zn_clean) and zn_clean not in idx_num_match:
                    idx_num_match[zn_clean] = (zname, zip_bytes)

            for ident_xml in idents_xml:
                cufe = _norm_cufe(ident_xml.get("CUFE") or "")
                num_raw = (ident_xml.get("NUMERO") or "").strip()
                fec_raw = (ident_xml.get("FECHA") or "").strip()
                fec_norm = (normalizar_fecha(fec_raw) or fec_raw) if fec_raw else ""

                try:
                    print(
                        f"[AIDX DEBUG] XML en ZIP={zname} | "
                        f"xml_name={ident_xml.get('xml_name')} | "
                        f"CUFE={cufe or '-'} | NUMERO={num_raw or '-'} | FECHA={fec_norm or '-'}"
                    )
                except Exception:
                    pass

                try:
                    aidx.upsert_zip(
                        cufe=cufe,
                        numero=num_raw,
                        fecha=fec_norm,
                        msg_id=mid,
                        att_id=zid,
                        att_name=zname,
                        received_dt_iso=imsg.get("receivedDateTime", "") or "",
                    )
                except Exception as e:
                    print(f"[AIDX] No pude upsert ZIP index: {e}")

                if cufe and cufe not in idx_cufe:
                    idx_cufe[cufe] = (zname, zip_bytes)

                if num_raw:
                    for k in _numero_variantes(num_raw):
                        if k and k not in idx_num:
                            idx_num[k] = (zname, zip_bytes)

                    for mk in _variantes_match_numero(num_raw):
                        if mk and mk not in idx_num_match:
                            idx_num_match[mk] = (zname, zip_bytes)

            if len(idx_cufe) >= 2500 and len(idx_num_match) >= 2500:
                print("🛑 Índice suficiente; deteniendo prefetch temprano.")
                print(f"✅ Índice parcial: {len(idx_cufe)} por CUFE, {len(idx_num)} por NUMERO, {len(idx_num_match)} por MATCH")
                return idx_cufe, idx_num, idx_num_match

    print(f"✅ Índice listo: {len(idx_cufe)} por CUFE, {len(idx_num)} por NUMERO, {len(idx_num_match)} por MATCH")
    return idx_cufe, idx_num, idx_num_match


def _buscar_zip_por_numero(
    idx_num: Dict[str, Tuple[str, bytes]],
    *numeros: str
) -> Tuple[Optional[str], Optional[bytes], Optional[str]]:
    vistos = []
    for n in numeros:
        for v in _numero_variantes(n):
            if v not in vistos:
                vistos.append(v)

    print(f"[DEBUG NUMERO] candidatos exactos idx_num={vistos}")

    for cand in vistos:
        if cand in idx_num:
            zname, zbytes = idx_num[cand]
            print(f"[DEBUG NUMERO] MATCH EXACTO -> cand={cand} | zip={zname}")
            return zname, zbytes, cand

    cand_alnum = [_solo_alnum(x) for x in vistos if x]
    print(f"[DEBUG NUMERO] candidatos alnum idx_num={cand_alnum}")

    for k, val in idx_num.items():
        k_alnum = _solo_alnum(k)
        for c in cand_alnum:
            if c and k_alnum == c:
                zname, zbytes = val
                print(f"[DEBUG NUMERO] MATCH ALNUM -> cand={c} | idx={k} | zip={zname}")
                return zname, zbytes, c

    print("[DEBUG NUMERO] sin match en idx_num")
    return None, None, None


def _buscar_zip_por_numero_match(
    idx_num_match: Dict[str, Tuple[str, bytes]],
    *numeros: str
) -> Tuple[Optional[str], Optional[bytes], Optional[str]]:
    vistos = []

    for n in numeros:
        if not n:
            continue
        for v in _variantes_match_numero(n):
            if v not in vistos:
                vistos.append(v)

    print(f"[DEBUG MATCH] candidatos idx_num_match={vistos}")

    for cand in vistos:
        if cand in idx_num_match:
            zname, zbytes = idx_num_match[cand]
            print(f"[DEBUG MATCH] MATCH -> cand={cand} | zip={zname}")
            return zname, zbytes, cand

    print("[DEBUG MATCH] sin match en idx_num_match")
    return None, None, None


def _buscar_pdf_en_correo_validaciones_dian(
    target_ident: Dict[str, str],
    target_pdf_name: str,
    since_days: int,
    top_msgs: int = 200
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    dian_tmp = os.path.join(TMP_DIR, "dian_pdf_only")
    os.makedirs(dian_tmp, exist_ok=True)

    print("\n[DIAN] ================= INICIO BÚSQUEDA DIAN =================")
    print(f"[DIAN] target_pdf_name={target_pdf_name}")
    print(f"[DIAN] target CUFE={target_ident.get('CUFE')}")
    print(f"[DIAN] target NUMERO={target_ident.get('NUMERO')}")
    print(f"[DIAN] target NUMERO_APROB={target_ident.get('NUMERO_APROB')}")
    print(f"[DIAN] since_days={since_days} | top_msgs={top_msgs}")

    tokens_obj = _tokens_dian_objetivo(target_ident, target_pdf_name)
    print(f"[DIAN] tokens objetivo={tokens_obj}")

    candidatos_totales = []

    busquedas = [
        "DIAN",
        "Validacion",
        "Validaciones",
        "JOYCO",
        "VALIDACION JOYCO",
        "VALIDACION JOYCO SAS",
        "",
    ]

    for term in busquedas:
        try:
            lote = buscar_mensajes_inbox_por_asunto(
                asunto_contiene=term,
                top=top_msgs,
                since_days=since_days
            ) or []
            candidatos_totales.extend(lote)
            print(f"[DIAN] candidatos búsqueda {term!r}: {len(lote)}")
        except Exception as e:
            print(f"[DIAN] Error buscando {term!r}: {e}")

    msgs = []
    seen = set()
    for m in candidatos_totales:
        mid = m.get("id")
        if not mid or mid in seen:
            continue
        seen.add(mid)
        msgs.append(m)

    print(f"[DIAN] candidatos deduplicados: {len(msgs)}")

    if not msgs:
        print("[DIAN] ❌ No pude listar Inbox para buscar contenedores DIAN.")
        print("[DIAN] ================= FIN BÚSQUEDA DIAN =================\n")
        return None, None, None

    contenedores = []
    for m in msgs:
        subj = m.get("subject") or ""
        subj_norm = normalize_text(subj)
        subj_alnum = _solo_alnum(subj)

        if _is_inbox_dian_container_subject(subj):
            contenedores.append(m)
            continue

        if ("dian" in subj_norm) or ("joyco" in subj_norm) or _contains_validacion(subj_norm):
            if any(tk and tk in subj_alnum for tk in tokens_obj):
                contenedores.append(m)

    tmp = []
    seen_ids = set()
    for m in contenedores:
        mid = m.get("id")
        if mid and mid not in seen_ids:
            seen_ids.add(mid)
            tmp.append(m)
    contenedores = tmp

    print(f"[DIAN] contenedores filtrados: {len(contenedores)}")
    for i, m in enumerate(contenedores[:20], start=1):
        print(f"[DIAN] contenedor[{i}] asunto={m.get('subject')}")

    if not contenedores:
        print("[DIAN] ❌ No encontré correos contenedores con asunto tipo validación dian/joyco.")
        print("[DIAN] ================= FIN BÚSQUEDA DIAN =================\n")
        return None, None, None

    revisados = 0
    target_pdf_clean = _solo_alnum(Path(target_pdf_name).stem)

    for m in contenedores:
        mid = m.get("id")
        subj = m.get("subject") or ""

        if not mid:
            continue

        print(f"\n[DIAN] Revisando contenedor: {subj} | id={mid}")

        pdfs = listar_adjuntos_pdf(mid) or []
        print(f"[DIAN] PDFs adjuntos en contenedor: {len(pdfs)}")

        if not pdfs:
            continue

        def _prio(att: dict) -> int:
            aname = att.get("name") or ""
            clean = _solo_alnum(Path(aname).stem)
            p = 0
            if target_pdf_clean and clean == target_pdf_clean:
                p += 1000
            if target_pdf_clean and target_pdf_clean in clean:
                p += 400
            if any(tk and tk in clean for tk in tokens_obj):
                p += 200
            return -p

        pdfs = sorted(pdfs, key=_prio)

        for att in pdfs:
            revisados += 1
            aname = att.get("name") or f"{att.get('id')}.pdf"
            aid = att.get("id")
            if not aid:
                continue

            aname_alnum = _solo_alnum(Path(aname).stem)

            nombre_sugiere_match = False
            if target_pdf_clean and (
                aname_alnum == target_pdf_clean or
                target_pdf_clean in aname_alnum or
                aname_alnum in target_pdf_clean
            ):
                nombre_sugiere_match = True

            if not nombre_sugiere_match:
                for tk in tokens_obj:
                    if tk and tk in aname_alnum:
                        nombre_sugiere_match = True
                        break

            safe_name = re.sub(r"[^A-Za-z0-9_. -]", "_", aname)
            local = os.path.join(dian_tmp, f"{mid[:8]}_{aid[:8]}_{safe_name}")

            ok = descargar_adjunto_por_id(mid, aid, local)
            if not ok or not os.path.exists(local):
                print(f"[DIAN] ⚠️ No pude descargar adjunto PDF: {aname}")
                try:
                    if os.path.exists(local):
                        os.remove(local)
                except Exception:
                    pass
                continue

            ident = _safe_pdf_ident(local)

            print(
                f"[DIAN] PDF revisado: {aname} | "
                f"CUFE={ident.get('CUFE')} | "
                f"NUMERO={ident.get('NUMERO')} | "
                f"NUMERO_APROB={ident.get('NUMERO_APROB')} | "
                f"FECHA={ident.get('FECHA')} | "
                f"nombre_sugiere_match={nombre_sugiere_match}"
            )

            if _match_pdf_candidate(target_ident, target_pdf_name, ident, aname):
                print(f"[DIAN] ✅ Match PDF dentro de correo contenedor: {aname}")
                print(f"[DIAN] revisados_total={revisados}")
                print("[DIAN] ================= FIN BÚSQUEDA DIAN =================\n")
                return local, mid, aid

            if nombre_sugiere_match:
                print(f"[DIAN] ✅ Match por nombre candidato DIAN: {aname}")
                print(f"[DIAN] revisados_total={revisados}")
                print("[DIAN] ================= FIN BÚSQUEDA DIAN =================\n")
                return local, mid, aid

            try:
                os.remove(local)
            except Exception:
                pass

    print(f"[DIAN] ❌ No encontré PDF coincidente dentro de los correos contenedores.")
    print(f"[DIAN] revisados_total={revisados}")
    print("[DIAN] ================= FIN BÚSQUEDA DIAN =================\n")
    return None, None, None


def _buscar_zip_en_correo_validaciones_dian(
    target_ident: Dict[str, str],
    target_pdf_name: str,
    since_days: int,
    top_msgs: int = 200
) -> Tuple[Optional[str], Optional[bytes], Optional[str], Optional[str]]:
    print("\n[DIAN ZIP] ================ INICIO BÚSQUEDA ZIP DIAN ================")
    print(f"[DIAN ZIP] target_pdf_name={target_pdf_name}")
    print(f"[DIAN ZIP] target CUFE={target_ident.get('CUFE')}")
    print(f"[DIAN ZIP] target NUMERO={target_ident.get('NUMERO')}")
    print(f"[DIAN ZIP] target NUMERO_APROB={target_ident.get('NUMERO_APROB')}")

    tokens_obj = _tokens_dian_objetivo(target_ident, target_pdf_name)
    print(f"[DIAN ZIP] tokens objetivo={tokens_obj}")

    candidatos_totales = []

    busquedas = [
        "DIAN",
        "Validacion",
        "Validaciones",
        "JOYCO",
        "VALIDACION JOYCO",
        "VALIDACION JOYCO SAS",
        "",
    ]

    for term in busquedas:
        try:
            lote = buscar_mensajes_inbox_por_asunto(
                asunto_contiene=term,
                top=top_msgs,
                since_days=since_days
            ) or []
            candidatos_totales.extend(lote)
            print(f"[DIAN ZIP] candidatos búsqueda {term!r}: {len(lote)}")
        except Exception as e:
            print(f"[DIAN ZIP] Error buscando {term!r}: {e}")

    msgs = []
    seen = set()
    for m in candidatos_totales:
        mid = m.get("id")
        if not mid or mid in seen:
            continue
        seen.add(mid)
        msgs.append(m)

    if not msgs:
        print("[DIAN ZIP] ❌ No se encontraron mensajes candidatos.")
        print("[DIAN ZIP] ================ FIN BÚSQUEDA ZIP DIAN ================\n")
        return None, None, None, None

    contenedores = []
    for m in msgs:
        subj = m.get("subject") or ""
        subj_norm = normalize_text(subj)
        subj_alnum = _solo_alnum(subj)

        if _is_inbox_dian_container_subject(subj):
            contenedores.append(m)
            continue

        if ("dian" in subj_norm) or ("joyco" in subj_norm) or _contains_validacion(subj_norm):
            if any(tk and tk in subj_alnum for tk in tokens_obj):
                contenedores.append(m)

    dedup_cont = []
    seen_ids = set()
    for m in contenedores:
        mid = m.get("id")
        if mid and mid not in seen_ids:
            seen_ids.add(mid)
            dedup_cont.append(m)
    contenedores = dedup_cont

    print(f"[DIAN ZIP] contenedores filtrados: {len(contenedores)}")

    if not contenedores:
        print("[DIAN ZIP] ❌ No encontré correos contenedores DIAN/JOYCO.")
        print("[DIAN ZIP] ================ FIN BÚSQUEDA ZIP DIAN ================\n")
        return None, None, None, None

    revisados = 0
    target_pdf_clean = _solo_alnum(Path(target_pdf_name).stem)

    for m in contenedores:
        mid = m.get("id")
        subj = m.get("subject") or ""

        if not mid:
            continue

        print(f"\n[DIAN ZIP] Revisando contenedor: {subj} | id={mid}")

        try:
            zips = listar_adjuntos_zip(mid) or []
        except Exception as e:
            print(f"[DIAN ZIP] Error listando ZIPs en {mid}: {e}")
            zips = []

        print(f"[DIAN ZIP] ZIPs adjuntos en contenedor: {len(zips)}")

        if not zips:
            continue

        def _prio_zip(att: dict) -> int:
            zname = att.get("name") or ""
            clean = _solo_alnum(Path(zname).stem)
            p = 0
            if target_pdf_clean and target_pdf_clean in clean:
                p += 400
            if any(tk and tk in clean for tk in tokens_obj):
                p += 200
            return -p

        zips = sorted(zips, key=_prio_zip)

        for att in zips:
            revisados += 1
            aid = att.get("id")
            zname = att.get("name") or f"{aid}.zip"

            if not aid:
                continue

            tmp_zip = os.path.join(
                TMP_DIR,
                f"dian_zip_{uuid.uuid4().hex}_{re.sub(r'[^A-Za-z0-9_. -]', '_', zname)}"
            )

            ok = descargar_adjunto_por_id(mid, aid, tmp_zip)
            if not ok or not os.path.exists(tmp_zip):
                print(f"[DIAN ZIP] ⚠️ No pude descargar ZIP: {zname}")
                try:
                    if os.path.exists(tmp_zip):
                        os.remove(tmp_zip)
                except Exception:
                    pass
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
            print(f"[DIAN ZIP] ZIP revisado: {zname} | xmls_detectados={len(idents_xml)}")

            target_cufe = _norm_cufe(target_ident.get("CUFE") or "")
            if target_cufe:
                for ident_xml in idents_xml:
                    cufe_xml = _norm_cufe(ident_xml.get("CUFE") or "")
                    if cufe_xml and cufe_xml == target_cufe:
                        print(f"[DIAN ZIP] ✅ Match por CUFE en ZIP: {zname}")
                        print("[DIAN ZIP] ================ FIN BÚSQUEDA ZIP DIAN ================\n")
                        return zname, zip_bytes, mid, aid

            target_nums = _build_num_set(target_ident)
            for ident_xml in idents_xml:
                xml_nums = _build_num_set({
                    "NUMERO": ident_xml.get("NUMERO") or "",
                    "NUMERO_APROB": ident_xml.get("NUMERO") or "",
                })
                inter = target_nums & xml_nums
                if inter:
                    print(f"[DIAN ZIP] ✅ Match por número en ZIP: {zname} | inter={sorted(list(inter))[:5]}")
                    print("[DIAN ZIP] ================ FIN BÚSQUEDA ZIP DIAN ================\n")
                    return zname, zip_bytes, mid, aid

            zname_alnum = _solo_alnum(Path(zname).stem)

            if target_pdf_clean and target_pdf_clean in zname_alnum:
                print(f"[DIAN ZIP] ✅ Match por nombre ZIP: {zname}")
                print("[DIAN ZIP] ================ FIN BÚSQUEDA ZIP DIAN ================\n")
                return zname, zip_bytes, mid, aid

            for tk in tokens_obj:
                if tk and tk in zname_alnum:
                    print(f"[DIAN ZIP] ✅ Match por token en nombre ZIP: {zname}")
                    print("[DIAN ZIP] ================ FIN BÚSQUEDA ZIP DIAN ================\n")
                    return zname, zip_bytes, mid, aid

    print(f"[DIAN ZIP] ❌ No encontré ZIP coincidente dentro de contenedores DIAN. revisados={revisados}")
    print("[DIAN ZIP] ================ FIN BÚSQUEDA ZIP DIAN ================\n")
    return None, None, None, None


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


def _upload_replace_with_retries(
    local_path: str,
    sp_path: str,
    retries: int = 2,
    sleep_s: float = 1.5
) -> bool:
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


def _subir_excels_a_sharepoint(
    sp_excel_root: str,
    hubo_cambios_excel: bool,
    historial_actualizado: bool
):
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
            out.add(f"{stem}.xml")
            out.add(f"{stem}.pdf")
            out.add(f"{stem}.zip")
            out.add(f"{stem}.XML")
            out.add(f"{stem}.PDF")
            out.add(f"{stem}.ZIP")

    return out


def _try_workbook_append(
    sp_excel_root: str,
    archivos_ref: set[str],
    table_name: str = "TblFacturas"
) -> int:
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

def _registrar_desde_pdf_aprobado_fallback(
    *,
    msg_id: str,
    subj: str,
    pdf_name: str,
    pdf_tmp: str,
    ident_pdf: Dict[str, str],
    fecha_pdf: str,
    cufe_pdf: str,
    numero_aprob: str,
    fecha_local: str,
    hora_local: str,
    run_id: str,
    detalle_rows: List[Dict[str, object]],
    resumen: List[Tuple[str, float, str, int]],
    t0: float,
    usar_processed_store: bool,
    store: ProcessedStore,
    cufes_existentes: set,
    norm_cufes_existentes: set,
) -> Tuple[bool, int, int]:
    """
    Último recurso:
    si no hubo ZIP ni PDF externo y el PDF aprobado tiene CUFE válido,
    registrar directamente desde ese mismo PDF.
    Retorna: (aplico_fallback, total_nuevos, enriquecidas)
    """
    if not cufe_pdf or not _cufe_is_valid(cufe_pdf):
        print(f"⛔ Fallback PDF_APROBADAS no aplica para {pdf_name}: PDF sin CUFE válido.")
        return False, 0, 0

    print(f"✅ Fallback PDF_APROBADAS habilitado para {pdf_name} (CUFE válido).")

    try:
        reg = _generar_registro_pdf_only(pdf_tmp, pdf_name)

        numero_final = (
            ident_pdf.get("NUMERO_APROB")
            or ident_pdf.get("NUMERO")
            or reg.get("Número de factura")
            or ""
        )
        if numero_final and len(str(numero_final).strip()) >= 3:
            reg["Número de factura"] = str(numero_final).strip()

        total_nuevos = guardar_en_excel([reg])

        historial_actualizado = False
        if total_nuevos > 0:
            registrar_historial_por_zip([{
                "Fecha": fecha_local,
                "Hora": hora_local,
                "Archivo ZIP": "(PDF_APROBADAS_FALLBACK)",
                "Nuevos XML guardados": total_nuevos,
                "Errores encontrados": 0
            }])
            historial_actualizado = True

        enriquecidas = 0
        try:
            enriquecidas = sincronizar_aprobaciones_en_facturas()
        except Exception as e:
            print(f"[APROB] Error al sincronizar aprobaciones: {e}")

        print("☁️  Subiendo a SharePoint (fallback PDF aprobadas)...")
        sp_ext_root = f"{BASE_SP}/extraidos/pdf_aprobadas_fallback"
        sp_excel = f"{BASE_SP}/excel"

        sp_disponible = True
        try:
            ensure_folder(sp_ext_root)
            ensure_folder(sp_excel)
        except Exception as e:
            sp_disponible = False
            print(f"⚠️ SharePoint no disponible en fallback PDF aprobadas: {e}")

        if sp_disponible:
            try:
                upload_small_file(pdf_tmp, f"{sp_ext_root}/{os.path.basename(pdf_name)}", mode="skip")
            except Exception as e:
                print(f"⚠️ No pude subir PDF fallback a SharePoint: {e}")
        else:
            print("⚠️ Se omite subida a SharePoint por error temporal.")

        hubo_cambios_excel = (total_nuevos > 0) or (enriquecidas > 0)

        insertadas = 0
        if sp_disponible and hubo_cambios_excel:
            try:
                archivos_ref = {
                    os.path.basename(pdf_name),
                    str(reg.get("Archivo") or os.path.basename(pdf_name)),
                }
                insertadas = _try_workbook_append(sp_excel, archivos_ref, table_name="TblFacturas")
                print(f"✅ Workbook API (PDF_APROBADAS_FALLBACK): +{insertadas} fila(s) nuevas en TblFacturas.")
            except Exception as e:
                print(f"⚠️ Workbook API falló (PDF_APROBADAS_FALLBACK): {e}")
        elif hubo_cambios_excel:
            print("⚠️ Workbook API omitida por indisponibilidad temporal de SharePoint.")

        if sp_disponible:
            _subir_excels_a_sharepoint(sp_excel, hubo_cambios_excel, historial_actualizado)
        else:
            print("⚠️ No se suben excels a SharePoint en esta iteración.")

        if usar_processed_store:
            store.mark_processed(msg_id, {
                "status": "ok_pdf_aprobadas_fallback",
                "pdf": pdf_name,
                "nuevos": int(total_nuevos),
                "enriquecidas": int(enriquecidas),
                "cufe": cufe_pdf,
            })

        try:
            marcar_mensaje_como_leido(msg_id)
        except Exception as e:
            print(f"⚠️ No se pudo marcar como leído: {e}")

        secs = time.perf_counter() - t0
        resumen.append((pdf_name, secs, "fallback pdf aprobadas", total_nuevos))

        _push_detalle(
            detalle_rows, run_id, msg_id, subj,
            pdf_name=pdf_name,
            cufe=cufe_pdf,
            numero=reg.get("Número de factura") or "",
            fecha_factura=ident_pdf.get("FECHA") or fecha_pdf,
            zip_match="(PDF_APROBADAS_FALLBACK)",
            estado="ok_pdf_aprobadas_fallback",
            duracion_s=secs,
            nuevos=int(total_nuevos or 0),
            enriquecidas=int(enriquecidas or 0),
            fuente="PDF_APROBADAS"
        )

        if total_nuevos > 0 and cufe_pdf:
            cufes_existentes.add(cufe_pdf)
            norm_cufes_existentes.add(cufe_pdf)

        return True, int(total_nuevos or 0), int(enriquecidas or 0)

    except Exception as e:
        print(f"⚠️ Falló fallback PDF_APROBADAS para {pdf_name}: {e}")
        return False, 0, 0

def _procesar_pdf_aprobadas_como_ultimo_recurso(
    *,
    msg_id: str,
    subj: str,
    pdf_name: str,
    pdf_tmp: str,
    ident_pdf: Dict[str, str],
    fecha_pdf: str,
    fecha_local: str,
    hora_local: str,
    cufe_pdf: str,
    numero_aprob: str,
    detalle_rows: list,
    run_id: str,
    t0: float,
    usar_processed_store: bool,
    store,
):
    """
    Último último recurso:
    usa el MISMO PDF de solo aprobadas como fuente de registro,
    únicamente si el PDF tiene CUFE válido.

    Retorna:
    {
        "handled": bool,
        "ok": bool,
        "nuevos": int,
        "enriquecidas": int,
        "estado": str
    }
    """

    if not cufe_pdf or not _cufe_is_valid(cufe_pdf):
        estado = "sin_match_pdf_aprobadas_sin_cufe"
        print(f"❌ Fallback PDF_APROBADAS NO aplicado para {pdf_name}: el PDF no tiene CUFE válido.")

        if usar_processed_store:
            store.mark_processed(msg_id, {
                "status": estado,
                "pdf": pdf_name,
                "cufe": cufe_pdf or "",
            })

        _push_detalle(
            detalle_rows, run_id, msg_id, subj,
            pdf_name=pdf_name,
            cufe=cufe_pdf or "",
            numero=ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "",
            fecha_factura=ident_pdf.get("FECHA") or fecha_pdf,
            zip_match="(PDF_APROBADAS_FALLBACK_NO_CUFE)",
            estado=estado,
            duracion_s=(time.perf_counter() - t0),
            nuevos=0,
            enriquecidas=0,
            fuente="PDF_APROBADAS",
            error="No se pudo leer por solo PDF porque no se encontró CUFE válido"
        )

        return {
            "handled": True,
            "ok": False,
            "nuevos": 0,
            "enriquecidas": 0,
            "estado": estado,
        }

    print(f"✅ Fallback PDF_APROBADAS habilitado para {pdf_name} (CUFE válido).")

    reg = _generar_registro_pdf_only(pdf_tmp, pdf_name)

    numero_final = (
        ident_pdf.get("NUMERO_APROB")
        or ident_pdf.get("NUMERO")
        or reg.get("Número de factura")
        or ""
    )

    if numero_aprob and len(str(numero_aprob).strip()) >= 3:
        numero_final = str(numero_aprob).strip()

    if numero_final and len(str(numero_final).strip()) >= 3:
        reg["Número de factura"] = str(numero_final).strip()

    total_nuevos = guardar_en_excel([reg])

    historial_actualizado = False
    if total_nuevos > 0:
        registrar_historial_por_zip([{
            "Fecha": fecha_local,
            "Hora": hora_local,
            "Archivo ZIP": "(PDF_APROBADAS_FALLBACK)",
            "Nuevos XML guardados": total_nuevos,
            "Errores encontrados": 0
        }])
        historial_actualizado = True

    enriquecidas = 0
    try:
        enriquecidas = sincronizar_aprobaciones_en_facturas()
    except Exception as e:
        print(f"[APROB] Error al sincronizar aprobaciones: {e}")

    print("☁️  Subiendo a SharePoint (fallback PDF aprobadas)...")
    sp_ext_root = f"{BASE_SP}/extraidos/pdf_aprobadas_fallback"
    sp_excel = f"{BASE_SP}/excel"

    sp_disponible = True
    try:
        ensure_folder(sp_ext_root)
        ensure_folder(sp_excel)
    except Exception as e:
        sp_disponible = False
        print(f"⚠️ SharePoint no disponible en fallback PDF aprobadas: {e}")

    if sp_disponible:
        try:
            upload_small_file(pdf_tmp, f"{sp_ext_root}/{os.path.basename(pdf_name)}", mode="skip")
        except Exception as e:
            print(f"⚠️ No pude subir PDF fallback a SharePoint: {e}")
    else:
        print("⚠️ Se omite subida a SharePoint por error temporal.")

    hubo_cambios_excel = (total_nuevos > 0) or (enriquecidas > 0)

    insertadas = 0
    if sp_disponible and hubo_cambios_excel:
        try:
            archivos_ref = {
                os.path.basename(pdf_name),
                str(reg.get("Archivo") or os.path.basename(pdf_name)),
            }
            insertadas = _try_workbook_append(sp_excel, archivos_ref, table_name="TblFacturas")
            print(f"✅ Workbook API (PDF_APROBADAS_FALLBACK): +{insertadas} fila(s) nuevas en TblFacturas.")
        except Exception as e:
            print(f"⚠️ Workbook API falló (PDF_APROBADAS_FALLBACK): {e}")
    elif hubo_cambios_excel:
        print("⚠️ Workbook API omitida por indisponibilidad temporal de SharePoint.")

    if sp_disponible:
        _subir_excels_a_sharepoint(sp_excel, hubo_cambios_excel, historial_actualizado)
    else:
        print("⚠️ No se suben excels a SharePoint en esta iteración.")

    if usar_processed_store:
        store.mark_processed(msg_id, {
            "status": "ok_pdf_aprobadas_fallback",
            "pdf": pdf_name,
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
        numero=reg.get("Número de factura") or "",
        fecha_factura=ident_pdf.get("FECHA") or fecha_pdf,
        zip_match="(PDF_APROBADAS_FALLBACK)",
        estado="ok_pdf_aprobadas_fallback",
        duracion_s=(time.perf_counter() - t0),
        nuevos=int(total_nuevos or 0),
        enriquecidas=int(enriquecidas or 0),
        fuente="PDF_APROBADAS"
    )

    return {
        "handled": True,
        "ok": True,
        "nuevos": int(total_nuevos or 0),
        "enriquecidas": int(enriquecidas or 0),
        "estado": "ok_pdf_aprobadas_fallback",
    }


def run_desde_aprobadas(
    max_aprobados: int = 50,
    max_zip_buscar: int = 150,
    since_days: Optional[int] = None,
    unread_only: Optional[bool] = None,
    usar_processed_store: bool = True,
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

    modo_lectura = "solo NO leídos" if (unread_only is True or unread_only is None) else "todos"
    print(f"📬 Leyendo carpeta de aprobadas ({modo_lectura}): {APROB_FOLDER_NAME}")

    msgs = listar_mensajes_en_carpeta(
        folder_id,
        top=max_aprobados,
        unread_only=unread_only,
        since_days=since_days,
    )
    msgs_leidos = len(msgs) if msgs else 0

    if not msgs:
        print("ℹ️ No hay mensajes para procesar en la carpeta de aprobadas.")
        total_secs = time.perf_counter() - t0_total
        print(f"⏱️ Tiempo total real: {total_secs:.2f} s")
        try:
            lock.release()
        except Exception:
            pass
        return

    msgs_pendientes = []
    for m in msgs:
        mid = m.get("id")
        if not mid:
            continue

        if usar_processed_store:
            if not store.is_processed(mid):
                msgs_pendientes.append(m)
        else:
            msgs_pendientes.append(m)

    msgs_pendientes_count = len(msgs_pendientes)

    if not msgs_pendientes:
        print("✅ No hay mensajes nuevos para procesar según ProcessedStore.")
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

    idx_cufe, idx_num, idx_num_match = _build_zip_index(
        since_days=since_days,
        max_zip_buscar=max_zip_buscar,
        aidx=aidx
    )

    cufes_existentes = obtener_cufes_existentes()
    norm_cufes_existentes = {_norm_cufe(x) for x in cufes_existentes}
    print(f"ℹ️ CUFEs ya registrados en facturas.xlsx: {len(cufes_existentes)}")
    if len(cufes_existentes) == 0:
        print("⚠️ ALERTA: obtener_cufes_existentes() devolvió 0. Revisa si ARCHIVO_EXCEL apunta al archivo correcto o si el Excel está vacío.")
        print(f"⚠️ ARCHIVO_EXCEL={ARCHIVO_EXCEL}")

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

        if usar_processed_store and store.is_processed(msg_id):
            print(f"⏭️  Mensaje ya procesado (store). Se omite. id={msg_id}")
            continue

        pdf_atts = listar_adjuntos_pdf(msg_id)
        if not pdf_atts:
            if usar_processed_store:
                store.mark_processed(msg_id, {"status": "sin_pdf"})
            cnt_sin_pdf += 1
            msgs_procesados += 1

            _push_detalle(
                detalle_rows, run_id, msg_id, subj,
                estado="sin_pdf",
                duracion_s=(time.perf_counter() - t0)
            )
            continue

        pdf = None
        pdf_tmp = None
        ident_pdf = {}

        if len(pdf_atts) == 1:
            pdf = pdf_atts[0]
            pdf_name = pdf.get("name") or f"{pdf['id']}.pdf"
            pdf_tmp = os.path.join(TMP_DIR, pdf_name)

            if not descargar_adjunto_por_id(msg_id, pdf["id"], pdf_tmp):
                print(f"[APROB] No pude descargar PDF {pdf_name}")
                if usar_processed_store:
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
            ident_pdf = parse_identificadores_pdf(texto) or {}
        else:
            pdf, pdf_tmp, ident_pdf = _seleccionar_mejor_pdf(msg_id, subj, pdf_atts)
            if not pdf or not pdf_tmp:
                pdf = pdf_atts[0]
                pdf_name = pdf.get("name") or f"{pdf['id']}.pdf"
                pdf_tmp = os.path.join(TMP_DIR, pdf_name)

                if not descargar_adjunto_por_id(msg_id, pdf["id"], pdf_tmp):
                    print(f"[APROB] No pude descargar PDF (fallback) {pdf_name}")
                    if usar_processed_store:
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
                ident_pdf = parse_identificadores_pdf(texto) or {}

        pdf_name = pdf.get("name") or f"{pdf['id']}.pdf"

        subj_num = _numero_from_subject(subj)
        best_num = _prefer_subject_numero(ident_pdf.get("NUMERO"), subj_num)
        if best_num:
            ident_pdf["NUMERO"] = best_num

        numero_aprob = (ident_pdf.get("NUMERO_APROB") or "").strip()
        if not numero_aprob and subj_num and subj_num.strip() and subj_num.strip() != (ident_pdf.get("NUMERO") or "").strip():
            numero_aprob = subj_num.strip()

        if numero_aprob:
            ident_pdf["NUMERO_APROB"] = numero_aprob

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
        print(f"→ NUMERO_APROB detectado: {ident_pdf.get('NUMERO_APROB')}")
        print(f"→ FECHA detectada: {ident_pdf.get('FECHA')}")
        print("===========================\n")

        numero_principal = _elegir_numero_principal(ident_pdf, subj, pdf_name)

        if numero_principal and not ident_pdf.get("NUMERO_APROB"):
            ident_pdf["NUMERO_APROB"] = numero_principal

        if not _es_probable_factura_electronica(subj, pdf_name, ident_pdf):
            print(f"⏭️ Se omite PDF no probable factura electrónica: {pdf_name}")
            secs = time.perf_counter() - t0
            resumen.append((pdf_name, secs, "omitido no factura", 0))

            if usar_processed_store:
                store.mark_processed(msg_id, {
                    "status": "omitido_no_factura",
                    "pdf": pdf_name,
                    "cufe": cufe_pdf,
                })

            _push_detalle(
                detalle_rows, run_id, msg_id, subj,
                pdf_name=pdf_name,
                cufe=cufe_pdf,
                numero=ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "",
                fecha_factura=fecha_pdf,
                estado="omitido_no_factura",
                duracion_s=(time.perf_counter() - t0),
                fuente="FILTRO"
            )

            msgs_procesados += 1
            procesados += 1
            continue

        if _is_dian_trigger_message(msg):
            print(f"[DIAN] Detectado mensaje DIAN/JOYCO-validación en aprobadas (asunto/cuerpo): {subj!r}")

            pdf_real_path, mid_src, aid_src = _buscar_pdf_en_correo_validaciones_dian(
                target_ident=ident_pdf,
                target_pdf_name=pdf_name,
                since_days=since_days,
                top_msgs=400
            )

            zip_dian_name = None
            zip_dian_bytes = None
            zip_mid_src = None
            zip_aid_src = None

            if not pdf_real_path:
                zip_dian_name, zip_dian_bytes, zip_mid_src, zip_aid_src = _buscar_zip_en_correo_validaciones_dian(
                    target_ident=ident_pdf,
                    target_pdf_name=pdf_name,
                    since_days=since_days,
                    top_msgs=400
                )

            # -------------------------------------------------
            # 1) Si encontró PDF externo DIAN -> procesa PDF-only DIAN
            # -------------------------------------------------
            if pdf_real_path:
                pdf_real_name = os.path.basename(pdf_real_path)
                reg = _generar_registro_pdf_only(pdf_real_path, pdf_real_name)

                if numero_aprob and len(numero_aprob) >= 3:
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

                sp_disponible = True
                try:
                    ensure_folder(sp_ext_root)
                    ensure_folder(sp_excel)
                except Exception as e:
                    sp_disponible = False
                    print(f"⚠️ SharePoint no disponible en rama DIAN: {e}")

                if sp_disponible:
                    try:
                        upload_small_file(pdf_real_path, f"{sp_ext_root}/{pdf_real_name}", mode="skip")
                    except Exception as e:
                        print(f"[DIAN] No pude subir PDF real: {e}")
                else:
                    print("[DIAN] ⚠️ Se omite subida a SharePoint por error temporal.")

                hubo_cambios_excel = (total_nuevos > 0) or (enriquecidas > 0)

                insertadas = 0
                if sp_disponible and hubo_cambios_excel:
                    try:
                        archivos_ref = {str(reg.get("Archivo") or pdf_real_name)}
                        insertadas = _try_workbook_append(sp_excel, archivos_ref, table_name="TblFacturas")
                        print(f"✅ Workbook API (DIAN/PDF): +{insertadas} fila(s) nuevas en TblFacturas.")
                    except Exception as e:
                        print(f"⚠️ Workbook API falló (DIAN/PDF): {e}")

                if sp_disponible:
                    _subir_excels_a_sharepoint(sp_excel, hubo_cambios_excel, historial_actualizado)
                else:
                    print("[DIAN] ⚠️ No se suben excels a SharePoint en esta iteración.")

                if usar_processed_store:
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
                resumen.append((pdf_name, secs, "match dian pdf", total_nuevos))

                _push_detalle(
                    detalle_rows, run_id, msg_id, subj,
                    pdf_name=pdf_name,
                    cufe=cufe_pdf,
                    numero=reg.get("Número de factura") or "",
                    fecha_factura=ident_pdf.get("FECHA") or fecha_pdf,
                    zip_match="(PDF-ONLY) VALIDACIONES DIAN",
                    estado="ok_dian_pdf_only",
                    duracion_s=secs,
                    nuevos=int(total_nuevos or 0),
                    enriquecidas=int(enriquecidas or 0),
                    fuente="DIAN_PDF"
                )

                cnt_dian += 1
                msgs_procesados += 1
                nuevos_total += int(total_nuevos or 0)
                enriq_total += int(enriquecidas or 0)

                sin_match_consec = 0
                sin_nuevos_consec = 0 if total_nuevos > 0 else (sin_nuevos_consec + 1)
                procesados += 1
                continue

            # -------------------------------------------------
            # 2) Si encontró ZIP externo DIAN -> procesa ZIP DIAN
            # -------------------------------------------------
            if zip_dian_name and zip_dian_bytes:
                b1 = _limpiar_adj_hoy()
                if b1:
                    print(f"🧹 Limpieza ADJ_HOY: borrados {b1} ZIP(s) viejos.")

                b2 = _limpiar_ext_hoy()
                if b2:
                    print(f"🧹 Limpieza EXT_HOY: borrados {b2} elemento(s) viejos.")

                found_zip_name = zip_dian_name
                found_zip_bytes = zip_dian_bytes

                zip_local_path = Path(ADJ_HOY) / found_zip_name
                with open(zip_local_path, "wb") as f:
                    f.write(found_zip_bytes)

                print(f"🗜️  Extrayendo ZIP DIAN {found_zip_name} ...")
                resultados = extraer_por_zip(ADJ_HOY, EXT_HOY)
                print("🧾 Procesando XMLs DIAN...")

                historial_rows = []
                total_nuevos = 0
                carpeta_obj = None
                ruta_obj = None
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
                            if old != numero_aprob and len(numero_aprob) >= 3:
                                dct["Número de factura"] = numero_aprob

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

                print(f"✅ Excel local actualizado por ZIP DIAN (+{total_nuevos}): {ARCHIVO_EXCEL}")

                historial_actualizado = False
                if historial_rows:
                    registrar_historial_por_zip(historial_rows)
                    historial_actualizado = True

                enriquecidas = 0
                try:
                    enriquecidas = sincronizar_aprobaciones_en_facturas()
                except Exception as e:
                    print(f"[APROB] Error al sincronizar aprobaciones: {e}")

                print("☁️  Subiendo a SharePoint (DIAN / ZIP)...")
                if USE_DATE_SUBFOLDERS:
                    sp_adj_root = f"{BASE_SP}/adjuntos/{fecha_local}"
                    sp_ext_root = f"{BASE_SP}/extraidos/{fecha_local}"
                else:
                    sp_adj_root = f"{BASE_SP}/adjuntos"
                    sp_ext_root = f"{BASE_SP}/extraidos"
                sp_excel = f"{BASE_SP}/excel"

                sp_disponible = True
                try:
                    ensure_folder(sp_adj_root)
                    ensure_folder(sp_ext_root)
                    ensure_folder(sp_excel)
                except Exception as e:
                    sp_disponible = False
                    print(f"⚠️ SharePoint no disponible en rama DIAN ZIP: {e}")

                if sp_disponible:
                    try:
                        upload_small_file(str(zip_local_path), f"{sp_adj_root}/{found_zip_name}", mode="skip")
                    except Exception as e:
                        print(f"⚠️ Error subiendo ZIP DIAN a SharePoint: {e}")

                    try:
                        if carpeta_obj and ruta_obj and os.path.exists(ruta_obj):
                            upload_directory(ruta_obj, f"{sp_ext_root}/{carpeta_obj}", mode="skip")
                        else:
                            upload_directory(EXT_HOY, sp_ext_root, mode="skip")
                    except Exception as e:
                        print(f"⚠️ Error subiendo extraídos DIAN a SharePoint: {e}")
                else:
                    print("⚠️ Se omite subida a SharePoint para ZIP DIAN por error temporal.")

                hubo_cambios_excel = (total_nuevos > 0) or (enriquecidas > 0)

                insertadas = 0
                if sp_disponible and hubo_cambios_excel:
                    try:
                        archivos_xml = set()
                        if ruta_obj and os.path.isdir(ruta_obj):
                            for fn in os.listdir(ruta_obj):
                                if fn.lower().endswith(".xml"):
                                    archivos_xml.add(fn)

                        archivos_ref = set(archivos_realmente_guardados)
                        archivos_ref |= set(archivos_xml)
                        archivos_ref.add(os.path.basename(pdf_name))
                        if found_zip_name:
                            archivos_ref.add(os.path.basename(found_zip_name))

                        insertadas = _try_workbook_append(sp_excel, archivos_ref, table_name="TblFacturas")
                        print(f"✅ Workbook API (DIAN/ZIP): +{insertadas} fila(s) nuevas en TblFacturas.")
                    except Exception as e:
                        print(f"⚠️ Workbook API falló (DIAN/ZIP): {e}")
                elif hubo_cambios_excel:
                    print("⚠️ Workbook API DIAN/ZIP omitida por indisponibilidad temporal de SharePoint.")

                if sp_disponible:
                    _subir_excels_a_sharepoint(sp_excel, hubo_cambios_excel, historial_actualizado)
                else:
                    print("⚠️ No se suben excels a SharePoint en esta iteración DIAN/ZIP.")

                if usar_processed_store:
                    store.mark_processed(msg_id, {
                        "status": "ok_dian_zip",
                        "pdf": pdf_name,
                        "zip": found_zip_name,
                        "nuevos": int(total_nuevos),
                        "enriquecidas": int(enriquecidas),
                        "cufe": cufe_pdf,
                        "src_msg": zip_mid_src,
                        "src_att": zip_aid_src,
                    })

                try:
                    marcar_mensaje_como_leido(msg_id)
                except Exception as e:
                    print(f"⚠️ No se pudo marcar como leído: {e}")

                secs = time.perf_counter() - t0
                resumen.append((pdf_name, secs, "match dian zip", total_nuevos))

                _push_detalle(
                    detalle_rows, run_id, msg_id, subj,
                    pdf_name=pdf_name,
                    cufe=cufe_pdf,
                    numero=ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "",
                    fecha_factura=fecha_pdf,
                    zip_match=found_zip_name,
                    estado="ok_dian_zip",
                    duracion_s=secs,
                    nuevos=int(total_nuevos or 0),
                    enriquecidas=int(enriquecidas or 0),
                    fuente="DIAN_ZIP"
                )

                cnt_dian += 1
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
                continue

            # -------------------------------------------------
            # 3) Si NO encontró ni PDF DIAN ni ZIP DIAN,
            # aplicar también el fallback final del mismo PDF aprobado
            # -------------------------------------------------
            print(f"[DIAN] No encontré PDF/ZIP externo para {pdf_name}. Intentando fallback con el mismo PDF aprobado...")

            aplico_fallback, total_nuevos, enriquecidas = _registrar_desde_pdf_aprobado_fallback(
                msg_id=msg_id,
                subj=subj,
                pdf_name=pdf_name,
                pdf_tmp=pdf_tmp,
                ident_pdf=ident_pdf,
                fecha_pdf=fecha_pdf,
                cufe_pdf=cufe_pdf,
                numero_aprob=numero_aprob,
                fecha_local=fecha_local,
                hora_local=hora_local,
                run_id=run_id,
                detalle_rows=detalle_rows,
                resumen=resumen,
                t0=t0,
                usar_processed_store=usar_processed_store,
                store=store,
                cufes_existentes=cufes_existentes,
                norm_cufes_existentes=norm_cufes_existentes,
            )

            if aplico_fallback:
                cnt_dian += 1
                msgs_procesados += 1
                nuevos_total += int(total_nuevos or 0)
                enriq_total += int(enriquecidas or 0)

                sin_match_consec = 0
                if total_nuevos == 0:
                    sin_nuevos_consec += 1
                else:
                    sin_nuevos_consec = 0

                procesados += 1
                continue

            # -------------------------------------------------
            # 4) Si tampoco aplicó fallback, cerrar como sin_match_dian
            # -------------------------------------------------
            secs = time.perf_counter() - t0
            resumen.append((pdf_name, secs, "sin match dian", 0))

            motivo_dian = "sin_match_dian"
            if not cufe_pdf:
                motivo_dian = "sin_match_dian_pdf_sin_cufe"
            elif cufe_pdf and not _cufe_is_valid(cufe_pdf):
                motivo_dian = "sin_match_dian_cufe_debil"

            if usar_processed_store:
                store.mark_processed(msg_id, {
                    "status": motivo_dian,
                    "pdf": pdf_name,
                    "cufe": cufe_pdf
                })

            cnt_sin_match += 1
            msgs_procesados += 1
            sin_match_consec += 1
            sin_nuevos_consec = 0
            procesados += 1

            _push_detalle(
                detalle_rows, run_id, msg_id, subj,
                pdf_name=pdf_name,
                cufe=cufe_pdf,
                numero=ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "",
                fecha_factura=fecha_pdf,
                estado=motivo_dian,
                duracion_s=secs,
                fuente="DIAN"
            )
            continue

        found_match = False
        found_zip_name = None
        found_zip_bytes = None
        fuente_match = "normal"

        if cufe_pdf and cufe_pdf in idx_cufe:
            found_zip_name, found_zip_bytes = idx_cufe[cufe_pdf]
            found_match = True
            fuente_match = "CUFE"

        if not found_match:
            num_pdf = ident_pdf.get("NUMERO") or ""
            num_aprob = ident_pdf.get("NUMERO_APROB") or ""
            num_asunto = subj_num or ""
            num_principal = numero_principal or ""

            found_zip_name, found_zip_bytes, variante = _buscar_zip_por_numero_match(
                idx_num_match,
                num_aprob,
                num_asunto,
                num_principal,
                num_pdf,
            )
            if found_zip_name and found_zip_bytes:
                found_match = True
                fuente_match = f"NUM_MATCH:{variante}"

        if not found_match:
            num_pdf = ident_pdf.get("NUMERO") or ""
            num_aprob = ident_pdf.get("NUMERO_APROB") or ""
            num_asunto = subj_num or ""

            found_zip_name, found_zip_bytes, variante = _buscar_zip_por_numero(
                idx_num,
                num_aprob,
                num_asunto,
                num_pdf,
            )
            if found_zip_name and found_zip_bytes:
                found_match = True
                fuente_match = f"NUMERO:{variante}"

        if not found_match and not _is_uuid_like_name(pdf_name):
            for tk in _tokens_match_from_text(Path(pdf_name).stem):
                if not _token_es_util_para_match(tk):
                    continue
                found_zip_name, found_zip_bytes, variante = _buscar_zip_por_numero_match(
                    idx_num_match, tk
                )
                if found_zip_name and found_zip_bytes:
                    found_match = True
                    fuente_match = f"PDF_TOKEN:{variante}"
                    break

        if not found_match:
            pdf_base = Path(pdf_name).stem.lower()

            if not _is_uuid_like_name(pdf_name):
                pdf_clean = re.sub(r"[^a-z0-9]", "", pdf_base)

                vistos = set()
                for zn, zbytes in list(idx_cufe.values()) + list(idx_num.values()) + list(idx_num_match.values()):
                    if zn in vistos:
                        continue
                    vistos.add(zn)

                    zbase = Path(zn).stem.lower()
                    zclean = re.sub(r"[^a-z0-9]", "", zbase)

                    if len(pdf_clean) >= 8 and (pdf_clean == zclean or pdf_clean in zclean or zclean in pdf_clean):
                        found_zip_name, found_zip_bytes = zn, zbytes
                        found_match = True
                        fuente_match = "NOMBRE"
                        print(f"🔄 Emparejado por nombre: {pdf_name} ↔ {zn}")
                        break

        print("\n[MATCH DEBUG] =====================================")
        print(f"[MATCH DEBUG] PDF: {pdf_name}")
        print(f"[MATCH DEBUG] ASUNTO: {subj}")
        print(f"[MATCH DEBUG] CUFE PDF: {cufe_pdf}")
        print(f"[MATCH DEBUG] FECHA PDF: {fecha_pdf}")
        print(f"[MATCH DEBUG] NUMERO PDF: {ident_pdf.get('NUMERO')}")
        print(f"[MATCH DEBUG] NUMERO_APROB: {ident_pdf.get('NUMERO_APROB')}")
        print(f"[MATCH DEBUG] NUMERO PRINCIPAL: {numero_principal}")
        print(f"[MATCH DEBUG] SUBJECT NUM: {subj_num}")
        print(f"[MATCH DEBUG] found_match={found_match}")
        print(f"[MATCH DEBUG] found_zip_name={found_zip_name}")
        print(f"[MATCH DEBUG] fuente_match={fuente_match}")
        print(f"[MATCH DEBUG] idx_cufe_size={len(idx_cufe)} | idx_num_size={len(idx_num)} | idx_num_match_size={len(idx_num_match)}")

        if cufe_pdf:
            print(f"[MATCH DEBUG] cufe_pdf_en_idx_cufe={cufe_pdf in idx_cufe}")

        for base_num in [
            ident_pdf.get("NUMERO_APROB") or "",
            subj_num or "",
            numero_principal or "",
            ident_pdf.get("NUMERO") or "",
        ]:
            if base_num:
                print(f"[MATCH DEBUG] numero base={base_num}")
                print(f"[MATCH DEBUG] variantes numero={_numero_variantes(base_num)}")
                print(f"[MATCH DEBUG] variantes match={_variantes_match_numero(base_num)}")
                _debug_top_similares_idx(idx_num, base_num, limite=10)

        print("[MATCH DEBUG] =====================================\n")

        if not found_match or not found_zip_name or not found_zip_bytes:
            entry = None

            if cufe_pdf:
                try:
                    entry = aidx.find_zip_by_cufe(cufe_pdf)
                except Exception:
                    entry = None

            if not entry:
                try:
                    num_pdf = ident_pdf.get("NUMERO") or ""
                    num_aprob = ident_pdf.get("NUMERO_APROB") or ""
                    num_asunto = subj_num or ""

                    for n in [num_aprob, num_asunto, numero_principal, num_pdf]:
                        if not n:
                            continue

                        print(f"[AIDX DEBUG BUSQ] buscando por numero base={n}")

                        variantes = _numero_variantes(n)
                        variantes_match = _variantes_match_numero(n)
                        todas = []

                        for x in variantes + variantes_match + [n]:
                            if x and x not in todas:
                                todas.append(x)

                        print(f"[AIDX DEBUG BUSQ] variantes={todas}")

                        for vn in todas:
                            entry = aidx.find_zip_by_numero(vn)
                            if entry:
                                print(f"[AIDX DEBUG BUSQ] MATCH AIDX por numero={vn} -> {entry.get('att_name')}")
                                break

                        if entry:
                            break
                except Exception as e:
                    print(f"[AIDX DEBUG BUSQ] error buscando en AIDX: {e}")
                    entry = None

            if entry:
                try:
                    zname = entry.get("att_name") or "factura.zip"
                    mid = entry.get("msg_id")
                    aid = entry.get("att_id")

                    if mid and aid:
                        print(f"🧠 [AIDX] Encontré ZIP histórico: {zname} (descargando directo por IDs)...")
                        tmp_zip = os.path.join(TMP_DIR, f"aidx_{re.sub(r'[^A-Za-z0-9_. -]', '_', zname)}")
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
            print("🔁 Intentando último recurso con el MISMO PDF de solo aprobadas...")

            try:
                resultado_fallback = _procesar_pdf_aprobadas_como_ultimo_recurso(
                    msg_id=msg_id,
                    subj=subj,
                    pdf_name=pdf_name,
                    pdf_tmp=pdf_tmp,
                    ident_pdf=ident_pdf,
                    fecha_pdf=fecha_pdf,
                    fecha_local=fecha_local,
                    hora_local=hora_local,
                    cufe_pdf=cufe_pdf,
                    numero_aprob=numero_aprob,
                    detalle_rows=detalle_rows,
                    run_id=run_id,
                    t0=t0,
                    usar_processed_store=usar_processed_store,
                    store=store,
                )

                if resultado_fallback["handled"]:
                    secs = time.perf_counter() - t0

                    if resultado_fallback["ok"]:
                        resumen.append((pdf_name, secs, "fallback pdf aprobadas", resultado_fallback["nuevos"]))

                        cnt_ok += 1
                        msgs_procesados += 1
                        nuevos_total += int(resultado_fallback["nuevos"] or 0)
                        enriq_total += int(resultado_fallback["enriquecidas"] or 0)

                        sin_match_consec = 0
                        if int(resultado_fallback["nuevos"] or 0) == 0:
                            sin_nuevos_consec += 1
                        else:
                            sin_nuevos_consec = 0
                            if cufe_pdf:
                                cufes_existentes.add(cufe_pdf)
                                norm_cufes_existentes.add(cufe_pdf)

                        procesados += 1
                        continue

                    else:
                        resumen.append((pdf_name, secs, "sin match pdf aprobadas sin cufe", 0))

                        cnt_sin_match += 1
                        msgs_procesados += 1
                        sin_match_consec += 1
                        sin_nuevos_consec = 0
                        procesados += 1

                        if (procesados >= AUTO_STOP_MIN_PROCESADOS) and (sin_match_consec >= AUTO_STOP_SIN_MATCH_CONSEC):
                            print("🛑 Deteniendo flujo: varios PDFs consecutivos sin match.")
                            break
                        continue

            except Exception as e:
                print(f"⚠️ Falló el último recurso PDF_APROBADAS para {pdf_name}: {e}")

            secs = time.perf_counter() - t0
            resumen.append((pdf_name, secs, "sin match", 0))

            motivo_sin_match = "sin_match"
            num_candidato = ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or ""

            if cufe_pdf and not _cufe_is_valid(cufe_pdf):
                motivo_sin_match = "sin_match_pdf_sin_cufe"
            elif not cufe_pdf:
                motivo_sin_match = "sin_match_pdf_sin_cufe"
            elif not num_candidato and not cufe_pdf:
                motivo_sin_match = "sin_match_sin_identificadores"
            elif num_candidato and not _numero_parece_valido(num_candidato):
                motivo_sin_match = "sin_match_numero_no_valido"
            else:
                motivo_sin_match = "sin_match_no_zip"

            if usar_processed_store:
                store.mark_processed(msg_id, {
                    "status": motivo_sin_match,
                    "pdf": pdf_name,
                    "cufe": cufe_pdf
                })

            _push_detalle(
                detalle_rows, run_id, msg_id, subj,
                pdf_name=pdf_name,
                cufe=cufe_pdf,
                numero=ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "",
                fecha_factura=fecha_pdf,
                estado=motivo_sin_match,
                duracion_s=(time.perf_counter() - t0),
                fuente=fuente_match,
                error="No encontró ZIP, no encontró DIAN y no pudo cerrarse por fallback PDF_APROBADAS"
            )

            cnt_sin_match += 1
            msgs_procesados += 1

            sin_match_consec += 1
            sin_nuevos_consec = 0
            procesados += 1

            if (procesados >= AUTO_STOP_MIN_PROCESADOS) and (sin_match_consec >= AUTO_STOP_SIN_MATCH_CONSEC):
                print("🛑 Deteniendo flujo: varios PDFs consecutivos sin match.")
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
                    if old != numero_aprob and len(numero_aprob) >= 3:
                        dct["Número de factura"] = numero_aprob

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

        sp_disponible = True
        try:
            ensure_folder(sp_adj_root)
            ensure_folder(sp_ext_root)
            ensure_folder(sp_excel)
        except Exception as e:
            sp_disponible = False
            print(f"⚠️ SharePoint no disponible en este momento: {e}")

        if sp_disponible:
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
        else:
            print("⚠️ Se omite subida a SharePoint para este correo por error temporal.")

        hubo_cambios_excel = (total_nuevos > 0) or (enriquecidas > 0)

        insertadas = 0
        if sp_disponible and hubo_cambios_excel:
            try:
                archivos_xml = set()
                if ruta_obj and os.path.isdir(ruta_obj):
                    for fn in os.listdir(ruta_obj):
                        if fn.lower().endswith(".xml"):
                            archivos_xml.add(fn)

                archivos_ref = set(archivos_realmente_guardados)
                archivos_ref |= set(archivos_xml)
                archivos_ref.add(os.path.basename(pdf_name))
                if found_zip_name:
                    archivos_ref.add(os.path.basename(found_zip_name))

                insertadas = _try_workbook_append(sp_excel, archivos_ref, table_name="TblFacturas")
                print(f"✅ Workbook API: +{insertadas} fila(s) nuevas en TblFacturas.")
            except Exception as e:
                print(f"⚠️ Workbook API falló: {e}")
        elif hubo_cambios_excel:
            print("⚠️ Workbook API omitida por indisponibilidad temporal de SharePoint.")

        if sp_disponible:
            _subir_excels_a_sharepoint(sp_excel, hubo_cambios_excel, historial_actualizado)
        else:
            print("⚠️ No se suben excels a SharePoint en esta iteración.")

        print("🎉 Proceso por aprobadas finalizado para:", found_zip_name)
        secs = time.perf_counter() - t0
        resumen.append((pdf_name, secs, "match", total_nuevos))

        if usar_processed_store:
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
            numero=ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "",
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
            print("🛑 Deteniendo flujo: varios PDFs con match pero sin nuevos registros.")
            break

    try:
        n = borrar_pdfs_en_arbol(TMP_DIR)
        print(f"🧹 Limpieza temp_check: borrados {n} PDF(s).")
    except Exception:
        print("⚠️ Limpieza temp_check: no se pudo completar.")

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
                "nota": "",
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