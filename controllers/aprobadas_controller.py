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


CONCEPTOS_BASE_FIJOS = (
    "Subtotal",
    "IVA 5%",
    "IVA 19%",
    "Retención de IVA",
    "Retención de ICA",
    "Retención en la fuente",
    "Total",
)


def _float_seguro(valor) -> float:
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        try:
            return float(valor)
        except Exception:
            return 0.0

    s = str(valor).strip()
    if not s:
        return 0.0

    s = s.replace("\xa0", " ").replace("$", "").replace("COP", "")
    s = s.replace(" ", "")

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")

    try:
        return float(s)
    except Exception:
        return 0.0
def _forzar_texto_numero_factura(valor) -> str:
    """
    Fuerza el número de factura a texto para evitar notación científica
    o conversión rara al subir filas al Excel web.
    """
    if valor is None:
        return ""

    if isinstance(valor, str):
        s = valor.strip()
    else:
        s = str(valor).strip()

    if not s:
        return ""

    s_upper = s.upper().replace(",", ".")

    if "E+" in s_upper or "E-" in s_upper:
        try:
            num = float(s_upper)
            return "{:.0f}".format(num)
        except Exception:
            return s

    if re.fullmatch(r"\d+\.0", s):
        try:
            return str(int(float(s)))
        except Exception:
            return s

    return s

def _asegurar_reg_7_conceptos(reg: Dict[str, object]) -> Dict[str, object]:
    """
    Garantiza que cada registro registrable tenga SIEMPRE los 7 conceptos base
    que usa guardar_en_excel para expandir la factura.
    """
    if not reg:
        reg = {}

    out = dict(reg)

    # Campos de texto que no deberían faltar nunca
    for k in (
        "Archivo", "Empresa emisora", "CUFE", "Ciudad emisora", "Código ciudad",
        "NIT", "Cliente", "Número de factura", "Año", "Mes", "Día",
        "Tipo de contribuyente", "Actividad económica", "DescripcionLineas",
        "Radicado", "ProyectoProceso",
    ):
        out.setdefault(k, "")

    # Conceptos numéricos fijos
    for concepto in CONCEPTOS_BASE_FIJOS:
        out[concepto] = _float_seguro(out.get(concepto, 0.0))

    return out


def _asegurar_regs_registrables_7_conceptos(regs: Optional[List[Dict[str, object]]]) -> List[Dict[str, object]]:
    if not regs:
        return []

    salida: List[Dict[str, object]] = []
    for reg in regs:
        salida.append(_asegurar_reg_7_conceptos(reg or {}))
    return salida


def _normalizar_numero_para_match_local(s: str) -> str:
    return _solo_alnum(_normalizar_numero_match(s or ""))


def _forzar_campos_obligatorios_en_excel_local(
    excel_path: str,
    *,
    archivos_ref: Optional[set[str]] = None,
    numeros_ref: Optional[set[str]] = None,
    radicado: str = "",
    proyecto: str = "",
) -> Tuple[int, int]:
    """
    Fuerza Radicado y ProyectoProceso en el Excel local para TODAS las filas
    recién guardadas de la factura actual.
    """
    if not excel_path or not os.path.exists(excel_path):
        return 0, 0

    radicado = str(radicado or "").strip()
    proyecto = str(proyecto or "").strip()
    if not radicado and not proyecto:
        return 0, 0

    archivos_ref = {str(x).strip() for x in (archivos_ref or set()) if str(x).strip()}
    archivos_ref = _expand_archivos_ref(archivos_ref) if archivos_ref else set()

    nums_norm = set()
    for n in (numeros_ref or set()):
        nn = _normalizar_numero_para_match_local(str(n or "").strip())
        if nn:
            nums_norm.add(nn)

    if not archivos_ref and not nums_norm:
        return 0, 0

    try:
        from openpyxl import load_workbook

        wb = load_workbook(excel_path)
        ws = wb["Facturas"] if "Facturas" in wb.sheetnames else wb.active

        headers = {}
        for c in range(1, ws.max_column + 1):
            key = str(ws.cell(row=1, column=c).value or "").strip()
            if key:
                headers[key] = c

        col_arc = headers.get("Archivo")
        col_num = headers.get("Número de factura")
        col_rad = headers.get("Radicado")
        col_pro = headers.get("ProyectoProceso")

        if not col_rad or not col_pro:
            return 0, 0

        filas_match = 0
        filas_upd = 0

        for r in range(2, ws.max_row + 1):
            arc_val = str(ws.cell(r, col_arc).value or "").strip() if col_arc else ""
            num_val = str(ws.cell(r, col_num).value or "").strip() if col_num else ""

            arc_hit = False
            num_hit = False

            if archivos_ref and arc_val:
                arc_hit = (arc_val in archivos_ref) or (os.path.basename(arc_val) in archivos_ref)

            if nums_norm and num_val:
                num_hit = _normalizar_numero_para_match_local(num_val) in nums_norm

            if not arc_hit and not num_hit:
                continue

            filas_match += 1

            cell_rad = ws.cell(r, col_rad)
            cell_pro = ws.cell(r, col_pro)
            cambio = False

            if radicado and not str(cell_rad.value or "").strip():
                cell_rad.value = radicado
                cambio = True

            if proyecto and not str(cell_pro.value or "").strip():
                cell_pro.value = proyecto
                cambio = True

            if cambio:
                filas_upd += 1

        if filas_upd > 0:
            wb.save(excel_path)

        return filas_match, filas_upd

    except Exception as e:
        print(f"[LOCAL FORCE] Error forzando campos obligatorios en Excel local: {e}")
        return 0, 0


ADJ_HOY = os.path.join(DATA_DIR, "adjuntos", "hoy")
EXT_HOY = os.path.join(DATA_DIR, "extraidos", "hoy")

USE_DATE_SUBFOLDERS = False

_CTRL_REGEX = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
_AMP_FIX = re.compile(r"&(?!(?:[a-zA-Z]+|#\d+|#x[0-9A-Fa-f]+);)")
_NON_INVOICE_PREFIXES = {"DDI", "RAD", "RDI", "RDC", "REC", "RCP", "DOC", "REF"}


def _clasificar_auditoria_detalle(
    *,
    estado: str = "",
    nuevos: int = 0,
    pdf_name: str = "",
    subj: str = "",
    error: str = "",
) -> Tuple[str, int, str]:
    """
    Clasificación operativa para audit_detalle SIN tocar la lógica de negocio
    ya aprobada de nuevos/enriquecidas.
    """
    estado_norm = str(estado or "").strip().lower()
    nuevos_i = int(nuevos or 0)
    s = normalize_text(f"{subj or ''} {pdf_name or ''} {error or ''}")

    if nuevos_i > 0:
        if estado_norm == "ok_dian_pdf_only":
            return "DIAN_PDF_ONLY_REGISTRADA", nuevos_i, ""
        if estado_norm == "ok_dian_zip":
            return "DIAN_ZIP_REGISTRADA", nuevos_i, ""
        if estado_norm == "ok_pdf_aprobadas_fallback":
            return "OK_REGISTRADA_FALLBACK_PDF", nuevos_i, ""
        if estado_norm == "ok":
            return "OK_REGISTRADA", nuevos_i, ""
        return "REGISTRADA", nuevos_i, ""

    # nuevos == 0
    if estado_norm in {"ok", "ok_pdf_aprobadas_fallback", "ok_dian_pdf_only", "ok_dian_zip", "ok_registro_minimo"}:
        motivo = "SIN_DATOS_REGISTRABLES"
        if any(x in s for x in ["etb", "enel", "tigo", "claro", "movistar", "gas", "energia", "energía", "acueducto", "telefonia", "telefonía"]):
            motivo = "SERVICIO_PUBLICO"
        elif any(x in s for x in ["tv", "television", "televisión", "internet", "celular", "plan movil", "plan móvil"]):
            motivo = "SERVICIO_TECNOLOGIA"
        return "OK_NO_REGISTRABLE", 0, motivo

    if estado_norm == "omitido_no_factura":
        return "OMITIDO_NO_FACTURA", 0, "NO_FACTURA"

    if estado_norm == "sin_pdf":
        return "SIN_PDF", 0, ""

    if "sin_match" in estado_norm:
        return "SIN_MATCH", 0, ""

    if "error" in estado_norm:
        return "ERROR", 0, (error or "")[:120]

    return "OTRO", 0, ""


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
    error: str = "",
    tipo_resultado: str = "",
    filas_generadas: Optional[int] = None,
    motivo_no_registro: str = "",
):
    if not tipo_resultado:
        tipo_resultado, filas_generadas_auto, motivo_auto = _clasificar_auditoria_detalle(
            estado=estado,
            nuevos=int(nuevos or 0),
            pdf_name=pdf_name,
            subj=subj,
            error=error,
        )
        if filas_generadas is None:
            filas_generadas = filas_generadas_auto
        if not motivo_no_registro:
            motivo_no_registro = motivo_auto

    if filas_generadas is None:
        filas_generadas = int(nuevos or 0)

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
        "tipo_resultado": tipo_resultado or "",
        "filas_generadas": int(filas_generadas or 0),
        "motivo_no_registro": motivo_no_registro or "",
        "duracion_s": round(float(duracion_s or 0.0), 3),
        "nuevos": int(nuevos or 0),
        "enriquecidas": int(enriquecidas or 0),
        "fuente": fuente or "",
        "error": (error or "")[:500],
    })

def _resolver_cufe_numero_final(
    *,
    regs: Optional[List[Dict[str, object]]] = None,
    reg_pdf: Optional[Dict[str, object]] = None,
    cufe_pdf: str = "",
    numero_pdf: str = "",
) -> Tuple[str, str]:
    """
    Prioriza el CUFE/Número final real que terminó en Excel.
    Orden:
    1) regs del XML procesado
    2) reg_pdf del fallback PDF
    3) datos preliminares del PDF aprobado
    """
    cufe_final = ""
    numero_final = ""

    if regs:
        for d in regs:
            c = str(d.get("CUFE") or "").strip()
            n = str(d.get("Número de factura") or "").strip()
            if c and not cufe_final:
                cufe_final = c
            if n and not numero_final:
                numero_final = n
            if cufe_final and numero_final:
                break

    if reg_pdf:
        if not cufe_final:
            cufe_final = str(reg_pdf.get("CUFE") or "").strip()
        if not numero_final:
            numero_final = str(reg_pdf.get("Número de factura") or "").strip()

    if not cufe_final:
        cufe_final = str(cufe_pdf or "").strip()

    if not numero_final:
        numero_final = str(numero_pdf or "").strip()

    return cufe_final, numero_final


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

    # Evitar tokens demasiado genéricos como años puros: 2024, 2025, 2026...
    if re.fullmatch(r"\d{4}", t):
        try:
            year = int(t)
            if 1900 <= year <= 2100:
                return False
        except Exception:
            pass

    # Tokens puramente numéricos solo sirven si son suficientemente informativos.
    # Rechazamos años y números muy cortos / genéricos.
    if re.fullmatch(r"\d{4,20}", t):
        if len(t) < 5:
            return False
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

def _parse_datos_desde_subject_aprobado(subj: str) -> Dict[str, str]:
    """
    Extrae desde el asunto del correo aprobado:
    - numero_subject
    - radicado_subject
    - proyecto_subject
    - empresa_subject

    Ejemplo:
    Aprobado- Factura  -  FPP - 463496 - Radicado 191297 - PEOPLE PASS SAS - Joyco Consultores S.A.S. - NA
    """
    out = {
        "numero_subject": "",
        "radicado_subject": "",
        "proyecto_subject": "",
        "empresa_subject": "",
    }

    s = (subj or "").strip()
    if not s:
        return out

    s = s.replace("–", "-").replace("—", "-")
    s = re.sub(r"\s+", " ", s).strip()

    # -------- Radicado --------
    m_rad = re.search(r"Radicado\s+(\d{4,20})", s, flags=re.IGNORECASE)
    if m_rad:
        out["radicado_subject"] = m_rad.group(1).strip()

    # -------- Número de factura desde asunto --------
    # Captura lo que esté entre "Factura -" y "- Radicado"
    m_num = re.search(
        r"Factura(?:\s+de\s+servicio\s+p[uú]blico)?\s*-\s*(.*?)\s*-\s*Radicado\s+\d+",
        s,
        flags=re.IGNORECASE
    )
    if m_num:
        numero_raw = (m_num.group(1) or "").strip()
        numero_raw = numero_raw.replace("–", "-").replace("—", "-")
        numero_raw = re.sub(r"\s*-\s*", "-", numero_raw)
        numero_raw = re.sub(r"\s+", " ", numero_raw).strip()
        out["numero_subject"] = numero_raw

    # -------- Empresa / Proyecto --------
    # Todo lo que viene después del Radicado
    m_post = re.search(
        r"Radicado\s+\d+\s*-\s*(.*?)\s*$",
        s,
        flags=re.IGNORECASE
    )
    if m_post:
        cola = (m_post.group(1) or "").strip()
        partes = [p.strip() for p in cola.split(" - ") if p.strip()]

        # estructura típica:
        # EMPRESA - PROYECTO - NA
        if len(partes) >= 1:
            out["empresa_subject"] = partes[0]

        if len(partes) >= 2:
            if partes[-1].upper() == "NA":
                out["proyecto_subject"] = partes[-2].strip()
            else:
                out["proyecto_subject"] = partes[1].strip()

    return out


def _enriquecer_regs_desde_subject_si_falta(
    regs: List[Dict[str, object]],
    subj: str
) -> Tuple[int, str, str]:
    """
    En memoria, llena Radicado y ProyectoProceso en los regs
    usando el subject del correo aprobado, SOLO si vienen vacíos.

    Retorna:
        (filas_tocadas, radicado_subject, proyecto_subject)
    """
    if not regs:
        return 0, "", ""

    datos = _parse_datos_desde_subject_aprobado(subj)
    rad = (datos.get("radicado_subject") or "").strip()
    proy = (datos.get("proyecto_subject") or "").strip()

    if not rad and not proy:
        return 0, "", ""

    tocadas = 0

    for r in regs:
        cambio = False

        rad_actual = str(r.get("Radicado") or "").strip()
        proy_actual = str(r.get("ProyectoProceso") or "").strip()

        if rad and not rad_actual:
            r["Radicado"] = rad
            cambio = True

        if proy and not proy_actual:
            r["ProyectoProceso"] = proy
            cambio = True

        if cambio:
            tocadas += 1

    return tocadas, rad, proy


def _forzar_radicado_y_proyecto_en_filas(
    filas: List[Dict[str, object]],
    subj: str,
    estado: str,
) -> Tuple[List[Dict[str, object]], int, str, str]:
    """
    Regla obligatoria:
    si una factura ya quedó en estado OK y tiene filas para guardar,
    TODAS las filas deben salir con Radicado y ProyectoProceso.

    Prioridad:
    1) lo que ya venga en la fila
    2) subject del correo aprobado
    3) marcadores SIN_...
    """
    if not filas:
        return filas, 0, "", ""

    datos = _parse_datos_desde_subject_aprobado(subj)
    radicado_subject = str(datos.get("radicado_subject") or "").strip()
    proyecto_subject = str(datos.get("proyecto_subject") or "").strip()

    radicado_final = ""
    proyecto_final = ""

    # 1) Intentar tomar lo que ya venga en las filas
    for f in filas:
        if not radicado_final:
            radicado_final = str(f.get("Radicado") or "").strip()
        if not proyecto_final:
            proyecto_final = str(f.get("ProyectoProceso") or "").strip()
        if radicado_final and proyecto_final:
            break

    # 2) Completar desde subject
    if not radicado_final:
        radicado_final = radicado_subject
    if not proyecto_final:
        proyecto_final = proyecto_subject

    # 3) Si sigue vacío y ya hubo match/ok, obligarlo
    estado_norm = (estado or "").strip().lower()
    if estado_norm in {"ok", "ok_pdf_aprobadas_fallback", "ok_dian_pdf_only", "ok_dian_zip", "ok_registro_minimo"}:
        if not radicado_final:
            radicado_final = "SIN_RADICADO"
        if not proyecto_final:
            proyecto_final = "SIN_PROYECTO"

    enriquecidas = 0
    for f in filas:
        f["Radicado"] = radicado_final
        f["ProyectoProceso"] = proyecto_final

        if str(f.get("Radicado") or "").strip() and str(f.get("ProyectoProceso") or "").strip():
            enriquecidas += 1

    return filas, enriquecidas, radicado_final, proyecto_final

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
    archivos_ref: Optional[set[str]] = None,
    table_name: str = "TblFacturas",
    rows_dicts: Optional[List[Dict[str, object]]] = None,
) -> int:
    filas: List[Dict[str, object]] = []

    if rows_dicts:
        filas = [dict(x or {}) for x in rows_dicts if isinstance(x, dict)]
    elif archivos_ref:
        archivos_ref = _expand_archivos_ref(set(archivos_ref))
        filas = obtener_filas_por_archivos(archivos_ref)

    if not filas:
        return 0

    sp_facturas_path = f"{sp_excel_root}/facturas.xlsx".strip("/")
    xl = ExcelWorkbookGraph(sp_facturas_path)

    insertadas = xl.append_rows_dedup(
        table_name=table_name,
        rows_dicts=filas,
        key_cols=("Radicado", "Archivo", "Concepto"),
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
) -> Tuple[bool, int, int, int]:
    """
    Último recurso:
    si no hubo ZIP ni PDF externo y el PDF aprobado tiene CUFE válido,
    registrar directamente desde ese mismo PDF.

    Retorna:
        (aplico_fallback, total_nuevos, enriquecidas, insertadas_web)
    """
    if not cufe_pdf or not _cufe_is_valid(cufe_pdf):
        print(f"⛔ Fallback PDF_APROBADAS no aplica para {pdf_name}: PDF sin CUFE válido.")
        return False, 0, 0, 0

    print(f"✅ Fallback PDF_APROBADAS habilitado para {pdf_name} (CUFE válido).")

    try:
        reg = _asegurar_reg_7_conceptos(_generar_registro_pdf_only(pdf_tmp, pdf_name))

        numero_final = (
            ident_pdf.get("NUMERO_APROB")
            or ident_pdf.get("NUMERO")
            or reg.get("Número de factura")
            or ""
        )
        if numero_final and len(str(numero_final).strip()) >= 3:
            reg["Número de factura"] = str(numero_final).strip()

        total_nuevos = guardar_en_excel([reg])
        datos_subject_local = _parse_datos_desde_subject_aprobado(subj)
        radicado_local_force = str(reg.get("Radicado") or datos_subject_local.get("radicado_subject") or "").strip()
        proyecto_local_force = str(reg.get("ProyectoProceso") or datos_subject_local.get("proyecto_subject") or "").strip()
        archivos_force = {os.path.basename(pdf_name), str(reg.get("Archivo") or os.path.basename(pdf_name))}
        numeros_force = {str(reg.get("Número de factura") or numero_aprob or ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "").strip()}
        filas_match_force, filas_upd_force = _forzar_campos_obligatorios_en_excel_local(
            ARCHIVO_EXCEL,
            archivos_ref=archivos_force,
            numeros_ref=numeros_force,
            radicado=radicado_local_force,
            proyecto=proyecto_local_force,
        )
        if filas_upd_force > 0:
            print(
                f"[LOCAL FORCE] PDF/FALLBACK -> match={filas_match_force} | "
                f"actualizadas={filas_upd_force} | radicado={radicado_local_force} | proyecto={proyecto_local_force}"
            )

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

        enriquecidas = int(total_nuevos or 0)
        try:
            sincronizar_aprobaciones_en_facturas()
        except Exception as e:
            print(f"[APROB] Error al sincronizar aprobaciones: {e}")

        enriquecidas_subject = 0
        radicado_subject = ""
        proyecto_subject = ""
        try:
            regs_subject = [reg]
            enriquecidas_subject, radicado_subject, proyecto_subject = _enriquecer_regs_desde_subject_si_falta(regs_subject, subj)
            if enriquecidas_subject > 0:
                reg = regs_subject[0]
                print(
                    f"[SUBJECT] Fallback PDF_APROBADAS en memoria: "
                    f"filas={enriquecidas_subject} | radicado={radicado_subject} | proyecto={proyecto_subject}"
                )
        except Exception as e:
            print(f"[SUBJECT] Error enriqueciendo PDF_APROBADAS desde subject: {e}")

        try:
            filas_tmp, _, rad_tmp, proy_tmp = _forzar_radicado_y_proyecto_en_filas(
                filas=[reg],
                subj=subj,
                estado="ok_pdf_aprobadas_fallback",
            )
            reg = filas_tmp[0]
            print(f"[FORCE MEM] PDF/FALLBACK -> radicado={rad_tmp} | proyecto={proy_tmp}")
        except Exception as e:
            print(f"[FORCE MEM] Error reforzando obligatorios en fallback PDF: {e}")

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
                insertadas = _subir_factura_a_web_desde_local(
                    sp_excel_root=sp_excel,
                    archivos_ref=archivos_ref,
                    numeros_ref={str(reg.get("Número de factura") or ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "").strip()},
                    expected_rows=int(total_nuevos or 0),
                    table_name="TblFacturas",
                    rows_dicts=[reg],
                )
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

        cufe_final, numero_final = _resolver_cufe_numero_final(
            reg_pdf=reg,
            cufe_pdf=cufe_pdf,
            numero_pdf=reg.get("Número de factura") or ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "",
        )

        _push_detalle(
            detalle_rows, run_id, msg_id, subj,
            pdf_name=pdf_name,
            cufe=cufe_final,
            numero=numero_final,
            fecha_factura=ident_pdf.get("FECHA") or fecha_pdf,
            zip_match="(PDF_APROBADAS_FALLBACK)",
            estado="ok_pdf_aprobadas_fallback",
            duracion_s=(time.perf_counter() - t0),
            nuevos=int(total_nuevos or 0),
            enriquecidas=int(enriquecidas or 0),
            fuente=("PDF_APROBADAS|SUBJECT" if (radicado_subject or proyecto_subject) else "PDF_APROBADAS")
        )

        if total_nuevos > 0 and cufe_pdf:
            cufes_existentes.add(cufe_pdf)
            norm_cufes_existentes.add(cufe_pdf)

        return True, int(total_nuevos or 0), int(enriquecidas or 0), int(insertadas or 0)

    except Exception as e:
        print(f"⚠️ Falló fallback PDF_APROBADAS para {pdf_name}: {e}")
        return False, 0, 0, 0

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
        "insertadas": int,
        "estado": str
    }
    """

    if not cufe_pdf or not _cufe_is_valid(cufe_pdf):
        print(f"⚠️ Fallback PDF_APROBADAS sin CUFE válido para {pdf_name}. Se aplicará registro mínimo obligatorio.")
        return _registrar_minimo_obligatorio_desde_aprobadas(
            msg_id=msg_id,
            subj=subj,
            pdf_name=pdf_name,
            pdf_tmp=pdf_tmp,
            ident_pdf=ident_pdf,
            fecha_pdf=fecha_pdf,
            fecha_local=fecha_local,
            hora_local=hora_local,
            numero_aprob=numero_aprob,
            detalle_rows=detalle_rows,
            run_id=run_id,
            t0=t0,
            usar_processed_store=usar_processed_store,
            store=store,
            motivo="sin_cufe_o_pdf_no_legible",
        )

    print(f"✅ Fallback PDF_APROBADAS habilitado para {pdf_name} (CUFE válido).")

    reg = _asegurar_reg_7_conceptos(_generar_registro_pdf_only(pdf_tmp, pdf_name))

    numero_final_preferido = (
        ident_pdf.get("NUMERO_APROB")
        or ident_pdf.get("NUMERO")
        or reg.get("Número de factura")
        or ""
    )

    if numero_aprob and len(str(numero_aprob).strip()) >= 3:
        numero_final_preferido = str(numero_aprob).strip()

    if numero_final_preferido and len(str(numero_final_preferido).strip()) >= 3:
        reg["Número de factura"] = str(numero_final_preferido).strip()

    regs_tmp = [reg]
    regs_tmp, enriquecidas_forzadas, radicado_final, proyecto_final = _forzar_radicado_y_proyecto_en_filas(
        filas=regs_tmp,
        subj=subj,
        estado="ok_pdf_aprobadas_fallback",
    )
    reg = regs_tmp[0]

    cufe_final, numero_final = _resolver_cufe_numero_final(
        reg_pdf=reg,
        cufe_pdf=cufe_pdf,
        numero_pdf=reg.get("Número de factura") or ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "",
    )

    print(
        f"[FORZADO] PDF_APROBADAS -> radicado={radicado_final} | "
        f"proyecto={proyecto_final} | enriquecidas={enriquecidas_forzadas}"
    )

    total_nuevos = guardar_en_excel([reg])

    datos_subject_local = _parse_datos_desde_subject_aprobado(subj)
    radicado_local_force = str(reg.get("Radicado") or datos_subject_local.get("radicado_subject") or "").strip()
    proyecto_local_force = str(reg.get("ProyectoProceso") or datos_subject_local.get("proyecto_subject") or "").strip()
    archivos_force = {os.path.basename(pdf_name), str(reg.get("Archivo") or os.path.basename(pdf_name))}
    numeros_force = {str(reg.get("Número de factura") or numero_aprob or ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "").strip()}
    filas_match_force, filas_upd_force = _forzar_campos_obligatorios_en_excel_local(
        ARCHIVO_EXCEL,
        archivos_ref=archivos_force,
        numeros_ref=numeros_force,
        radicado=radicado_local_force,
        proyecto=proyecto_local_force,
    )
    if filas_upd_force > 0:
        print(
            f"[LOCAL FORCE] PDF/FALLBACK -> match={filas_match_force} | "
            f"actualizadas={filas_upd_force} | radicado={radicado_local_force} | proyecto={proyecto_local_force}"
        )

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

    enriquecidas = int(total_nuevos or 0)
    try:
        sincronizar_aprobaciones_en_facturas()
    except Exception as e:
        print(f"[APROB] Error al sincronizar aprobaciones: {e}")

    enriquecidas_subject = 0
    radicado_subject = ""
    proyecto_subject = ""
    try:
        regs_subject = [reg]
        enriquecidas_subject, radicado_subject, proyecto_subject = _enriquecer_regs_desde_subject_si_falta(regs_subject, subj)
        if enriquecidas_subject > 0:
            reg = regs_subject[0]
            print(
                f"[SUBJECT] Fallback último recurso PDF_APROBADAS en memoria: "
                f"filas={enriquecidas_subject} | radicado={radicado_subject} | proyecto={proyecto_subject}"
            )
    except Exception as e:
        print(f"[SUBJECT] Error enriqueciendo último recurso desde subject: {e}")

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
            insertadas = _subir_factura_a_web_desde_local(
                sp_excel_root=sp_excel,
                archivos_ref=archivos_ref,
                numeros_ref={str(reg.get("Número de factura") or numero_aprob or ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "").strip()},
                expected_rows=int(total_nuevos or 0),
                table_name="TblFacturas",
                rows_dicts=[reg],
            )
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
            "cufe": cufe_final,
        })

    try:
        marcar_mensaje_como_leido(msg_id)
    except Exception as e:
        print(f"⚠️ No se pudo marcar como leído: {e}")

    _push_detalle(
        detalle_rows, run_id, msg_id, subj,
        pdf_name=pdf_name,
        cufe=cufe_final,
        numero=numero_final,
        fecha_factura=ident_pdf.get("FECHA") or fecha_pdf,
        zip_match="(PDF_APROBADAS_FALLBACK)",
        estado="ok_pdf_aprobadas_fallback",
        duracion_s=(time.perf_counter() - t0),
        nuevos=int(total_nuevos or 0),
        enriquecidas=int(enriquecidas or 0),
        fuente=("PDF_APROBADAS|SUBJECT" if (radicado_subject or proyecto_subject) else "PDF_APROBADAS")
    )

    return {
        "handled": True,
        "ok": True,
        "nuevos": int(total_nuevos or 0),
        "enriquecidas": int(enriquecidas or 0),
        "insertadas": int(insertadas or 0),
        "estado": "ok_pdf_aprobadas_fallback",
    }



def _buscar_correo_origen_con_pdf_unico(
    *,
    msg_id_aprobado: str,
    subj_aprobado: str,
    ident_pdf: Dict[str, str],
    since_days: int,
    top_msgs: int = 80,
) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str]]:
    """
    Busca en Bandeja de entrada el correo origen correspondiente a la factura aprobada.
    Si encuentra un correo con EXACTAMENTE 1 PDF y 0 ZIPs, retorna ese PDF como fuente origen.
    """
    datos_subj = _parse_datos_desde_subject_aprobado(subj_aprobado)
    terminos = []
    for t in [
        ident_pdf.get("NUMERO_APROB") or "",
        ident_pdf.get("NUMERO") or "",
        datos_subj.get("radicado_subject") or "",
        _numero_from_subject(subj_aprobado) or "",
        Path((subj_aprobado or "")).stem,
    ]:
        t = str(t or "").strip()
        if t and t not in terminos:
            terminos.append(t)

    candidatos = []
    vistos = set()
    for term in terminos:
        try:
            lote = buscar_mensajes_inbox_por_asunto(
                asunto_contiene=term,
                top=top_msgs,
                since_days=since_days,
            ) or []
        except Exception as e:
            print(f"[ORIGEN PDF] Error buscando asunto={term!r}: {e}")
            lote = []

        for m in lote:
            mid = m.get("id")
            if not mid or mid == msg_id_aprobado or mid in vistos:
                continue
            vistos.add(mid)
            candidatos.append(m)

    print(f"[ORIGEN PDF] candidatos encontrados={len(candidatos)}")

    for m in candidatos:
        mid = m.get("id")
        subj = m.get("subject") or ""

        # evitar reusar el mismo correo de aprobadas
        if "aprobado" in normalize_text(subj):
            continue

        try:
            zips = listar_adjuntos_zip(mid) or []
        except Exception:
            zips = []

        try:
            pdfs = listar_adjuntos_pdf(mid) or []
        except Exception:
            pdfs = []

        if len(zips) == 0 and len(pdfs) == 1:
            att = pdfs[0]
            aid = att.get("id")
            if not aid:
                continue

            aname = att.get("name") or f"{aid}.pdf"
            safe_name = re.sub(r"[^A-Za-z0-9_. -]", "_", aname)
            local = os.path.join(TMP_DIR, f"origen_unico_{uuid.uuid4().hex}_{safe_name}")

            ok = descargar_adjunto_por_id(mid, aid, local)
            if ok and os.path.exists(local):
                print(f"[ORIGEN PDF] ✅ Correo origen con PDF único: asunto={subj} | pdf={aname}")
                return local, aname, mid, aid

            try:
                if os.path.exists(local):
                    os.remove(local)
            except Exception:
                pass

    print("[ORIGEN PDF] No se encontró correo origen con PDF único utilizable.")
    return None, None, None, None


def _registrar_desde_pdf_origen_unico(
    *,
    msg_id: str,
    subj: str,
    pdf_name_aprobado: str,
    pdf_origen_path: str,
    pdf_origen_name: str,
    ident_pdf_aprobado: Dict[str, str],
    fecha_local: str,
    hora_local: str,
    run_id: str,
    detalle_rows: List[Dict[str, object]],
    resumen: List[Tuple[str, float, str, int]],
    t0: float,
    usar_processed_store: bool,
    store: ProcessedStore,
    source_msg_id: str = "",
    source_att_id: str = "",
) -> Tuple[bool, int, int, int]:
    """
    Registra usando el PDF único encontrado en el correo origen.
    """
    try:
        ident_origen = parse_identificadores_pdf(extraer_texto_pdf(pdf_origen_path)) or {}

        reg = _asegurar_reg_7_conceptos(_generar_registro_pdf_only(pdf_origen_path, pdf_origen_name))

        numero_final = (
            ident_pdf_aprobado.get("NUMERO_APROB")
            or ident_pdf_aprobado.get("NUMERO")
            or ident_origen.get("NUMERO_APROB")
            or ident_origen.get("NUMERO")
            or reg.get("Número de factura")
            or ""
        )
        if numero_final and len(str(numero_final).strip()) >= 3:
            reg["Número de factura"] = str(numero_final).strip()

        regs_tmp = [reg]
        regs_tmp, enriquecidas_forzadas, radicado_final, proyecto_final = _forzar_radicado_y_proyecto_en_filas(
            filas=regs_tmp,
            subj=subj,
            estado="ok_pdf_origen_unico",
        )
        reg = regs_tmp[0]

        total_nuevos = guardar_en_excel([reg])

        datos_subject_local = _parse_datos_desde_subject_aprobado(subj)
        radicado_local_force = str(reg.get("Radicado") or datos_subject_local.get("radicado_subject") or "").strip()
        proyecto_local_force = str(reg.get("ProyectoProceso") or datos_subject_local.get("proyecto_subject") or "").strip()
        archivos_force = {os.path.basename(pdf_origen_name), str(reg.get("Archivo") or os.path.basename(pdf_origen_name)), os.path.basename(pdf_name_aprobado)}
        numeros_force = {str(reg.get("Número de factura") or ident_pdf_aprobado.get("NUMERO_APROB") or ident_pdf_aprobado.get("NUMERO") or "").strip()}
        _forzar_campos_obligatorios_en_excel_local(
            ARCHIVO_EXCEL,
            archivos_ref=archivos_force,
            numeros_ref=numeros_force,
            radicado=radicado_local_force,
            proyecto=proyecto_local_force,
        )

        historial_actualizado = False
        if total_nuevos > 0:
            registrar_historial_por_zip([{
                "Fecha": fecha_local,
                "Hora": hora_local,
                "Archivo ZIP": "(PDF_ORIGEN_UNICO)",
                "Nuevos XML guardados": total_nuevos,
                "Errores encontrados": 0
            }])
            historial_actualizado = True

        enriquecidas = int(total_nuevos or 0)
        try:
            sincronizar_aprobaciones_en_facturas()
        except Exception as e:
            print(f"[APROB] Error al sincronizar aprobaciones: {e}")

        sp_ext_root = f"{BASE_SP}/extraidos/pdf_origen_unico"
        sp_excel = f"{BASE_SP}/excel"

        sp_disponible = True
        try:
            ensure_folder(sp_ext_root)
            ensure_folder(sp_excel)
        except Exception as e:
            sp_disponible = False
            print(f"⚠️ SharePoint no disponible en PDF origen único: {e}")

        if sp_disponible:
            try:
                upload_small_file(pdf_origen_path, f"{sp_ext_root}/{os.path.basename(pdf_origen_name)}", mode="skip")
            except Exception as e:
                print(f"⚠️ No pude subir PDF origen único a SharePoint: {e}")

        hubo_cambios_excel = (total_nuevos > 0) or (enriquecidas > 0)
        insertadas = 0
        if sp_disponible and hubo_cambios_excel:
            try:
                archivos_ref = {
                    os.path.basename(pdf_origen_name),
                    str(reg.get("Archivo") or os.path.basename(pdf_origen_name)),
                    os.path.basename(pdf_name_aprobado),
                }
                insertadas = _try_workbook_append(sp_excel, archivos_ref, table_name="TblFacturas")
                print(f"✅ Workbook API (PDF_ORIGEN_UNICO): +{insertadas} fila(s) nuevas en TblFacturas.")
            except Exception as e:
                print(f"⚠️ Workbook API falló (PDF_ORIGEN_UNICO): {e}")

        if sp_disponible:
            _subir_excels_a_sharepoint(sp_excel, hubo_cambios_excel, historial_actualizado)

        if usar_processed_store:
            store.mark_processed(msg_id, {
                "status": "ok_pdf_origen_unico",
                "pdf": pdf_name_aprobado,
                "origen_pdf": pdf_origen_name,
                "nuevos": int(total_nuevos),
                "enriquecidas": int(enriquecidas),
                "src_msg": source_msg_id,
                "src_att": source_att_id,
            })

        try:
            marcar_mensaje_como_leido(msg_id)
        except Exception as e:
            print(f"⚠️ No se pudo marcar como leído: {e}")

        secs = time.perf_counter() - t0
        resumen.append((pdf_name_aprobado, secs, "pdf origen unico", total_nuevos))

        cufe_final, numero_final = _resolver_cufe_numero_final(
            reg_pdf=reg,
            cufe_pdf=_norm_cufe(ident_origen.get("CUFE") or ident_pdf_aprobado.get("CUFE") or ""),
            numero_pdf=reg.get("Número de factura") or ident_origen.get("NUMERO") or ident_pdf_aprobado.get("NUMERO_APROB") or ident_pdf_aprobado.get("NUMERO") or "",
        )

        _push_detalle(
            detalle_rows, run_id, msg_id, subj,
            pdf_name=pdf_name_aprobado,
            cufe=cufe_final,
            numero=numero_final,
            fecha_factura=ident_origen.get("FECHA") or ident_pdf_aprobado.get("FECHA") or "",
            zip_match="(PDF_ORIGEN_UNICO)",
            estado="ok_pdf_origen_unico",
            duracion_s=secs,
            nuevos=int(total_nuevos or 0),
            enriquecidas=int(enriquecidas or 0),
            fuente="PDF_ORIGEN_UNICO"
        )

        return True, int(total_nuevos or 0), int(enriquecidas or 0), int(insertadas or 0)

    except Exception as e:
        print(f"⚠️ Falló PDF_ORIGEN_UNICO para {pdf_name_aprobado}: {e}")
        return False, 0, 0, 0



def _construir_registro_minimo_obligatorio(
    *,
    subj: str,
    pdf_name: str,
    numero_aprob: str = "",
    fecha_pdf: str = "",
    ident_pdf: Optional[Dict[str, str]] = None,
) -> Dict[str, object]:
    """
    Construye un registro mínimo obligatorio para facturas de solo aprobadas
    cuando no fue posible leer CUFE/XML/PDF, pero la regla de negocio exige
    registrar sí o sí la factura con 7 conceptos.
    """
    ident_pdf = ident_pdf or {}
    datos_subj = _parse_datos_desde_subject_aprobado(subj)

    numero_final = (
        str(numero_aprob or "").strip()
        or str(ident_pdf.get("NUMERO_APROB") or "").strip()
        or str(ident_pdf.get("NUMERO") or "").strip()
        or str(datos_subj.get("numero_subject") or "").strip()
        or Path(pdf_name or "").stem
    ).strip()

    fecha_final = str(fecha_pdf or ident_pdf.get("FECHA") or "").strip()
    fecha_final = normalizar_fecha(fecha_final) or fecha_final if fecha_final else ""

    y = fecha_final[:4] if len(fecha_final) >= 4 else ""
    mo = fecha_final[5:7] if len(fecha_final) >= 7 else ""
    d = fecha_final[8:10] if len(fecha_final) >= 10 else ""

    empresa = str(datos_subj.get("empresa_subject") or "").strip()

    reg = {
        "Archivo": os.path.basename(pdf_name or ""),
        "Empresa emisora": empresa,
        "CUFE": "",
        "Ciudad emisora": "",
        "Código ciudad": "",
        "NIT": "",
        "Cliente": "",
        "Número de factura": numero_final,
        "Año": y,
        "Mes": mo,
        "Día": d,
        "Tipo de contribuyente": "",
        "Actividad económica": "",
        "DescripcionLineas": "REGISTRO MÍNIMO OBLIGATORIO",
        "Radicado": str(datos_subj.get("radicado_subject") or "").strip(),
        "ProyectoProceso": str(datos_subj.get("proyecto_subject") or "").strip(),
        "Subtotal": 0.0,
        "IVA 5%": 0.0,
        "IVA 19%": 0.0,
        "Retención de IVA": 0.0,
        "Retención de ICA": 0.0,
        "Retención en la fuente": 0.0,
        "Total": 0.0,
    }
    return _asegurar_reg_7_conceptos(reg)


def _registrar_minimo_obligatorio_desde_aprobadas(
    *,
    msg_id: str,
    subj: str,
    pdf_name: str,
    pdf_tmp: str,
    ident_pdf: Optional[Dict[str, str]],
    fecha_pdf: str,
    fecha_local: str,
    hora_local: str,
    numero_aprob: str,
    detalle_rows: List[Dict[str, object]],
    run_id: str,
    t0: float,
    usar_processed_store: bool,
    store,
    motivo: str = "ok_registro_minimo",
) -> Dict[str, object]:
    """
    Registro mínimo obligatorio:
    si una factura de solo aprobadas no pudo procesarse por CUFE/ZIP/PDF,
    se registra de todos modos con Radicado, ProyectoProceso y los 7 conceptos.
    """
    try:
        reg = _construir_registro_minimo_obligatorio(
            subj=subj,
            pdf_name=pdf_name,
            numero_aprob=numero_aprob,
            fecha_pdf=fecha_pdf,
            ident_pdf=ident_pdf or {},
        )

        regs_tmp = [reg]
        regs_tmp, _, radicado_final, proyecto_final = _forzar_radicado_y_proyecto_en_filas(
            filas=regs_tmp,
            subj=subj,
            estado="ok_registro_minimo",
        )
        reg = regs_tmp[0]

        total_nuevos = guardar_en_excel([reg])

        datos_subject_local = _parse_datos_desde_subject_aprobado(subj)
        radicado_local_force = str(reg.get("Radicado") or datos_subject_local.get("radicado_subject") or "").strip()
        proyecto_local_force = str(reg.get("ProyectoProceso") or datos_subject_local.get("proyecto_subject") or "").strip()
        archivos_force = {os.path.basename(pdf_name), str(reg.get("Archivo") or os.path.basename(pdf_name))}
        numeros_force = {str(reg.get("Número de factura") or numero_aprob or "").strip()}

        filas_match_force, filas_upd_force = _forzar_campos_obligatorios_en_excel_local(
            ARCHIVO_EXCEL,
            archivos_ref=archivos_force,
            numeros_ref=numeros_force,
            radicado=radicado_local_force,
            proyecto=proyecto_local_force,
        )
        if filas_upd_force <= 0 and int(total_nuevos or 0) > 0:
            filas_match_force, filas_upd_force = _forzar_campos_obligatorios_en_ultimas_filas(
                ARCHIVO_EXCEL,
                expected_rows=int(total_nuevos or 0),
                radicado=radicado_local_force,
                proyecto=proyecto_local_force,
            )

        historial_actualizado = False
        if total_nuevos > 0:
            registrar_historial_por_zip([{
                "Fecha": fecha_local,
                "Hora": hora_local,
                "Archivo ZIP": f"(REGISTRO_MINIMO:{os.path.basename(pdf_name)})",
                "Nuevos XML guardados": total_nuevos,
                "Errores encontrados": 0,
            }])
            historial_actualizado = True

        enriquecidas = int(total_nuevos or 0)
        try:
            sincronizar_aprobaciones_en_facturas()
        except Exception as e:
            print(f"[APROB] Error al sincronizar aprobaciones (registro mínimo): {e}")

        sp_ext_root = f"{BASE_SP}/extraidos/pdf_aprobadas_fallback"
        sp_excel = f"{BASE_SP}/excel"

        sp_disponible = True
        try:
            ensure_folder(sp_ext_root)
            ensure_folder(sp_excel)
        except Exception as e:
            sp_disponible = False
            print(f"⚠️ SharePoint no disponible en registro mínimo: {e}")

        if sp_disponible and pdf_tmp and os.path.exists(pdf_tmp):
            try:
                upload_small_file(pdf_tmp, f"{sp_ext_root}/{os.path.basename(pdf_name)}", mode="skip")
            except Exception as e:
                print(f"⚠️ No pude subir PDF de registro mínimo a SharePoint: {e}")

        hubo_cambios_excel = (total_nuevos > 0) or (enriquecidas > 0)
        insertadas = 0
        if sp_disponible and hubo_cambios_excel:
            try:
                archivos_ref = {
                    os.path.basename(pdf_name),
                    str(reg.get("Archivo") or os.path.basename(pdf_name)),
                }
                insertadas = _subir_factura_a_web_desde_local(
                    sp_excel_root=sp_excel,
                    archivos_ref=archivos_ref,
                    numeros_ref={str(reg.get("Número de factura") or numero_aprob or "").strip()},
                    expected_rows=int(total_nuevos or 0),
                    table_name="TblFacturas",
                    rows_dicts=[reg],
                )
            except Exception as e:
                print(f"⚠️ Workbook API falló (registro mínimo): {e}")

        if sp_disponible:
            _subir_excels_a_sharepoint(sp_excel, hubo_cambios_excel, historial_actualizado)

        if usar_processed_store:
            store.mark_processed(msg_id, {
                "status": "ok_registro_minimo",
                "pdf": pdf_name,
                "nuevos": int(total_nuevos or 0),
                "enriquecidas": int(enriquecidas or 0),
                "motivo": motivo,
            })

        try:
            marcar_mensaje_como_leido(msg_id)
        except Exception as e:
            print(f"⚠️ No se pudo marcar como leído (registro mínimo): {e}")

        secs = time.perf_counter() - t0
        _push_detalle(
            detalle_rows, run_id, msg_id, subj,
            pdf_name=pdf_name,
            cufe="",
            numero=str(reg.get("Número de factura") or numero_aprob or ""),
            fecha_factura=fecha_pdf or "",
            zip_match=f"(REGISTRO_MINIMO:{motivo})",
            estado="ok_registro_minimo",
            duracion_s=secs,
            nuevos=int(total_nuevos or 0),
            enriquecidas=int(enriquecidas or 0),
            fuente="REGISTRO_MINIMO",
            error=motivo,
        )

        return {
            "handled": True,
            "ok": True,
            "nuevos": int(total_nuevos or 0),
            "enriquecidas": int(enriquecidas or 0),
            "insertadas": int(insertadas or 0),
            "estado": "ok_registro_minimo",
        }
    except Exception as e:
        print(f"⚠️ Falló registro mínimo obligatorio para {pdf_name}: {e}")
        return {
            "handled": True,
            "ok": False,
            "nuevos": 0,
            "enriquecidas": 0,
            "insertadas": 0,
            "estado": "error_registro_minimo",
        }



def _agregar_filas_al_buffer_web_run(
    buffer: List[Dict[str, object]],
    filas: Optional[List[Dict[str, object]]],
    *,
    origen: str = "",
) -> int:
    """
    Buffer incremental del run para reconciliar LOCAL -> WEB al final.

    Importante:
    - No sube todo el histórico.
    - Solo guarda en memoria las filas generadas durante este run.
    - Filtra filas sin Concepto para evitar la fila fantasma del web.
    - No modifica la lógica local que ya funciona.
    """
    if buffer is None or not filas:
        return 0

    agregadas = 0
    for row in filas:
        if not isinstance(row, dict):
            continue

        d = dict(row)

        concepto = str(d.get("Concepto") or "").strip()
        if not concepto:
            continue

        if "Número de factura" in d:
            d["Número de factura"] = _forzar_texto_numero_factura(d.get("Número de factura", ""))

        # La llave web actual es Radicado + Archivo + Concepto.
        # Si falta Radicado/Archivo, igual intentamos conservar la fila si tiene Concepto,
        # porque el Workbook Graph filtrará llaves incompletas si corresponde.
        # En la práctica el controller ya fuerza Radicado y Archivo antes de guardar.
        buffer.append(d)
        agregadas += 1

    if agregadas:
        print(f"[WEB BUFFER] +{agregadas} fila(s) agregadas al buffer del run. origen={origen}")

    return agregadas


def _reconciliar_web_desde_buffer_run(
    *,
    sp_excel_root: str,
    rows_web_run_buffer: List[Dict[str, object]],
    table_name: str = "TblFacturas",
) -> int:
    """
    Reconciliación final incremental del run.

    En vez de leer las últimas N filas del Excel local, usa el buffer real
    de filas generadas durante esta ejecución. Esto evita perder filas cuando
    el Excel local se reordena, deduplica o cambia físicamente de orden.

    El Workbook API deduplica por Radicado + Archivo + Concepto,
    así que esta llamada no duplica lo ya insertado; solo completa faltantes.
    """
    if not rows_web_run_buffer:
        print("[WEB BUFFER] No hay filas en buffer para reconciliar.")
        return 0

    filas_validas: List[Dict[str, object]] = []
    vistos = set()

    for row in rows_web_run_buffer:
        if not isinstance(row, dict):
            continue

        d = dict(row)
        concepto = str(d.get("Concepto") or "").strip()
        if not concepto:
            continue

        if "Número de factura" in d:
            d["Número de factura"] = _forzar_texto_numero_factura(d.get("Número de factura", ""))

        # Deduplicación interna del buffer para no mandar la misma fila varias veces.
        k = (
            str(d.get("Radicado") or "").strip(),
            str(d.get("Archivo") or "").strip(),
            concepto,
        )
        if k in vistos:
            continue
        vistos.add(k)
        filas_validas.append(d)

    if not filas_validas:
        print("[WEB BUFFER] Buffer sin filas válidas con Concepto.")
        return 0

    print(
        f"[WEB BUFFER] Reconciliación final: buffer={len(rows_web_run_buffer)} | "
        f"validas={len(filas_validas)} | tabla={table_name}"
    )

    try:
        ensure_folder(sp_excel_root)
    except Exception as e:
        print(f"[WEB BUFFER] No se pudo asegurar carpeta Excel en SharePoint: {e}")
        return 0

    try:
        insertadas = _try_workbook_append_rows(
            sp_excel_root,
            filas_validas,
            table_name=table_name,
        )
        print(f"[WEB BUFFER] Reconciliación final insertó/completó: {insertadas} fila(s)")
        return int(insertadas or 0)
    except Exception as e:
        print(f"[WEB BUFFER] Falló reconciliación final: {e}")
        return 0


def _try_workbook_append_rows(
    sp_excel_root: str,
    rows_dicts: List[Dict[str, object]],
    table_name: str = "TblFacturas"
) -> int:
    if not rows_dicts:
        return 0

    filas = []
    for row in rows_dicts:
        if not isinstance(row, dict):
            continue

        fila = dict(row)

        if "Número de factura" in fila:
            fila["Número de factura"] = _forzar_texto_numero_factura(
                fila.get("Número de factura", "")
            )

        filas.append(fila)

    if not filas:
        return 0

    sp_facturas_path = f"{sp_excel_root}/facturas.xlsx".strip("/")
    xl = ExcelWorkbookGraph(sp_facturas_path)

    insertadas = xl.append_rows_dedup(
        table_name=table_name,
        rows_dicts=filas,
        key_cols=("Radicado", "Archivo", "Concepto"),
        require_table=True,
    )
    return int(insertadas or 0)


def _obtener_filas_locales_por_archivos_y_numeros(
    excel_path: str,
    *,
    archivos_ref: Optional[set[str]] = None,
    numeros_ref: Optional[set[str]] = None,
    last_n: int = 0,
) -> List[Dict[str, object]]:
    if not excel_path or not os.path.exists(excel_path):
        return []

    try:
        from openpyxl import load_workbook

        wb = load_workbook(excel_path, data_only=True)
        ws = wb["Facturas"] if "Facturas" in wb.sheetnames else wb.active

        headers = [str(ws.cell(row=1, column=c).value or "").strip() for c in range(1, ws.max_column + 1)]

        archivos_ref = _expand_archivos_ref(set(archivos_ref or set()))
        nums_norm = {_normalizar_numero_para_match_local(str(n or "").strip()) for n in (numeros_ref or set()) if str(n or "").strip()}
        nums_norm.discard("")

        rows = []
        for r in range(2, ws.max_row + 1):
            values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
            row = {headers[i]: values[i] for i in range(len(headers)) if headers[i]}

            arc_val = str(row.get("Archivo") or "").strip()
            num_val = str(row.get("Número de factura") or "").strip()

            arc_hit = False
            num_hit = False

            if archivos_ref and arc_val:
                arc_hit = (arc_val in archivos_ref) or (os.path.basename(arc_val) in archivos_ref)

            if nums_norm and num_val:
                num_hit = _normalizar_numero_para_match_local(num_val) in nums_norm

            if archivos_ref or nums_norm:
                if not arc_hit and not num_hit:
                    continue

            rows.append(row)

        if rows:
            return rows

        if last_n > 0 and ws.max_row > 1:
            start = max(2, ws.max_row - last_n + 1)
            out = []
            for r in range(start, ws.max_row + 1):
                values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
                out.append({headers[i]: values[i] for i in range(len(headers)) if headers[i]})
            return out

        return []
    except Exception as e:
        print(f"[LOCAL->WEB] Error leyendo filas locales: {e}")
        return []


def _forzar_campos_obligatorios_en_ultimas_filas(
    excel_path: str,
    *,
    expected_rows: int,
    radicado: str = "",
    proyecto: str = "",
) -> Tuple[int, int]:
    if not excel_path or not os.path.exists(excel_path) or expected_rows <= 0:
        return 0, 0

    radicado = str(radicado or "").strip() or "SIN_RADICADO"
    proyecto = str(proyecto or "").strip() or "SIN_PROYECTO"

    try:
        from openpyxl import load_workbook

        wb = load_workbook(excel_path)
        ws = wb["Facturas"] if "Facturas" in wb.sheetnames else wb.active

        headers = {}
        for c in range(1, ws.max_column + 1):
            key = str(ws.cell(row=1, column=c).value or "").strip()
            if key:
                headers[key] = c

        col_rad = headers.get("Radicado")
        col_pro = headers.get("ProyectoProceso")
        if not col_rad or not col_pro:
            return 0, 0

        start = max(2, ws.max_row - expected_rows + 1)
        filas_match = 0
        filas_upd = 0
        for r in range(start, ws.max_row + 1):
            filas_match += 1
            cell_rad = ws.cell(r, col_rad)
            cell_pro = ws.cell(r, col_pro)
            cambio = False
            if not str(cell_rad.value or "").strip():
                cell_rad.value = radicado
                cambio = True
            if not str(cell_pro.value or "").strip():
                cell_pro.value = proyecto
                cambio = True
            if cambio:
                filas_upd += 1

        if filas_upd > 0:
            wb.save(excel_path)
        return filas_match, filas_upd
    except Exception as e:
        print(f"[LOCAL FORCE LAST] Error forzando últimas filas: {e}")
        return 0, 0


def _subir_factura_a_web_desde_local(
    *,
    sp_excel_root: str,
    archivos_ref: Optional[set[str]] = None,
    numeros_ref: Optional[set[str]] = None,
    expected_rows: int = 0,
    table_name: str = "TblFacturas",
    rows_dicts: Optional[List[Dict[str, object]]] = None,
) -> int:
    insertadas = 0

    rows_dicts = [dict(x or {}) for x in (rows_dicts or []) if isinstance(x, dict)]
    if rows_dicts:
        try:
            insertadas = _try_workbook_append_rows(
                sp_excel_root,
                rows_dicts,
                table_name=table_name,
            )
        except Exception as e:
            print(f"[LOCAL->WEB] append por filas en memoria falló: {e}")

        if expected_rows > 0 and insertadas >= expected_rows:
            return int(insertadas or 0)

        try:
            if archivos_ref:
                extra_arch = _try_workbook_append(
                    sp_excel_root,
                    set(archivos_ref),
                    table_name=table_name,
                )
                insertadas = max(int(insertadas or 0), int(extra_arch or 0))
        except Exception as e:
            print(f"[LOCAL->WEB] append por archivos falló: {e}")

    if expected_rows > 0 and insertadas >= expected_rows:
        return int(insertadas or 0)

    rows_local = _obtener_filas_locales_por_archivos_y_numeros(
        ARCHIVO_EXCEL,
        archivos_ref=set(archivos_ref or set()),
        numeros_ref=set(numeros_ref or set()),
        last_n=expected_rows,
    )
    if not rows_local:
        return int(insertadas or 0)

    try:
        extra = _try_workbook_append_rows(sp_excel_root, rows_local, table_name=table_name)
        insertadas = max(int(insertadas or 0), int(extra or 0))
    except Exception as e:
        print(f"[LOCAL->WEB] append por filas locales falló: {e}")

    return int(insertadas or 0)


def _obtener_ultimas_filas_locales_para_reconciliacion_web(
    excel_path: str,
    *,
    total_filas_run: int,
) -> List[Dict[str, object]]:
    """
    Respaldo de reconciliación:
    lee las últimas N filas del Excel local del run actual.
    Se usa además del buffer para cubrir ramas auxiliares que generan filas
    dentro de helpers y no tienen acceso directo al buffer.
    """
    if not excel_path or not os.path.exists(excel_path):
        return []

    total_filas_run = int(total_filas_run or 0)
    if total_filas_run <= 0:
        return []

    try:
        from openpyxl import load_workbook

        wb = load_workbook(excel_path, data_only=True)
        ws = wb["Facturas"] if "Facturas" in wb.sheetnames else wb.active

        headers = [
            str(ws.cell(row=1, column=c).value or "").strip()
            for c in range(1, ws.max_column + 1)
        ]

        if not headers:
            return []

        start_row = max(2, ws.max_row - total_filas_run + 1)
        filas: List[Dict[str, object]] = []

        for r in range(start_row, ws.max_row + 1):
            row = {}
            for idx, header in enumerate(headers, start=1):
                if not header:
                    continue
                row[header] = ws.cell(row=r, column=idx).value

            if not str(row.get("Concepto") or "").strip():
                continue

            if "Número de factura" in row:
                row["Número de factura"] = _forzar_texto_numero_factura(row.get("Número de factura", ""))

            filas.append(row)

        print(
            f"[WEB LOCAL BACKUP] Últimas filas locales leídas: "
            f"esperadas={total_filas_run} | encontradas={len(filas)}"
        )
        return filas

    except Exception as e:
        print(f"[WEB LOCAL BACKUP] Error leyendo últimas filas locales: {e}")
        return []


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
    filas_local_total = 0
    filas_web_total = 0

    # Resumen inteligente por factura (PASO 3.2)
    ok_total = 0
    ok_registradas = 0
    ok_no_registrables = 0

    dian_total = 0
    dian_registradas = 0
    dian_no_registrables = 0

    facturas_con_filas = 0
    facturas_sin_registro = 0

    detalle_rows: List[Dict[str, object]] = []

    # Buffer real de filas generadas en este run para reconciliar al web al final.
    # No contiene histórico, solo lo producido en esta ejecución.
    rows_web_run_buffer: List[Dict[str, object]] = []

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
            print("[APROB] Correo aprobado sin PDF detectado. Se fuerza REGISTRO MÍNIMO obligatorio.")
            resultado_min = _registrar_minimo_obligatorio_desde_aprobadas(
                msg_id=msg_id,
                subj=subj,
                pdf_name="SIN_ADJUNTO_APROBADO",
                pdf_tmp="",
                ident_pdf={},
                fecha_pdf="",
                fecha_local=fecha_local,
                hora_local=hora_local,
                numero_aprob=(_numero_from_subject(subj) or ""),
                detalle_rows=detalle_rows,
                run_id=run_id,
                t0=t0,
                usar_processed_store=usar_processed_store,
                store=store,
                motivo="sin_pdf_o_imagen_aprobada",
            )

            total_nuevos_min = int(resultado_min.get("nuevos", 0) or 0)
            insertadas_min = int(resultado_min.get("insertadas", 0) or 0)
            enriquecidas_min = int(resultado_min.get("enriquecidas", 0) or 0)

            cnt_ok += 1
            ok_total += 1
            if total_nuevos_min > 0:
                ok_registradas += 1
                facturas_con_filas += 1
            else:
                ok_no_registrables += 1
                facturas_sin_registro += 1

            msgs_procesados += 1
            nuevos_total += total_nuevos_min
            enriq_total += enriquecidas_min
            filas_local_total += total_nuevos_min
            filas_web_total += insertadas_min

            secs = time.perf_counter() - t0
            resumen.append(("SIN_ADJUNTO_APROBADO", secs, "registro minimo sin pdf/imagen", total_nuevos_min))

            procesados += 1
            sin_match_consec = 0
            sin_nuevos_consec = 0 if total_nuevos_min > 0 else (sin_nuevos_consec + 1)
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
            print(f"⚠️ PDF no probable factura electrónica: {pdf_name}. Se aplicará registro mínimo obligatorio.")
            resultado_min = _registrar_minimo_obligatorio_desde_aprobadas(
                msg_id=msg_id,
                subj=subj,
                pdf_name=pdf_name,
                pdf_tmp=pdf_tmp,
                ident_pdf=ident_pdf,
                fecha_pdf=fecha_pdf,
                fecha_local=fecha_local,
                hora_local=hora_local,
                numero_aprob=numero_aprob,
                detalle_rows=detalle_rows,
                run_id=run_id,
                t0=t0,
                usar_processed_store=usar_processed_store,
                store=store,
                motivo="filtro_no_factura",
            )

            secs = time.perf_counter() - t0
            resumen.append((pdf_name, secs, "registro minimo obligatorio", int(resultado_min.get("nuevos", 0) or 0)))

            cnt_ok += 1
            ok_total += 1
            if int(resultado_min.get("nuevos", 0) or 0) > 0:
                ok_registradas += 1
                facturas_con_filas += 1
            else:
                ok_no_registrables += 1
                facturas_sin_registro += 1

            msgs_procesados += 1
            nuevos_total += int(resultado_min.get("nuevos", 0) or 0)
            enriq_total += int(resultado_min.get("enriquecidas", 0) or 0)
            filas_local_total += int(resultado_min.get("nuevos", 0) or 0)
            filas_web_total += int(resultado_min.get("insertadas", 0) or 0)

            sin_match_consec = 0
            if int(resultado_min.get("nuevos", 0) or 0) == 0:
                sin_nuevos_consec += 1
            else:
                sin_nuevos_consec = 0

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

            if pdf_real_path:
                pdf_real_name = os.path.basename(pdf_real_path)
                reg = _asegurar_reg_7_conceptos(_generar_registro_pdf_only(pdf_real_path, pdf_real_name))

                if numero_aprob and len(numero_aprob) >= 3:
                    reg["Número de factura"] = numero_aprob

                regs_tmp = [reg]
                regs_tmp, enriquecidas_forzadas, radicado_final, proyecto_final = _forzar_radicado_y_proyecto_en_filas(
                    filas=regs_tmp,
                    subj=subj,
                    estado="ok_dian_pdf_only",
                )
                reg = regs_tmp[0]

                print(
                    f"[FORZADO] DIAN_PDF -> radicado={radicado_final} | "
                    f"proyecto={proyecto_final} | enriquecidas={enriquecidas_forzadas}"
                )

                total_nuevos = guardar_en_excel([reg])
                _agregar_filas_al_buffer_web_run(rows_web_run_buffer, [reg], origen="dian_pdf")

                datos_subject_local = _parse_datos_desde_subject_aprobado(subj)
                radicado_local_force = str(reg.get("Radicado") or datos_subject_local.get("radicado_subject") or "").strip()
                proyecto_local_force = str(reg.get("ProyectoProceso") or datos_subject_local.get("proyecto_subject") or "").strip()
                archivos_force = {os.path.basename(pdf_name), str(reg.get("Archivo") or os.path.basename(pdf_name))}
                numeros_force = {str(reg.get("Número de factura") or numero_aprob or ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "").strip()}
                filas_match_force, filas_upd_force = _forzar_campos_obligatorios_en_excel_local(
                    ARCHIVO_EXCEL,
                    archivos_ref=archivos_force,
                    numeros_ref=numeros_force,
                    radicado=radicado_local_force,
                    proyecto=proyecto_local_force,
                )
                if filas_upd_force <= 0 and int(total_nuevos or 0) > 0:
                    filas_match_force, filas_upd_force = _forzar_campos_obligatorios_en_ultimas_filas(
                        ARCHIVO_EXCEL,
                        expected_rows=int(total_nuevos or 0),
                        radicado=radicado_local_force,
                        proyecto=proyecto_local_force,
                    )
                if filas_upd_force > 0:
                    print(
                        f"[LOCAL FORCE] PDF/FALLBACK -> match={filas_match_force} | "
                        f"actualizadas={filas_upd_force} | radicado={radicado_local_force} | proyecto={proyecto_local_force}"
                    )

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

                enriquecidas = int(total_nuevos or 0)
                try:
                    sincronizar_aprobaciones_en_facturas()
                except Exception as e:
                    print(f"[APROB] Error al sincronizar aprobaciones: {e}")

                print("☁️  Subiendo a SharePoint (DIAN / PDF)...")
                sp_ext_root = f"{BASE_SP}/extraidos/pdf_dian_fallback"
                sp_excel = f"{BASE_SP}/excel"

                sp_disponible = True
                try:
                    ensure_folder(sp_ext_root)
                    ensure_folder(sp_excel)
                except Exception as e:
                    sp_disponible = False
                    print(f"⚠️ SharePoint no disponible en rama DIAN PDF: {e}")

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
                        archivos_ref = {
                            os.path.basename(pdf_real_name),
                            str(reg.get("Archivo") or pdf_real_name),
                            os.path.basename(pdf_name),
                        }
                        insertadas = _subir_factura_a_web_desde_local(
                            sp_excel_root=sp_excel,
                            archivos_ref=archivos_ref,
                            numeros_ref={str(reg.get("Número de factura") or numero_aprob or ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "").strip()},
                            expected_rows=int(total_nuevos or 0),
                            table_name="TblFacturas",
                            rows_dicts=[reg],
                        )
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

                cufe_final, numero_final = _resolver_cufe_numero_final(
                    reg_pdf=reg,
                    cufe_pdf=cufe_pdf,
                    numero_pdf=reg.get("Número de factura") or ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "",
                )
                
                radicado_subject = ""
                proyecto_subject = ""
                enriquecidas_subject = 0

                try:
                    regs_subject = [reg]
                    enriquecidas_subject, radicado_subject, proyecto_subject = _enriquecer_regs_desde_subject_si_falta(regs_subject, subj)

                    if enriquecidas_subject > 0:
                        reg = regs_subject[0]
                        print(
                            f"[SUBJECT] Fallback enriquecimiento DIAN/PDF en memoria: "
                            f"filas={enriquecidas_subject} | radicado={radicado_subject} | proyecto={proyecto_subject}"
                        )
                except Exception as e:
                    print(f"[SUBJECT] Error enriqueciendo DIAN/PDF desde subject: {e}")
                _push_detalle(
                    detalle_rows, run_id, msg_id, subj,
                    pdf_name=pdf_name,
                    cufe=cufe_final,
                    numero=numero_final,
                    fecha_factura=ident_pdf.get("FECHA") or fecha_pdf,
                    zip_match="(PDF-ONLY) VALIDACIONES DIAN",
                    estado="ok_dian_pdf_only",
                    duracion_s=secs,
                    nuevos=int(total_nuevos or 0),
                    enriquecidas=int(enriquecidas or 0),
                    fuente=("DIAN_PDF|SUBJECT" if (radicado_subject or proyecto_subject) else "DIAN_PDF")
                )

                cnt_dian += 1
                dian_total += 1
                if int(total_nuevos or 0) > 0:
                    dian_registradas += 1
                    facturas_con_filas += 1
                else:
                    dian_no_registrables += 1
                    facturas_sin_registro += 1

                msgs_procesados += 1
                nuevos_total += int(total_nuevos or 0)
                enriq_total += int(enriquecidas or 0)
                filas_local_total += int(total_nuevos or 0)
                filas_web_total += int(insertadas or 0)

                sin_match_consec = 0
                sin_nuevos_consec = 0 if total_nuevos > 0 else (sin_nuevos_consec + 1)
                procesados += 1
                continue

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
                regs_para_sharepoint: List[Dict[str, object]] = []

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
                    regs = _asegurar_regs_registrables_7_conceptos(regs)

                    if not regs:
                        print(
                            f"[DIAN ZIP][FALLBACK REGS] ZIP match sin regs útiles. "
                            f"Se fuerza registro mínimo desde PDF aprobado: {pdf_name}"
                        )
                        reg_min = _asegurar_reg_7_conceptos(_generar_registro_pdf_only(pdf_tmp, pdf_name))
                        numero_final_min = (
                            numero_aprob
                            or ident_pdf.get("NUMERO_APROB")
                            or ident_pdf.get("NUMERO")
                            or reg_min.get("Número de factura")
                            or ""
                        )
                        if numero_final_min and len(str(numero_final_min).strip()) >= 3:
                            reg_min["Número de factura"] = str(numero_final_min).strip()
                        regs = [reg_min]

                    if regs and numero_aprob:
                        for dct in regs:
                            old = str(dct.get("Número de factura", "")).strip()
                            if old != numero_aprob and len(numero_aprob) >= 3:
                                dct["Número de factura"] = numero_aprob

                    if regs:
                        regs, enriquecidas_forzadas_zip_tmp, radicado_final_zip_tmp, proyecto_final_zip_tmp = _forzar_radicado_y_proyecto_en_filas(
                            filas=regs,
                            subj=subj,
                            estado="ok_dian_zip",
                        )

                    if regs:
                        regs_para_sharepoint.extend(regs)

                    if regs:
                        for dct in regs:
                            av = dct.get("Archivo")
                            if av:
                                archivos_realmente_guardados.add(str(av).strip())

                    nuevos = guardar_en_excel(regs) if regs else 0
                    total_nuevos += nuevos
                    if regs and int(nuevos or 0) > 0:
                        _agregar_filas_al_buffer_web_run(rows_web_run_buffer, regs, origen="dian_zip")

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

                enriquecidas = int(total_nuevos or 0)
                if regs_para_sharepoint:
                    _, _enriquecidas_forzadas_zip, radicado_final_zip, proyecto_final_zip = _forzar_radicado_y_proyecto_en_filas(
                        filas=regs_para_sharepoint,
                        subj=subj,
                        estado="ok_dian_zip",
                    )
                    print(
                        f"[FORZADO] DIAN_ZIP -> radicado={radicado_final_zip} | "
                        f"proyecto={proyecto_final_zip} | enriquecidas={int(total_nuevos or 0)}"
                    )

                if int(total_nuevos or 0) <= 0:
                    print(f"[DIAN ZIP][FORZADO FINAL] Match con ZIP pero sin filas. Se fuerza registro mínimo obligatorio: {pdf_name}")
                    resultado_min = _registrar_minimo_obligatorio_desde_aprobadas(
                        msg_id=msg_id,
                        subj=subj,
                        pdf_name=pdf_name,
                        pdf_tmp=pdf_tmp,
                        ident_pdf=ident_pdf,
                        fecha_pdf=fecha_pdf,
                        fecha_local=fecha_local,
                        hora_local=hora_local,
                        numero_aprob=numero_aprob,
                        detalle_rows=detalle_rows,
                        run_id=run_id,
                        t0=t0,
                        usar_processed_store=usar_processed_store,
                        store=store,
                        motivo="match_zip_sin_filas",
                    )

                    secs = time.perf_counter() - t0
                    resumen.append((pdf_name, secs, "registro minimo obligatorio", int(resultado_min.get("nuevos", 0) or 0)))

                    cnt_dian += 1
                    dian_total += 1
                    if int(resultado_min.get("nuevos", 0) or 0) > 0:
                        dian_registradas += 1
                        facturas_con_filas += 1
                    else:
                        dian_no_registrables += 1
                        facturas_sin_registro += 1

                    msgs_procesados += 1
                    nuevos_total += int(resultado_min.get("nuevos", 0) or 0)
                    enriq_total += int(resultado_min.get("enriquecidas", 0) or 0)
                    filas_local_total += int(resultado_min.get("nuevos", 0) or 0)
                    filas_web_total += int(resultado_min.get("insertadas", 0) or 0)

                    sin_match_consec = 0
                    if int(resultado_min.get("nuevos", 0) or 0) == 0:
                        sin_nuevos_consec += 1
                    else:
                        sin_nuevos_consec = 0

                    procesados += 1
                    continue

                try:
                    sincronizar_aprobaciones_en_facturas()
                except Exception as e:
                    print(f"[APROB] Error al sincronizar aprobaciones: {e}")

                enriquecidas_subject = 0
                radicado_subject = ""
                proyecto_subject = ""
                try:
                    enriquecidas_subject, radicado_subject, proyecto_subject = _enriquecer_regs_desde_subject_si_falta(regs_para_sharepoint, subj)
                    if enriquecidas_subject > 0:
                        print(
                            f"[SUBJECT] Fallback enriquecimiento DIAN/ZIP en memoria: "
                            f"filas={enriquecidas_subject} | radicado={radicado_subject} | proyecto={proyecto_subject}"
                        )
                except Exception as e:
                    print(f"[SUBJECT] Error enriqueciendo DIAN/ZIP desde subject: {e}")

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

                        numeros_ref_web = {str(d.get("Número de factura") or "").strip() for d in regs_para_sharepoint if str(d.get("Número de factura") or "").strip()}
                        insertadas = _subir_factura_a_web_desde_local(
                            sp_excel_root=sp_excel,
                            archivos_ref=archivos_ref,
                            numeros_ref=numeros_ref_web,
                            expected_rows=int(total_nuevos or 0),
                            table_name="TblFacturas",
                            rows_dicts=regs_para_sharepoint,
                        )
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

                cufe_final, numero_final = _resolver_cufe_numero_final(
                    regs=regs_para_sharepoint,
                    cufe_pdf=cufe_pdf,
                    numero_pdf=ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "",
                )

                _push_detalle(
                    detalle_rows, run_id, msg_id, subj,
                    pdf_name=pdf_name,
                    cufe=cufe_final,
                    numero=numero_final,
                    fecha_factura=fecha_pdf,
                    zip_match=found_zip_name,
                    estado="ok_dian_zip",
                    duracion_s=secs,
                    nuevos=int(total_nuevos or 0),
                    enriquecidas=int(enriquecidas or 0),
                    fuente=("DIAN_ZIP|SUBJECT" if (radicado_subject or proyecto_subject) else "DIAN_ZIP")
                )

                cnt_dian += 1
                dian_total += 1
                if int(total_nuevos or 0) > 0:
                    dian_registradas += 1
                    facturas_con_filas += 1
                else:
                    dian_no_registrables += 1
                    facturas_sin_registro += 1

                msgs_procesados += 1
                nuevos_total += int(total_nuevos or 0)
                enriq_total += int(enriquecidas or 0)
                filas_local_total += int(total_nuevos or 0)
                filas_web_total += int(insertadas or 0)

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

            print(f"[DIAN] No encontré PDF/ZIP externo para {pdf_name}. Intentando fallback con el mismo PDF aprobado...")

            aplico_fallback, total_nuevos, enriquecidas, insertadas = _registrar_desde_pdf_aprobado_fallback(
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
                dian_total += 1
                if int(total_nuevos or 0) > 0:
                    dian_registradas += 1
                    facturas_con_filas += 1
                else:
                    dian_no_registrables += 1
                    facturas_sin_registro += 1

                msgs_procesados += 1
                nuevos_total += int(total_nuevos or 0)
                enriq_total += int(enriquecidas or 0)
                filas_local_total += int(total_nuevos or 0)
                filas_web_total += int(insertadas or 0)

                sin_match_consec = 0
                if total_nuevos == 0:
                    sin_nuevos_consec += 1
                else:
                    sin_nuevos_consec = 0

                procesados += 1
                continue

            motivo_dian = "sin_match_dian"
            if not cufe_pdf:
                motivo_dian = "sin_match_dian_pdf_sin_cufe"
            elif cufe_pdf and not _cufe_is_valid(cufe_pdf):
                motivo_dian = "sin_match_dian_cufe_debil"

            print(f"[DIAN][MINIMO FINAL] {motivo_dian}. Se fuerza registro mínimo obligatorio: {pdf_name}")
            resultado_min = _registrar_minimo_obligatorio_desde_aprobadas(
                msg_id=msg_id,
                subj=subj,
                pdf_name=pdf_name,
                pdf_tmp=pdf_tmp,
                ident_pdf=ident_pdf,
                fecha_pdf=fecha_pdf,
                fecha_local=fecha_local,
                hora_local=hora_local,
                numero_aprob=numero_aprob,
                detalle_rows=detalle_rows,
                run_id=run_id,
                t0=t0,
                usar_processed_store=usar_processed_store,
                store=store,
                motivo=motivo_dian,
            )

            total_nuevos_min = int(resultado_min.get("nuevos", 0) or 0)
            insertadas_min = int(resultado_min.get("insertadas", 0) or 0)
            enriquecidas_min = int(resultado_min.get("enriquecidas", 0) or 0)

            secs = time.perf_counter() - t0
            resumen.append((pdf_name, secs, "registro minimo obligatorio dian", total_nuevos_min))

            cnt_dian += 1
            dian_total += 1
            if total_nuevos_min > 0:
                dian_registradas += 1
                facturas_con_filas += 1
            else:
                dian_no_registrables += 1
                facturas_sin_registro += 1

            msgs_procesados += 1
            nuevos_total += total_nuevos_min
            enriq_total += enriquecidas_min
            filas_local_total += total_nuevos_min
            filas_web_total += insertadas_min

            sin_match_consec = 0
            sin_nuevos_consec = 0 if total_nuevos_min > 0 else (sin_nuevos_consec + 1)
            procesados += 1
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
            print("🔎 Intentando localizar correo origen con PDF único antes del fallback local...")

            try:
                pdf_origen_path, pdf_origen_name, source_msg_id, source_att_id = _buscar_correo_origen_con_pdf_unico(
                    msg_id_aprobado=msg_id,
                    subj_aprobado=subj,
                    ident_pdf=ident_pdf,
                    since_days=since_days,
                    top_msgs=80,
                )
                if pdf_origen_path and pdf_origen_name:
                    aplico_origen, nuevos_origen, enriquecidas_origen, insertadas_origen = _registrar_desde_pdf_origen_unico(
                        msg_id=msg_id,
                        subj=subj,
                        pdf_name_aprobado=pdf_name,
                        pdf_origen_path=pdf_origen_path,
                        pdf_origen_name=pdf_origen_name,
                        ident_pdf_aprobado=ident_pdf,
                        fecha_local=fecha_local,
                        hora_local=hora_local,
                        run_id=run_id,
                        detalle_rows=detalle_rows,
                        resumen=resumen,
                        t0=t0,
                        usar_processed_store=usar_processed_store,
                        store=store,
                        source_msg_id=source_msg_id or "",
                        source_att_id=source_att_id or "",
                    )
                    if aplico_origen:
                        cnt_ok += 1
                        ok_total += 1
                        if int(nuevos_origen or 0) > 0:
                            ok_registradas += 1
                            facturas_con_filas += 1
                        else:
                            ok_no_registrables += 1
                            facturas_sin_registro += 1

                        msgs_procesados += 1
                        nuevos_total += int(nuevos_origen or 0)
                        enriq_total += int(enriquecidas_origen or 0)
                        filas_local_total += int(nuevos_origen or 0)
                        filas_web_total += int(insertadas_origen or 0)

                        sin_match_consec = 0
                        if int(nuevos_origen or 0) == 0:
                            sin_nuevos_consec += 1
                        else:
                            sin_nuevos_consec = 0
                            if cufe_pdf:
                                cufes_existentes.add(cufe_pdf)
                                norm_cufes_existentes.add(cufe_pdf)

                        procesados += 1
                        continue
            except Exception as e:
                print(f"⚠️ Falló búsqueda/registro por PDF origen único para {pdf_name}: {e}")

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
                        ok_total += 1
                        if int(resultado_fallback.get("nuevos", 0) or 0) > 0:
                            ok_registradas += 1
                            facturas_con_filas += 1
                        else:
                            ok_no_registrables += 1
                            facturas_sin_registro += 1

                        msgs_procesados += 1
                        nuevos_total += int(resultado_fallback.get("nuevos", 0) or 0)
                        enriq_total += int(resultado_fallback.get("enriquecidas", 0) or 0)
                        filas_local_total += int(resultado_fallback.get("nuevos", 0) or 0)
                        filas_web_total += int(resultado_fallback.get("insertadas", 0) or 0)

                        sin_match_consec = 0
                        if int(resultado_fallback.get("nuevos", 0) or 0) == 0:
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
        regs_para_sharepoint: List[Dict[str, object]] = []

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
            regs = _asegurar_regs_registrables_7_conceptos(regs)

            if not regs:
                print(
                    f"[NORMAL ZIP][FALLBACK REGS] ZIP match sin regs útiles. "
                    f"Se fuerza registro mínimo desde PDF aprobado: {pdf_name}"
                )
                reg_min = _asegurar_reg_7_conceptos(_generar_registro_pdf_only(pdf_tmp, pdf_name))
                numero_final_min = (
                    numero_aprob
                    or ident_pdf.get("NUMERO_APROB")
                    or ident_pdf.get("NUMERO")
                    or reg_min.get("Número de factura")
                    or ""
                )
                if numero_final_min and len(str(numero_final_min).strip()) >= 3:
                    reg_min["Número de factura"] = str(numero_final_min).strip()
                regs = [reg_min]

            if regs and numero_aprob:
                for dct in regs:
                    old = str(dct.get("Número de factura", "")).strip()
                    if old != numero_aprob and len(numero_aprob) >= 3:
                        dct["Número de factura"] = numero_aprob

            if regs:
                regs, enriquecidas_forzadas_tmp, radicado_final_tmp, proyecto_final_tmp = _forzar_radicado_y_proyecto_en_filas(
                    filas=regs,
                    subj=subj,
                    estado="ok",
                )

            if regs:
                regs_para_sharepoint.extend(regs)

            if regs:
                for dct in regs:
                    av = dct.get("Archivo")
                    if av:
                        archivos_realmente_guardados.add(str(av).strip())

            nuevos = guardar_en_excel(regs) if regs else 0

            if regs:
                datos_subject_local = _parse_datos_desde_subject_aprobado(subj)
                radicado_local_force = str((regs[0].get("Radicado") if regs else "") or datos_subject_local.get("radicado_subject") or "").strip()
                proyecto_local_force = str((regs[0].get("ProyectoProceso") if regs else "") or datos_subject_local.get("proyecto_subject") or "").strip()
                archivos_force = {str(d.get("Archivo") or "").strip() for d in regs if str(d.get("Archivo") or "").strip()}
                numeros_force = {str(d.get("Número de factura") or "").strip() for d in regs if str(d.get("Número de factura") or "").strip()}
                filas_match_force, filas_upd_force = _forzar_campos_obligatorios_en_excel_local(
                    ARCHIVO_EXCEL,
                    archivos_ref=archivos_force,
                    numeros_ref=numeros_force,
                    radicado=radicado_local_force,
                    proyecto=proyecto_local_force,
                )
                if filas_upd_force <= 0 and int(nuevos or 0) > 0:
                    filas_match_force, filas_upd_force = _forzar_campos_obligatorios_en_ultimas_filas(
                        ARCHIVO_EXCEL,
                        expected_rows=int(nuevos or 0),
                        radicado=radicado_local_force,
                        proyecto=proyecto_local_force,
                    )
                if filas_upd_force > 0:
                    print(
                        f"[LOCAL FORCE] ZIP -> match={filas_match_force} | "
                        f"actualizadas={filas_upd_force} | radicado={radicado_local_force} | proyecto={proyecto_local_force}"
                    )

            total_nuevos += nuevos
            if regs and int(nuevos or 0) > 0:
                _agregar_filas_al_buffer_web_run(rows_web_run_buffer, regs, origen="normal_zip")

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

        enriquecidas = int(total_nuevos or 0)
        if regs_para_sharepoint:
            _, _enriquecidas_forzadas_zip, radicado_final_zip, proyecto_final_zip = _forzar_radicado_y_proyecto_en_filas(
                filas=regs_para_sharepoint,
                subj=subj,
                estado="ok",
            )
            print(
                f"[FORZADO] NORMAL_ZIP -> radicado={radicado_final_zip} | "
                f"proyecto={proyecto_final_zip} | enriquecidas={int(total_nuevos or 0)}"
            )

        if int(total_nuevos or 0) <= 0:
            print(f"[NORMAL ZIP][FORZADO FINAL] Match con ZIP pero sin filas. Se fuerza registro mínimo obligatorio: {pdf_name}")
            resultado_min = _registrar_minimo_obligatorio_desde_aprobadas(
                msg_id=msg_id,
                subj=subj,
                pdf_name=pdf_name,
                pdf_tmp=pdf_tmp,
                ident_pdf=ident_pdf,
                fecha_pdf=fecha_pdf,
                fecha_local=fecha_local,
                hora_local=hora_local,
                numero_aprob=numero_aprob,
                detalle_rows=detalle_rows,
                run_id=run_id,
                t0=t0,
                usar_processed_store=usar_processed_store,
                store=store,
                motivo="match_zip_sin_filas",
            )

            secs = time.perf_counter() - t0
            resumen.append((pdf_name, secs, "registro minimo obligatorio", int(resultado_min.get("nuevos", 0) or 0)))

            cnt_ok += 1
            ok_total += 1
            if int(resultado_min.get("nuevos", 0) or 0) > 0:
                ok_registradas += 1
                facturas_con_filas += 1
            else:
                ok_no_registrables += 1
                facturas_sin_registro += 1

            msgs_procesados += 1
            nuevos_total += int(resultado_min.get("nuevos", 0) or 0)
            enriq_total += int(resultado_min.get("enriquecidas", 0) or 0)
            filas_local_total += int(resultado_min.get("nuevos", 0) or 0)
            filas_web_total += int(resultado_min.get("insertadas", 0) or 0)

            sin_match_consec = 0
            if int(resultado_min.get("nuevos", 0) or 0) == 0:
                sin_nuevos_consec += 1
            else:
                sin_nuevos_consec = 0
                if cufe_pdf:
                    cufes_existentes.add(cufe_pdf)
                    norm_cufes_existentes.add(cufe_pdf)

            procesados += 1
            continue

        try:
            sincronizar_aprobaciones_en_facturas()
        except Exception as e:
            print(f"[APROB] Error al sincronizar aprobaciones: {e}")

        enriquecidas_subject = 0
        radicado_subject = ""
        proyecto_subject = ""
        try:
            enriquecidas_subject, radicado_subject, proyecto_subject = _enriquecer_regs_desde_subject_si_falta(regs_para_sharepoint, subj)
            if enriquecidas_subject > 0:
                print(
                    f"[SUBJECT] Fallback enriquecimiento NORMAL/ZIP en memoria: "
                    f"filas={enriquecidas_subject} | radicado={radicado_subject} | proyecto={proyecto_subject}"
                )
        except Exception as e:
            print(f"[SUBJECT] Error enriqueciendo NORMAL/ZIP desde subject: {e}")

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

                numeros_ref_web = {str(d.get("Número de factura") or "").strip() for d in regs_para_sharepoint if str(d.get("Número de factura") or "").strip()}
                insertadas = _subir_factura_a_web_desde_local(
                    sp_excel_root=sp_excel,
                    archivos_ref=archivos_ref,
                    numeros_ref=numeros_ref_web,
                    expected_rows=int(total_nuevos or 0),
                    table_name="TblFacturas",
                    rows_dicts=regs_para_sharepoint,
                )
                print(f"✅ Workbook API (NORMAL/ZIP): +{insertadas} fila(s) nuevas en TblFacturas.")
            except Exception as e:
                print(f"⚠️ Workbook API falló (NORMAL/ZIP): {e}")
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

        cufe_final, numero_final = _resolver_cufe_numero_final(
            regs=regs_para_sharepoint,
            cufe_pdf=cufe_pdf,
            numero_pdf=ident_pdf.get("NUMERO_APROB") or ident_pdf.get("NUMERO") or "",
        )

        fuente_detalle = fuente_match
        if radicado_subject or proyecto_subject:
            fuente_detalle = f"{fuente_match}|SUBJECT"

        _push_detalle(
            detalle_rows, run_id, msg_id, subj,
            pdf_name=pdf_name,
            cufe=cufe_final,
            numero=numero_final,
            fecha_factura=fecha_pdf,
            zip_match=found_zip_name,
            estado="ok",
            duracion_s=(time.perf_counter() - t0),
            nuevos=int(total_nuevos or 0),
            enriquecidas=int(enriquecidas or 0),
            fuente=fuente_detalle
        )

        cnt_ok += 1
        ok_total += 1
        if int(total_nuevos or 0) > 0:
            ok_registradas += 1
            facturas_con_filas += 1
        else:
            ok_no_registrables += 1
            facturas_sin_registro += 1

        msgs_procesados += 1
        nuevos_total += int(total_nuevos or 0)
        enriq_total += int(enriquecidas or 0)
        filas_local_total += int(total_nuevos or 0)
        filas_web_total += int(insertadas or 0)

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

    # Reconciliación final incremental LOCAL/RUN -> WEB.
    # 1) usa buffer real del run;
    # 2) agrega respaldo de últimas filas locales para cubrir helpers que no tienen acceso al buffer.
    try:
        sp_excel_final = f"{BASE_SP}/excel"
        filas_backup_local = _obtener_ultimas_filas_locales_para_reconciliacion_web(
            ARCHIVO_EXCEL,
            total_filas_run=int(nuevos_total or 0),
        )

        filas_reconciliacion = []
        if rows_web_run_buffer:
            filas_reconciliacion.extend(rows_web_run_buffer)
        if filas_backup_local:
            filas_reconciliacion.extend(filas_backup_local)

        if filas_reconciliacion:
            print(
                f"[WEB FINAL] Reconciliando web al cierre del run: "
                f"buffer={len(rows_web_run_buffer)} | backup_local={len(filas_backup_local)} | "
                f"total_envio={len(filas_reconciliacion)}"
            )
            insertadas_final = _reconciliar_web_desde_buffer_run(
                sp_excel_root=sp_excel_final,
                rows_web_run_buffer=filas_reconciliacion,
                table_name="TblFacturas",
            )
            if int(insertadas_final or 0) > 0:
                filas_web_total += int(insertadas_final or 0)
                print(
                    f"[WEB FINAL] filas_web_total actualizado={filas_web_total} | "
                    f"filas_local_total={filas_local_total} | nuevos_total={nuevos_total}"
                )
        else:
            print("[WEB FINAL] Sin filas para reconciliar al cierre.")
    except Exception as e:
        print(f"[WEB FINAL] Error no crítico en reconciliación final: {e}")

    total_secs = time.perf_counter() - t0_total
    fin_dt = datetime.datetime.now().isoformat(timespec="seconds")

    total_match = ok_total + dian_total

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
                "filas_local_total": filas_local_total,
                "filas_web_total": filas_web_total,

                # Resumen inteligente por factura (PASO 3.2)
                "total_match": total_match,
                "match_total": total_match,

                "ok_total": ok_total,
                "ok_match": ok_total,
                "ok_registradas": ok_registradas,
                "ok_con_filas": ok_registradas,
                "ok_no_registrables": ok_no_registrables,
                "ok_sin_filas": ok_no_registrables,

                "dian_total": dian_total,
                "dian_match": dian_total,
                "dian_registradas": dian_registradas,
                "dian_con_filas": dian_registradas,
                "dian_no_registrables": dian_no_registrables,
                "dian_sin_filas": dian_no_registrables,

                "facturas_con_filas": facturas_con_filas,
                "facturas_sin_registro": facturas_sin_registro,
                "facturas_sin_filas": facturas_sin_registro,

                "nota": "",
            })
        except Exception as e:
            print(f"⚠️ No pude escribir audit runs CSV: {e}")

    print("\n===== 📊 Resumen inteligente por factura =====")
    print(f"Procesadas: {msgs_procesados}")
    print(f"Match total: {total_match} | OK: {ok_total} | DIAN: {dian_total}")
    print(f"Facturas con filas: {facturas_con_filas} | Facturas sin filas: {facturas_sin_registro}")
    print(f"OK con filas: {ok_registradas} | OK sin filas: {ok_no_registrables}")
    print(f"DIAN con filas: {dian_registradas} | DIAN sin filas: {dian_no_registrables}")
    print("==============================================")

    print("\n===== ⏱️ Resumen de tiempos (aprobadas) =====")
    for name, secs, estado, nuevos in resumen:
        print(f"• {name} -> {secs:.2f}s | {estado} | nuevos={nuevos}")
    print(f"⏱️ Tiempo total real de ejecución: {total_secs:.2f} s")
    print("=============================================")

    try:
        lock.release()
    except Exception:
        pass