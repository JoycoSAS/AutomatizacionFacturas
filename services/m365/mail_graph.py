import os
import base64
import time
import requests
from urllib.parse import quote
from dotenv import load_dotenv
from datetime import datetime, timedelta, timezone
from typing import Optional

from .token import get_access_token
from config import STORE_NAME

load_dotenv()

GRAPH = "https://graph.microsoft.com/v1.0"

MAILBOX = (
    os.getenv("GRAPH_USER")
    or os.getenv("GRAPH_MAILBOX")
    or os.getenv("MAILBOX_UPN")
    or STORE_NAME
)

if not MAILBOX or not isinstance(MAILBOX, str):
    raise RuntimeError(
        "No hay buzón. Define GRAPH_USER/GRAPH_MAILBOX/MAILBOX_UPN en .env "
        "o usa STORE_NAME en config.py"
    )

# Timeout estándar para llamadas JSON simples
TIMEOUT = (10, 30)  # connect, read

# Timeout para adjuntos / mensajes un poco más pesados
ATTACH_TIMEOUT = (10, 40)

# Timeout para descarga puntual de adjunto grande
DOWNLOAD_TIMEOUT = (15, 90)

# Límite de páginas de adjuntos por mensaje
MAX_ATTACHMENT_PAGES = 6

# Límite de páginas de mensajes por consulta
MAX_MESSAGE_PAGES = int(os.getenv("GRAPH_MAX_MESSAGE_PAGES", "120"))

# Tamaño de página para mensajes
GRAPH_PAGE_SIZE = max(1, min(int(os.getenv("GRAPH_PAGE_SIZE", "200")), 999))

# Pausa mínima entre llamadas Graph
PAUSE_BETWEEN_GRAPH_CALLS = float(os.getenv("GRAPH_PAUSE_SECONDS", "0.03"))

# Reintentos globales
DEFAULT_RETRIES = int(os.getenv("GRAPH_RETRIES", "2"))

_SESSION = requests.Session()
_SESSION.headers.update({"Accept": "application/json"})


def _h(content_type: str | None = None, extra: dict | None = None) -> dict:
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    if content_type:
        headers["Content-Type"] = content_type
    if extra:
        headers.update(extra)
    return headers


def _user_segment() -> str:
    return f"users/{quote(MAILBOX, safe='')}"


def _sleep_graph():
    if PAUSE_BETWEEN_GRAPH_CALLS > 0:
        time.sleep(PAUSE_BETWEEN_GRAPH_CALLS)


def _close_response(response):
    try:
        if response is not None:
            response.close()
    except Exception:
        pass


def _safe_json_response(response):
    try:
        return response.json()
    except Exception as e:
        print(f"[Graph] Respuesta no-JSON: {e}")
        try:
            txt = (response.text or "").strip()
            if txt:
                print(f"[Graph] Cuerpo parcial: {txt[:300]}")
        except Exception:
            pass
        return None


def _get_json_with_retries(
    url: str,
    retries: int = DEFAULT_RETRIES,
    timeout=TIMEOUT,
    headers_extra: dict | None = None
):
    delay = 1.2

    for attempt in range(retries + 1):
        response = None
        try:
            response = _SESSION.get(
                url,
                headers=_h(extra=headers_extra),
                timeout=timeout,
            )

            if not response.ok:
                print(f"[Graph] HTTP {response.status_code} en {url}")
                data_err = _safe_json_response(response)
                if data_err:
                    print(f"[Graph] Error detallado: {data_err}")

                if response.status_code in (429, 500, 502, 503, 504):
                    if attempt < retries:
                        print(f"[Graph] Reintentando por HTTP {response.status_code}...")
                        time.sleep(delay)
                        delay *= 2
                        continue

                return None

            data = _safe_json_response(response)
            if data is not None:
                _sleep_graph()
                return data

        except (requests.Timeout, requests.ConnectionError) as e:
            print(f"[Graph] Timeout/conexión en GET {url}: {e}")
        except requests.RequestException as e:
            print(f"[Graph] RequestException en GET {url}: {e}")
        except Exception as e:
            print(f"[Graph] Error inesperado en GET {url}: {e}")
        finally:
            _close_response(response)

        if attempt < retries:
            print(f"[Graph] Reintentando ({attempt + 1}/{retries}) en {delay:.1f}s…")
            time.sleep(delay)
            delay *= 2

    return None


def _env_int(name: str, default: int | None = None) -> int | None:
    raw = (os.getenv(name) or "").strip()
    if raw.isdigit():
        return int(raw)
    return default


def _env_bool(name: str, default: bool = False) -> bool:
    raw = (os.getenv(name) or "").strip().lower()
    if raw in ("1", "true", "yes", "y", "si", "sí", "on"):
        return True
    if raw in ("0", "false", "no", "n", "off"):
        return False
    return default


def _build_messages_url(
    base_url: str,
    page_size: int,
    since_days: int | None = None,
    unread_only: bool | None = None,
) -> str:
    params = {
        "$select": "id,subject,hasAttachments,categories,receivedDateTime,conversationId,isRead,bodyPreview",
        "$orderby": "receivedDateTime desc",
        "$top": str(page_size),
    }

    filtros = []

    if since_days is not None and since_days > 0:
        dt = datetime.now(timezone.utc) - timedelta(days=int(since_days))
        iso = dt.isoformat().replace("+00:00", "Z")
        filtros.append(f"receivedDateTime ge {iso}")

    if unread_only is True:
        filtros.append("isRead eq false")

    if filtros:
        params["$filter"] = " and ".join(filtros)

    query = "&".join(
        [f"{k}={quote(v, safe='(),:$ ')}" for k, v in params.items()]
    )
    return f"{base_url}?{query}"


def _listar_mensajes_paginado(
    base_url: str,
    max_messages: int = 200,
    since_days: int | None = None,
    unread_only: bool | None = None,
) -> list:
    max_env = _env_int("MAX_MESSAGES", 200) or 200
    target = min(max(1, int(max_messages or 200)), max_env)

    page_size = min(GRAPH_PAGE_SIZE, target)
    next_url = _build_messages_url(
        base_url=base_url,
        page_size=page_size,
        since_days=since_days,
        unread_only=unread_only,
    )

    out = []
    page_count = 0

    while next_url and len(out) < target:
        page_count += 1

        if page_count > MAX_MESSAGE_PAGES:
            print(f"[Graph] Corte preventivo: demasiadas páginas de mensajes ({page_count}) en {base_url}")
            break

        data = _get_json_with_retries(next_url, retries=2, timeout=TIMEOUT)
        if not data:
            break

        values = data.get("value", []) or []
        if not values:
            break

        faltan = target - len(out)
        out.extend(values[:faltan])

        if len(out) >= target:
            break

        next_url = data.get("@odata.nextLink")
        _sleep_graph()

    return out


def _listar_mensajes(max_messages: int = 200, since_days: int | None = None) -> list:
    """
    Mantiene compatibilidad histórica.
    Ojo: este método lee mensajes del buzón general.
    Para búsquedas de Inbox real se usan funciones específicas.
    """
    base = f"{GRAPH}/{_user_segment()}/messages"
    return _listar_mensajes_paginado(
        base_url=base,
        max_messages=max_messages,
        since_days=since_days,
        unread_only=None,
    )


def _listar_mensajes_inbox(max_messages: int = 200, since_days: int | None = None) -> list:
    base = f"{GRAPH}/{_user_segment()}/mailFolders/inbox/messages"
    return _listar_mensajes_paginado(
        base_url=base,
        max_messages=max_messages,
        since_days=since_days,
        unread_only=None,
    )


def buscar_mensajes_inbox_por_asunto(
    asunto_contiene: str,
    top: int = 50,
    since_days: int | None = 7,
    solo_con_adjuntos: bool = True,
):
    top = int(top or 50)
    if top < 1:
        return []

    # Traemos más de lo pedido para permitir filtrar por asunto después
    # y no quedarnos cortos por resultados no coincidentes.
    lote_busqueda = max(top * 4, 200)

    msgs = _listar_mensajes_inbox(
        max_messages=lote_busqueda,
        since_days=since_days
    )

    if not (asunto_contiene or "").strip():
        if solo_con_adjuntos:
            msgs = [m for m in msgs if m.get("hasAttachments")]
        return msgs[:top]

    asunto_contiene = asunto_contiene.strip().lower()

    out = []
    for msg in msgs:
        subj = (msg.get("subject") or "").lower()
        if asunto_contiene in subj:
            if solo_con_adjuntos and not msg.get("hasAttachments"):
                continue
            out.append(msg)
            if len(out) >= top:
                break

    return out


def _listar_adjuntos(msg_id: str) -> list:
    mid = quote(msg_id, safe="")
    next_url = f"{GRAPH}/{_user_segment()}/messages/{mid}/attachments?$top=50"
    out = []
    page_count = 0

    while next_url:
        page_count += 1

        if page_count > MAX_ATTACHMENT_PAGES:
            print(f"[Graph] Corte preventivo: demasiadas páginas de adjuntos para msg_id={msg_id}")
            break

        print(f"[Graph] Listando adjuntos msg_id={msg_id} página={page_count}")

        data = _get_json_with_retries(
            next_url,
            retries=1,
            timeout=ATTACH_TIMEOUT
        )

        if not data:
            print(f"[Graph] No se pudieron listar adjuntos para msg_id={msg_id}")
            break

        values = data.get("value", []) or []
        out.extend(values)

        next_url = data.get("@odata.nextLink")
        _sleep_graph()

    return out


def _es_file_attachment(att: dict) -> bool:
    odata_type = str(att.get("@odata.type") or "").lower()
    return (
        "#microsoft.graph.fileattachment" in odata_type
        or odata_type.endswith("fileattachment")
    )


def _listar_adjuntos_zip(msg_id: str) -> list:
    items = _listar_adjuntos(msg_id)
    out = []

    for att in items:
        if not _es_file_attachment(att):
            continue

        name = str(att.get("name") or "").strip()
        cty = str(att.get("contentType") or "").lower().strip()

        is_zip = (
            name.lower().endswith(".zip")
            or cty in ("application/zip", "application/x-zip-compressed")
            or "zip" in cty
        )

        if is_zip:
            out.append(att)

    return out


def listar_adjuntos_pdf(msg_id: str) -> list:
    items = _listar_adjuntos(msg_id)
    pdfs = []

    for att in items:
        if not _es_file_attachment(att):
            continue

        name = str(att.get("name") or "").lower().strip()
        cty = str(att.get("contentType") or "").lower().strip()

        if name.endswith(".pdf") or "pdf" in cty:
            pdfs.append(att)

    return pdfs


def _descargar_adjunto(msg_id: str, att_id: str, dest_path: str) -> bool:
    mid = quote(msg_id, safe="")
    aid = quote(att_id, safe="")
    url = f"{GRAPH}/{_user_segment()}/messages/{mid}/attachments/{aid}"

    delay = 1.2

    for attempt in range(DEFAULT_RETRIES + 1):
        response = None
        try:
            response = _SESSION.get(
                url,
                headers=_h(),
                timeout=DOWNLOAD_TIMEOUT,
            )

            if not response.ok:
                print(f"[Graph] HTTP {response.status_code} descargando adjunto {att_id}")
                data_err = _safe_json_response(response)
                if data_err:
                    print(f"[Graph] Error detallado adjunto: {data_err}")

                if response.status_code in (429, 500, 502, 503, 504) and attempt < DEFAULT_RETRIES:
                    time.sleep(delay)
                    delay *= 2
                    continue

                return False

            data = _safe_json_response(response)
            if not data:
                return False

            if not _es_file_attachment(data):
                print(f"[Graph] El adjunto {att_id} no es fileAttachment.")
                return False

            content = data.get("contentBytes")
            if not content:
                print("[Graph] El adjunto no trae contentBytes.")
                return False

            try:
                raw = base64.b64decode(content)
            except Exception as e:
                print(f"[Graph] contentBytes inválido: {e}")
                return False

            try:
                folder = os.path.dirname(dest_path)
                if folder:
                    os.makedirs(folder, exist_ok=True)

                with open(dest_path, "wb") as f:
                    f.write(raw)

                _sleep_graph()
                return True
            except Exception as e:
                print(f"[FS] No se pudo escribir el adjunto en {dest_path}: {e}")
                return False

        except (requests.Timeout, requests.ConnectionError) as e:
            print(f"[Graph] Timeout/conexión descargando adjunto {att_id}: {e}")
        except requests.RequestException as e:
            print(f"[Graph] RequestException descargando adjunto {att_id}: {e}")
        except Exception as e:
            print(f"[Graph] Error inesperado descargando adjunto {att_id}: {e}")
        finally:
            _close_response(response)

        if attempt < DEFAULT_RETRIES:
            print(f"[Graph] Reintentando descarga adjunto ({attempt + 1}/{DEFAULT_RETRIES}) en {delay:.1f}s…")
            time.sleep(delay)
            delay *= 2

    return False


def descargar_adjunto_por_id(msg_id: str, att_id: str, dest_path: str) -> bool:
    return _descargar_adjunto(msg_id, att_id, dest_path)


def listar_mensajes_zip_inbox(top: int = 300, since_days: int | None = None) -> list:
    msgs = _listar_mensajes_inbox(max_messages=top, since_days=since_days)
    return [m for m in msgs if m.get("hasAttachments")]


def listar_adjuntos_zip(msg_id: str) -> list:
    try:
        zips = _listar_adjuntos_zip(msg_id)
    except Exception as e:
        print(f"[DEBUG ZIP] Error listando adjuntos ZIP para msg_id={msg_id}: {e}")
        return []

    try:
        print(
            f"[DEBUG ZIP] msg_id={msg_id} | "
            f"zips encontrados={len(zips)} | "
            f"nombres={[z.get('name') for z in zips]}"
        )
    except Exception:
        pass

    return zips


def get_folder_id_by_name(root_display: str, name: str) -> str | None:
    inbox_url = f"{GRAPH}/{_user_segment()}/mailFolders/inbox"
    data = _get_json_with_retries(inbox_url, retries=2, timeout=TIMEOUT)
    if not data:
        return None

    root_id = data.get("id")
    if not root_id:
        return None

    childs_url = f"{GRAPH}/{_user_segment()}/mailFolders/{quote(root_id, safe='')}/childFolders?$top=500"
    data = _get_json_with_retries(childs_url, retries=2, timeout=TIMEOUT)
    if not data:
        return None

    for item in data.get("value", []):
        if (item.get("displayName") or "").strip().lower() == name.strip().lower():
            return item["id"]

    return None


def find_folder_id_anywhere(name: str) -> str | None:
    next_url = f"{GRAPH}/{_user_segment()}/mailFolders?$top=1000&$select=id,displayName"

    while next_url:
        data = _get_json_with_retries(next_url, retries=2, timeout=TIMEOUT)
        if not data:
            return None

        for item in data.get("value", []):
            if (item.get("displayName") or "").strip().lower() == name.strip().lower():
                return item["id"]

        next_url = data.get("@odata.nextLink")

    return None


def listar_mensajes_en_carpeta(
    folder_id: str,
    top: int = 200,
    unread_only: bool | None = None,
    since_days: int | None = None,
):
    fid = quote(folder_id, safe="")
    url = f"{GRAPH}/{_user_segment()}/mailFolders/{fid}/messages"

    if unread_only is None:
        unread_only = _env_bool("MAIL_UNREAD_ONLY", True)

    if since_days is None:
        since_days = _env_int("MAIL_LOOKBACK_DAYS", None)

    return _listar_mensajes_paginado(
        base_url=url,
        max_messages=top,
        since_days=since_days,
        unread_only=unread_only,
    )


def marcar_mensaje_como_leido(msg_id: str) -> bool:
    response = None
    try:
        mid = quote(msg_id, safe="")
        url = f"{GRAPH}/{_user_segment()}/messages/{mid}"
        payload = {"isRead": True}

        response = _SESSION.patch(
            url,
            headers=_h("application/json"),
            json=payload,
            timeout=TIMEOUT,
        )

        if response.status_code >= 400:
            try:
                body = response.json()
            except Exception:
                body = response.text
            print(f"[Graph] No se pudo marcar como leído ({response.status_code}): {body}")
            return False

        return True
    except Exception as e:
        print(f"[Graph] Error al marcar como leído: {e}")
        return False
    finally:
        _close_response(response)