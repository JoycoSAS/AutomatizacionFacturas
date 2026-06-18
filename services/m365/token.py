# services/m365/token.py

import os
import time
import requests
from dotenv import load_dotenv

# Cargar variables del .env
load_dotenv()

TENANT = (os.getenv("TENANT_ID") or "").strip()
CLIENT = (os.getenv("CLIENT_ID") or "").strip()
SECRET = (os.getenv("CLIENT_SECRET") or "").strip()

_TOKEN_CACHE = {"value": None, "exp": 0}

TOKEN_TIMEOUT_SECONDS = int(os.getenv("GRAPH_TOKEN_TIMEOUT_SECONDS", "30") or "30")
TOKEN_DEBUG = str(os.getenv("GRAPH_TOKEN_DEBUG", "0") or "0").strip().lower() in {
    "1", "true", "yes", "y", "si", "sí", "on"
}

TOKEN_SSL_VERIFY = str(os.getenv("SSL_VERIFY", "true") or "true").strip().lower() not in {
    "0", "false", "no", "off"
}


def _credenciales_configuradas() -> bool:
    """Valida que existan las variables mínimas sin imprimir valores sensibles."""
    faltantes = []
    if not TENANT:
        faltantes.append("TENANT_ID")
    if not CLIENT:
        faltantes.append("CLIENT_ID")
    if not SECRET:
        faltantes.append("CLIENT_SECRET")

    if faltantes:
        print("❌ Faltan variables requeridas para Microsoft Graph:", ", ".join(faltantes))
        return False

    return True


def _debug_token_seguro():
    """
    Diagnóstico opcional y seguro.
    No imprime TENANT_ID, CLIENT_ID ni longitud/valor del CLIENT_SECRET.
    Solo se activa con GRAPH_TOKEN_DEBUG=1.
    """
    if not TOKEN_DEBUG:
        return

    print("🔐 DEBUG TOKEN GRAPH ACTIVADO")
    print("TENANT_ID configurado:", bool(TENANT))
    print("CLIENT_ID configurado:", bool(CLIENT))
    print("CLIENT_SECRET configurado:", bool(SECRET))
    print("-" * 40)


def _mensaje_error_token(status_code: int, response_text: str = "") -> None:
    """Imprime un diagnóstico seguro sin exponer credenciales."""
    print(f"❌ Error al solicitar token de Microsoft Graph. HTTP {status_code}")

    if status_code == 400:
        print("Posibles causas: tenant, client_id, secret, permisos o scope mal configurados.")
    elif status_code == 401:
        print("Posibles causas: CLIENT_SECRET vencido/incorrecto, CLIENT_ID incorrecto o TENANT_ID incorrecto.")
    elif status_code == 403:
        print("Posible causa: la aplicación no tiene permisos/consentimiento suficiente en Azure/Entra ID.")
    elif status_code in (429, 500, 502, 503, 504):
        print("Posible causa: límite temporal, servicio no disponible o error transitorio de Microsoft Graph.")
    else:
        print("Revisa la configuración de Azure/Entra ID y las variables del .env.")

    # Solo mostrar cuerpo parcial si se habilita debug explícitamente.
    # No imprimir por defecto para evitar ruido o exposición innecesaria.
    if TOKEN_DEBUG and response_text:
        txt = str(response_text).replace("\n", " ").strip()
        print("Respuesta Azure parcial:", txt[:500])


def get_access_token() -> str:
    """
    Obtiene token de acceso para Microsoft Graph usando Client Credentials Flow.
    Incluye cache y salida segura para producción.
    """
    global _TOKEN_CACHE

    now = time.time()

    # Si el token sigue vigente, usar cache.
    if _TOKEN_CACHE["value"] and now < _TOKEN_CACHE["exp"] - 60:
        return _TOKEN_CACHE["value"]

    if not _credenciales_configuradas():
        raise RuntimeError("Credenciales de Microsoft Graph incompletas en .env")

    _debug_token_seguro()

    url = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"

    data = {
        "client_id": CLIENT,
        "client_secret": SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }

    try:
        r = requests.post(url, data=data, timeout=TOKEN_TIMEOUT_SECONDS, verify=TOKEN_SSL_VERIFY)

        if r.status_code >= 400:
            _mensaje_error_token(r.status_code, r.text)
            r.raise_for_status()

        js = r.json()
        access_token = js.get("access_token")

        if not access_token:
            raise RuntimeError("La respuesta de Microsoft Graph no incluyó access_token")

        expires_in = int(js.get("expires_in", 3600) or 3600)

        _TOKEN_CACHE["value"] = access_token
        _TOKEN_CACHE["exp"] = now + expires_in

        print("✅ Token Graph obtenido correctamente")
        return access_token

    except requests.exceptions.RequestException as e:
        print("❌ Error de red al solicitar token de Graph")
        print("Detalle:", str(e))
        raise
    except Exception as e:
        print("❌ Error inesperado al obtener token de Graph")
        print("Detalle:", str(e))
        raise


print("🔥 TOKEN GRAPH PATCH 2026-06-05 ACTIVO: SAFE-CREDENTIAL-LOGS")
