# services/m365/token.py

import os
import time
import requests
from dotenv import load_dotenv

# Cargar variables del .env
load_dotenv()

TENANT = os.getenv("TENANT_ID")
CLIENT = os.getenv("CLIENT_ID")
SECRET = os.getenv("CLIENT_SECRET")

_TOKEN_CACHE = {"value": None, "exp": 0}


def _debug_credentials():
    """Imprime diagnóstico seguro de credenciales."""
    print("🔐 DEBUG CREDENCIALES GRAPH")
    print("TENANT_ID:", TENANT)
    print("CLIENT_ID:", CLIENT)
    print("CLIENT_SECRET cargado:", bool(SECRET))
    print("CLIENT_SECRET longitud:", len(SECRET or ""))
    print("-" * 40)


def get_access_token() -> str:
    """
    Obtiene token de acceso para Microsoft Graph usando Client Credentials Flow.
    Incluye cache y diagnóstico de errores.
    """
    global _TOKEN_CACHE

    now = time.time()

    # Si el token sigue vigente, usar cache
    if _TOKEN_CACHE["value"] and now < _TOKEN_CACHE["exp"] - 60:
        return _TOKEN_CACHE["value"]

    # DEBUG (solo cuando pide token nuevo)
    _debug_credentials()

    url = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"

    data = {
        "client_id": CLIENT,
        "client_secret": SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }

    try:
        r = requests.post(url, data=data, timeout=30)

        # Manejo especial para errores comunes
        if r.status_code == 401:
            print("❌ ERROR 401 - Unauthorized al solicitar token.")
            print("Esto casi siempre significa:")
            print("   • CLIENT_SECRET vencido")
            print("   • CLIENT_SECRET incorrecto")
            print("   • CLIENT_ID incorrecto")
            print("   • TENANT_ID incorrecto")
            print("Respuesta Azure:", r.text)
            raise Exception("401 Unauthorized al obtener token de Microsoft Graph")

        r.raise_for_status()

        js = r.json()

        access_token = js["access_token"]
        expires_in = int(js.get("expires_in", 3600))

        _TOKEN_CACHE["value"] = access_token
        _TOKEN_CACHE["exp"] = now + expires_in

        print("✅ Token Graph obtenido correctamente")

        return access_token

    except requests.exceptions.RequestException as e:
        print("❌ Error de red al solicitar token de Graph")
        print("Detalle:", str(e))
        raise