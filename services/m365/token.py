# services/m365/token.py
import os, time, requests
from dotenv import load_dotenv

load_dotenv()

TENANT = os.getenv("TENANT_ID")
CLIENT = os.getenv("CLIENT_ID")
SECRET = os.getenv("CLIENT_SECRET")

_TOKEN_CACHE = {"value": None, "exp": 0}

def get_access_token() -> str:
    """Client Credentials flow para Microsoft Graph."""
    global _TOKEN_CACHE
    now = time.time()
    if _TOKEN_CACHE["value"] and now < _TOKEN_CACHE["exp"] - 60:
        return _TOKEN_CACHE["value"]

    url = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT,
        "client_secret": SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()
    js = r.json()
    _TOKEN_CACHE["value"] = js["access_token"]
    _TOKEN_CACHE["exp"]   = now + int(js.get("expires_in", 3600))
    return _TOKEN_CACHE["value"]
