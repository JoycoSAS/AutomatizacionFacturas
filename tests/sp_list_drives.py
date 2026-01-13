# tests/sp_list_drives.py
import os
import requests
import urllib3
from dotenv import load_dotenv

from services.m365.token import get_access_token


def _env_bool(name: str, default: bool = True) -> bool:
    v = os.getenv(name)
    if v is None:
        return default
    return v.strip().lower() in ("1", "true", "yes", "y", "on")


def main():
    load_dotenv()

    GRAPH = "https://graph.microsoft.com/v1.0"
    SITE_HOST = os.getenv("SP_HOSTNAME")        # ej: joycocia.sharepoint.com
    SITE_PATH = os.getenv("SP_SITE_PATH")       # ej: /sites/ComunicacionesyMercadeo-Innovacion
    ssl_verify = _env_bool("SSL_VERIFY", True)

    if not SITE_HOST or not SITE_PATH:
        raise SystemExit("❌ Falta SP_HOSTNAME o SP_SITE_PATH en .env")

    if not ssl_verify:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    sess = requests.Session()
    headers = {"Authorization": f"Bearer {get_access_token()}"}

    # 1) Resolver el siteId
    site_url = f"{GRAPH}/sites/{SITE_HOST}:{SITE_PATH}"
    r = sess.get(site_url, headers=headers, timeout=(15, 60), verify=ssl_verify)
    r.raise_for_status()
    site = r.json()
    print(f"✅ siteId: {site.get('id')}")

    # 2) Listar drives (bibliotecas)
    r = sess.get(f"{GRAPH}/sites/{site['id']}/drives", headers=headers, timeout=(15, 60), verify=ssl_verify)
    r.raise_for_status()
    drives = r.json().get("value", [])

    print("\n== DRIVES EN EL SITIO ==")
    for d in drives:
        print(f"- name: {d.get('name',''):<30}  id: {d.get('id','')}")


if __name__ == "__main__":
    main()
