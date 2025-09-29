# tests/sp_list_drives.py
import os, requests
from dotenv import load_dotenv
from services.m365.token import get_access_token

load_dotenv()
GRAPH = "https://graph.microsoft.com/v1.0"
SITE_HOST = os.getenv("SP_HOSTNAME")          # joycocia.sharepoint.com
SITE_PATH = os.getenv("SP_SITE_PATH")         # /sites/Infraestructura

sess = requests.Session()
headers = {"Authorization": f"Bearer {get_access_token()}"}

# 1) Resolver el siteId
site_url = f"{GRAPH}/sites/{SITE_HOST}:{SITE_PATH}"
site = sess.get(site_url, headers=headers).json()
print("siteId:", site.get("id"))

# 2) Listar drives (bibliotecas de documentos) del sitio
drives = sess.get(f"{GRAPH}/sites/{site['id']}/drives", headers=headers).json()["value"]
print("\n== DRIVES EN EL SITIO ==")
for d in drives:
    print(f"- name: {d['name']:<30}  id: {d['id']}")
