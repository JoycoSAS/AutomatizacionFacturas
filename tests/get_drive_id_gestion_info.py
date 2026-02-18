import os
import json
import requests
from dotenv import load_dotenv

from services.m365.token import get_access_token

load_dotenv()
GRAPH = "https://graph.microsoft.com/v1.0"
HOSTNAME = os.getenv("SP_HOSTNAME", "joycocia.sharepoint.com")

def h():
    return {"Authorization": f"Bearer {get_access_token()}", "Accept": "application/json"}

def main():
    query = "Gestión información"  # puedes probar también: "Gestion informacion"
    url = f"{GRAPH}/sites?search={requests.utils.quote(query)}"

    r = requests.get(url, headers=h(), timeout=60)
    if r.status_code >= 400:
        print("❌ Error buscando sitios:", r.status_code, r.text)
        return

    data = r.json()
    sites = data.get("value", [])

    if not sites:
        print("❌ No encontré sitios con search. Prueba otra palabra (ej: 'Gestion').")
        return

    print("\n=== SITIOS ENCONTRADOS ===")
    for i, s in enumerate(sites, 1):
        print("-" * 70)
        print(f"[{i}] name  : {s.get('name')}")
        print(f"    id    : {s.get('id')}")
        print(f"    webUrl : {s.get('webUrl')}")
    print("-" * 70)

    # Elige el que tenga webUrl relacionado a “Gestión información”
    # Luego copia su ID y prueba drives:
    chosen = sites[0]
    site_id = chosen.get("id")

    print(f"\n👉 Probando drives del site_id: {site_id}")
    durl = f"{GRAPH}/sites/{site_id}/drives"
    r2 = requests.get(durl, headers=h(), timeout=60)
    if r2.status_code >= 400:
        print("❌ Error listando drives:", r2.status_code, r2.text)
        return

    drives = r2.json().get("value", [])
    print("\n=== DRIVES (Bibliotecas) DEL SITIO ===")
    for d in drives:
        print("-" * 70)
        print("name     :", d.get("name"))
        print("id       :", d.get("id"))
        print("webUrl   :", d.get("webUrl"))
        print("driveType:", d.get("driveType"))

    print("\n✅ Copia el 'id' del drive que corresponda a:")
    print("   '01 Correspondencia' donde está Control correspondencia Oficina Principal.xlsx")

if __name__ == "__main__":
    main()
