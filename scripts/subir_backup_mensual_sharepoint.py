import sys
import os
import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    import config  # noqa: F401
except Exception:
    pass

from services.m365.sp_graph import (
    upload_small_file,
    ensure_folder,
    get_item_by_path,
    SP_FOLDER as BASE_SP,
)

VERSION_UPLOAD = "2026-06-11-UPLOAD-BACKUP-MENSUAL-SP-V3-DOBLE-RUTA"

DATA_DIR = ROOT / "data"
BACKUPS_DIR = DATA_DIR / "backups_mensuales"

NOW = datetime.datetime.now()
MES = NOW.strftime("%Y-%m")
MES_FILE = NOW.strftime("%Y_%m")

MES_DIR = BACKUPS_DIR / MES

SP_BACKUP_DIR_PRINCIPAL = f"{BASE_SP}/Backups/02_Backups_Mensuales/{MES}".strip("/")

SP_BACKUP2_DRIVE_ID = (os.getenv("SP_BACKUP2_DRIVE_ID") or "").strip()
SP_BACKUP2_FOLDER = (os.getenv("SP_BACKUP2_FOLDER") or "").strip().strip("/")

SP_BACKUP_DIR_SECUNDARIA = (
    f"{SP_BACKUP2_FOLDER}/02_Backups_Mensuales/{MES}".strip("/")
    if SP_BACKUP2_FOLDER
    else ""
)


def buscar_ultimo_archivo(pattern: str) -> Path | None:
    if not MES_DIR.exists():
        return None

    archivos = sorted(
        MES_DIR.glob(pattern),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    return archivos[0] if archivos else None


def verificar_archivo(local_path: Path, sp_path: str, drive_id: str | None = None) -> bool:
    try:
        item = get_item_by_path(sp_path, drive_id=drive_id)
        size_sp = int(item.get("size", -1))
        size_local = local_path.stat().st_size

        if size_sp != size_local:
            print(f"❌ Verificación fallida: tamaño diferente para {local_path.name}")
            print(f"   Local: {size_local}")
            print(f"   SP:    {size_sp}")
            return False

        print(f"✅ Verificado por tamaño: {local_path.name} ({size_local} bytes)")
        return True

    except Exception as e:
        print(f"❌ No se pudo verificar {local_path.name}: {e}")
        return False


def subir_archivo(local_path: Path, sp_dir: str, nombre_destino: str, drive_id: str | None = None) -> bool:
    if not local_path or not local_path.exists():
        print(f"⚠️ No existe archivo local: {local_path}")
        return False

    sp_path = f"{sp_dir}/{local_path.name}".strip("/")

    print("☁️ Subiendo a SharePoint:")
    print(f"   Destino: {nombre_destino}")
    print(f"   Local:   {local_path}")
    print(f"   SP:      {sp_path}")

    try:
        upload_small_file(str(local_path), sp_path, mode="replace", drive_id=drive_id)
        print(f"✅ Subido correctamente: {local_path.name}")
    except Exception as e:
        print(f"❌ Error subiendo {local_path.name} a {nombre_destino}: {e}")
        return False

    return verificar_archivo(local_path, sp_path, drive_id=drive_id)


def preparar_destinos() -> list[dict]:
    destinos = [
        {
            "nombre": "PRINCIPAL_CONTABILIDAD",
            "sp_dir": SP_BACKUP_DIR_PRINCIPAL,
            "drive_id": None,
        }
    ]

    if not SP_BACKUP2_DRIVE_ID:
        print("❌ Falta SP_BACKUP2_DRIVE_ID en .env")
        return []

    if not SP_BACKUP2_FOLDER:
        print("❌ Falta SP_BACKUP2_FOLDER en .env")
        return []

    destinos.append(
        {
            "nombre": "SECUNDARIA_CONTROL_INTERNO",
            "sp_dir": SP_BACKUP_DIR_SECUNDARIA,
            "drive_id": SP_BACKUP2_DRIVE_ID,
        }
    )

    return destinos


def main() -> int:
    print("=" * 100)
    print("SUBIDA BACKUP MENSUAL A SHAREPOINT - FACTURAS JOYCO")
    print("=" * 100)
    print(f"Versión: {VERSION_UPLOAD}")
    print(f"Root: {ROOT}")
    print(f"Mes: {MES}")
    print(f"Carpeta local: {MES_DIR}")
    print(f"Ruta principal:   {SP_BACKUP_DIR_PRINCIPAL}")
    print(f"Ruta secundaria:  {SP_BACKUP_DIR_SECUNDARIA}")
    print("-" * 100)

    zip_file = buscar_ultimo_archivo(f"backup_mensual_{MES_FILE}_*.zip")
    manifest_file = buscar_ultimo_archivo(f"manifest_backup_mensual_{MES_FILE}_*.json")

    if not zip_file:
        print("❌ No se encontró ZIP mensual para subir.")
        return 1

    archivos = [zip_file]

    if manifest_file:
        archivos.append(manifest_file)
    else:
        print("⚠️ No se encontró manifest mensual. Se subirá solo el ZIP.")

    destinos = preparar_destinos()

    if not destinos:
        print("❌ No hay destinos válidos para subir.")
        return 1

    ok_total = True

    for destino in destinos:
        nombre = destino["nombre"]
        sp_dir = destino["sp_dir"]
        drive_id = destino["drive_id"]

        print("-" * 100)
        print(f"📁 Verificando/creando destino: {nombre}")
        print(f"   SP_DIR: {sp_dir}")

        try:
            ensure_folder(sp_dir, drive_id=drive_id)
            print(f"✅ Carpeta SharePoint verificada/creada: {nombre}")
        except Exception as e:
            print(f"❌ No se pudo verificar/crear carpeta SharePoint en {nombre}: {e}")
            ok_total = False
            continue

        for archivo in archivos:
            ok = subir_archivo(archivo, sp_dir, nombre, drive_id=drive_id)
            if not ok:
                ok_total = False

    print("-" * 100)

    if ok_total:
        print("✅ Subida mensual a SharePoint terminada correctamente en ambas rutas.")
        print("=" * 100)
        return 0

    print("❌ Subida mensual terminó con errores en una o más rutas.")
    print("=" * 100)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
