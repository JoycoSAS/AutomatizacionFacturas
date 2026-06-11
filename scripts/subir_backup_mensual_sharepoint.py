import sys
import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from services.m365.sp_graph import (
    upload_small_file,
    ensure_folder,
    SP_FOLDER as BASE_SP,
)

VERSION_UPLOAD = "2026-06-11-UPLOAD-BACKUP-MENSUAL-SP-V2-SYSPATH"

DATA_DIR = ROOT / "data"
BACKUPS_DIR = DATA_DIR / "backups_mensuales"

NOW = datetime.datetime.now()
MES = NOW.strftime("%Y-%m")
MES_FILE = NOW.strftime("%Y_%m")

MES_DIR = BACKUPS_DIR / MES
SP_BACKUP_DIR = f"{BASE_SP}/Backups/02_Backups_Mensuales/{MES}".strip("/")


def buscar_ultimo_archivo(pattern: str) -> Path | None:
    if not MES_DIR.exists():
        return None

    archivos = sorted(
        MES_DIR.glob(pattern),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    return archivos[0] if archivos else None


def subir_archivo(local_path: Path, sp_dir: str) -> bool:
    if not local_path or not local_path.exists():
        print(f"⚠️ No existe archivo local: {local_path}")
        return False

    sp_path = f"{sp_dir}/{local_path.name}".strip("/")

    print("☁️ Subiendo a SharePoint:")
    print(f"   Local: {local_path}")
    print(f"   SP:    {sp_path}")

    try:
        upload_small_file(str(local_path), sp_path, mode="replace")
        print(f"✅ Subido correctamente: {local_path.name}")
        return True
    except Exception as e:
        print(f"❌ Error subiendo {local_path.name}: {e}")
        return False


def main() -> int:
    print("=" * 100)
    print("SUBIDA BACKUP MENSUAL A SHAREPOINT - FACTURAS JOYCO")
    print("=" * 100)
    print(f"Versión: {VERSION_UPLOAD}")
    print(f"Root: {ROOT}")
    print(f"Mes: {MES}")
    print(f"Carpeta local: {MES_DIR}")
    print(f"Carpeta SharePoint: {SP_BACKUP_DIR}")
    print("-" * 100)

    zip_file = buscar_ultimo_archivo(f"backup_mensual_{MES_FILE}_*.zip")
    manifest_file = buscar_ultimo_archivo(f"manifest_backup_mensual_{MES_FILE}_*.json")

    if not zip_file:
        print("❌ No se encontró ZIP mensual para subir.")
        return 1

    if not manifest_file:
        print("⚠️ No se encontró manifest mensual. Se subirá solo el ZIP.")

    try:
        ensure_folder(SP_BACKUP_DIR)
        print("✅ Carpeta SharePoint verificada/creada.")
    except Exception as e:
        print(f"❌ No se pudo verificar/crear carpeta SharePoint: {e}")
        return 1

    ok_zip = subir_archivo(zip_file, SP_BACKUP_DIR)
    ok_manifest = True

    if manifest_file:
        ok_manifest = subir_archivo(manifest_file, SP_BACKUP_DIR)

    print("-" * 100)

    if ok_zip and ok_manifest:
        print("✅ Subida mensual a SharePoint terminada correctamente.")
        print("=" * 100)
        return 0

    print("❌ Subida mensual terminó con errores.")
    print("=" * 100)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
