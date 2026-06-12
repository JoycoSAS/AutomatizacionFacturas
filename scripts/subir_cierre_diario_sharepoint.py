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

VERSION_UPLOAD = "2026-06-11-UPLOAD-CIERRE-DIARIO-SP-V2-DOBLE-RUTA"

DATA_DIR = ROOT / "data"
CIERRES_DIR = DATA_DIR / "cierres_diarios"

NOW = datetime.datetime.now()
FECHA = NOW.strftime("%Y-%m-%d")
MES = NOW.strftime("%Y-%m")

CIERRE_DIA_DIR = CIERRES_DIR / FECHA

SP_CIERRE_DIR_PRINCIPAL = f"{BASE_SP}/Backups/01_Cierres_Diarios/{MES}/{FECHA}".strip("/")

SP_BACKUP2_DRIVE_ID = (os.getenv("SP_BACKUP2_DRIVE_ID") or "").strip()
SP_BACKUP2_FOLDER = (os.getenv("SP_BACKUP2_FOLDER") or "").strip().strip("/")

SP_CIERRE_DIR_SECUNDARIA = (
    f"{SP_BACKUP2_FOLDER}/01_Cierres_Diarios/{MES}/{FECHA}".strip("/")
    if SP_BACKUP2_FOLDER
    else ""
)

EXCLUIR_NOMBRES = {
    ".env",
}

EXCLUIR_EXT = {
    ".tmp",
    ".lock",
}


def debe_subir(path: Path) -> bool:
    if not path.is_file():
        return False

    if path.name.lower() in EXCLUIR_NOMBRES:
        return False

    if path.suffix.lower() in EXCLUIR_EXT:
        return False

    return True


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
            "sp_dir": SP_CIERRE_DIR_PRINCIPAL,
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
            "sp_dir": SP_CIERRE_DIR_SECUNDARIA,
            "drive_id": SP_BACKUP2_DRIVE_ID,
        }
    )

    return destinos


def main() -> int:
    print("=" * 100)
    print("SUBIDA CIERRE DIARIO A SHAREPOINT - FACTURAS JOYCO")
    print("=" * 100)
    print(f"Versión: {VERSION_UPLOAD}")
    print(f"Root: {ROOT}")
    print(f"Fecha: {FECHA}")
    print(f"Carpeta local: {CIERRE_DIA_DIR}")
    print(f"Ruta principal:   {SP_CIERRE_DIR_PRINCIPAL}")
    print(f"Ruta secundaria:  {SP_CIERRE_DIR_SECUNDARIA}")
    print("-" * 100)

    if not CIERRE_DIA_DIR.exists():
        print("❌ No existe carpeta de cierre diario local.")
        return 1

    archivos = [p for p in sorted(CIERRE_DIA_DIR.iterdir()) if debe_subir(p)]

    if not archivos:
        print("❌ No hay archivos válidos para subir.")
        return 1

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
        print("✅ Subida de cierre diario a SharePoint terminada correctamente en ambas rutas.")
        print("=" * 100)
        return 0

    print("❌ Subida de cierre diario terminó con errores en una o más rutas.")
    print("=" * 100)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
