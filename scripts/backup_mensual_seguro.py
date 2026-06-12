import os
import sys
import time
import datetime
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional, Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

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

SP_BACKUP_ROOT = (os.getenv("SP_BACKUP_ROOT") or f"{BASE_SP}/Backups").strip("/")
SP_BACKUP_MENSUALES_DIR = (os.getenv("SP_BACKUP_MENSUALES_DIR") or "02_Backups_Mensuales").strip("/")

SP_BACKUP2_HOSTNAME = (os.getenv("SP_BACKUP2_HOSTNAME") or "").strip()
SP_BACKUP2_SITE_PATH = (os.getenv("SP_BACKUP2_SITE_PATH") or "").strip()
SP_BACKUP2_DRIVE_ID = (os.getenv("SP_BACKUP2_DRIVE_ID") or "").strip()
SP_BACKUP2_FOLDER = (os.getenv("SP_BACKUP2_FOLDER") or "").strip("/")
SP_BACKUP2_MENSUALES_DIR = (os.getenv("SP_BACKUP2_MENSUALES_DIR") or "02_Backups_Mensuales").strip("/")


@dataclass(frozen=True)
class DestinoSharePoint:
    nombre: str
    carpeta: str
    drive_id: Optional[str] = None


def sp_join(*parts: str) -> str:
    return "/".join(str(p).strip("/") for p in parts if str(p).strip("/"))


def reintentar(descripcion: str, funcion: Callable[[], Any], intentos: int = 3, espera_base: int = 3) -> Any:
    ultimo_error = None

    for intento in range(1, intentos + 1):
        try:
            if intento > 1:
                print(f"🔁 Reintento {intento}/{intentos}: {descripcion}")
            return funcion()
        except Exception as e:
            ultimo_error = e
            print(f"⚠️ Error en {descripcion} intento {intento}/{intentos}: {e}")

            if intento < intentos:
                espera = espera_base * intento
                print(f"⏳ Esperando {espera}s antes de reintentar...")
                time.sleep(espera)

    raise ultimo_error


def construir_destinos() -> list[DestinoSharePoint]:
    destino_principal = DestinoSharePoint(
        nombre="PRINCIPAL_CONTABILIDAD",
        carpeta=sp_join(SP_BACKUP_ROOT, SP_BACKUP_MENSUALES_DIR, MES),
        drive_id=None,
    )

    if not SP_BACKUP2_DRIVE_ID:
        raise RuntimeError("Falta SP_BACKUP2_DRIVE_ID en .env para la ruta secundaria.")

    if not SP_BACKUP2_FOLDER:
        raise RuntimeError("Falta SP_BACKUP2_FOLDER en .env para la ruta secundaria.")

    destino_secundario = DestinoSharePoint(
        nombre="SECUNDARIA_CONTROL_INTERNO",
        carpeta=sp_join(SP_BACKUP2_FOLDER, SP_BACKUP2_MENSUALES_DIR, MES),
        drive_id=SP_BACKUP2_DRIVE_ID,
    )

    return [destino_principal, destino_secundario]


def buscar_ultimo_archivo(pattern: str) -> Path | None:
    if not MES_DIR.exists():
        return None

    archivos = sorted(
        MES_DIR.glob(pattern),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    return archivos[0] if archivos else None


def verificar_archivo(local_path: Path, sp_path: str, drive_id: Optional[str]) -> bool:
    try:
        item = reintentar(
            descripcion=f"verificar archivo {local_path.name}",
            funcion=lambda: get_item_by_path(sp_path, drive_id=drive_id),
            intentos=3,
        )

        size_sp = int(item.get("size", -1))
        size_local = int(local_path.stat().st_size)

        print("🔎 Verificación SharePoint:")
        print(f"   Archivo:      {local_path.name}")
        print(f"   Tamaño local: {size_local}")
        print(f"   Tamaño SP:    {size_sp}")
        print(f"   Modificado:   {item.get('lastModifiedDateTime')}")

        if size_sp != size_local:
            print("❌ Verificación fallida: el tamaño no coincide.")
            return False

        print("✅ Verificación correcta: archivo existe y tamaño coincide.")
        return True

    except Exception as e:
        print(f"❌ Error verificando {local_path.name}: {e}")
        return False


def subir_archivo_a_destino(local_path: Path, destino: DestinoSharePoint) -> bool:
    if not local_path or not local_path.exists():
        print(f"⚠️ No existe archivo local: {local_path}")
        return False

    sp_path = sp_join(destino.carpeta, local_path.name)

    print("☁️ Subiendo a SharePoint:")
    print(f"   Destino: {destino.nombre}")
    print(f"   Local:   {local_path}")
    print(f"   SP:      {sp_path}")

    try:
        reintentar(
            descripcion=f"subir {local_path.name} a {destino.nombre}",
            funcion=lambda: upload_small_file(
                str(local_path),
                sp_path,
                mode="replace",
                drive_id=destino.drive_id,
            ),
            intentos=3,
        )
        print(f"✅ Subido correctamente en {destino.nombre}: {local_path.name}")
    except Exception as e:
        print(f"❌ Error subiendo {local_path.name} en {destino.nombre}: {e}")
        return False

    return verificar_archivo(local_path, sp_path, destino.drive_id)


def preparar_destino(destino: DestinoSharePoint) -> bool:
    print("-" * 100)
    print(f"📁 Verificando/creando carpeta destino: {destino.nombre}")
    print(f"   SP: {destino.carpeta}")

    try:
        reintentar(
            descripcion=f"crear/verificar carpeta {destino.nombre}",
            funcion=lambda: ensure_folder(destino.carpeta, drive_id=destino.drive_id),
            intentos=3,
        )
        print(f"✅ Carpeta SharePoint verificada/creada: {destino.nombre}")
        return True
    except Exception as e:
        print(f"❌ No se pudo verificar/crear carpeta {destino.nombre}: {e}")
        return False


def main() -> int:
    print("=" * 100)
    print("SUBIDA BACKUP MENSUAL A SHAREPOINT - FACTURAS JOYCO")
    print("=" * 100)
    print(f"Versión: {VERSION_UPLOAD}")
    print(f"Root: {ROOT}")
    print(f"Mes: {MES}")
    print(f"Carpeta local: {MES_DIR}")
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

    try:
        destinos = construir_destinos()
    except Exception as e:
        print(f"❌ Configuración incompleta de destinos SharePoint: {e}")
        return 1

    print("📌 Destinos configurados:")
    for destino in destinos:
        print(f"   - {destino.nombre}: {destino.carpeta}")

    print("📄 Archivos a subir:")
    for archivo in archivos:
        print(f"   - {archivo.name} ({archivo.stat().st_size} bytes)")

    ok_total = True

    for destino in destinos:
        if not preparar_destino(destino):
            ok_total = False
            continue

        for archivo in archivos:
            ok = subir_archivo_a_destino(archivo, destino)
            if not ok:
                ok_total = False

    print("-" * 100)

    if ok_total:
        print("✅ Subida mensual terminada correctamente en AMBAS rutas.")
        print("=" * 100)
        return 0

    print("❌ Subida mensual terminó con errores en una o más rutas.")
    print("⚠️ No borres ni archives localmente hasta revisar el error.")
    print("=" * 100)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
