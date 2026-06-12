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

VERSION_UPLOAD = "2026-06-11-UPLOAD-CIERRE-DIARIO-SP-V2-DOBLE-RUTA"

DATA_DIR = ROOT / "data"
CIERRES_DIR = DATA_DIR / "cierres_diarios"

NOW = datetime.datetime.now()
FECHA = NOW.strftime("%Y-%m-%d")
MES = NOW.strftime("%Y-%m")

CIERRE_DIA_DIR = CIERRES_DIR / FECHA

SP_BACKUP_ROOT = (os.getenv("SP_BACKUP_ROOT") or f"{BASE_SP}/Backups").strip("/")
SP_BACKUP_CIERRES_DIR = (os.getenv("SP_BACKUP_CIERRES_DIR") or "01_Cierres_Diarios").strip("/")

SP_BACKUP2_HOSTNAME = (os.getenv("SP_BACKUP2_HOSTNAME") or "").strip()
SP_BACKUP2_SITE_PATH = (os.getenv("SP_BACKUP2_SITE_PATH") or "").strip()
SP_BACKUP2_DRIVE_ID = (os.getenv("SP_BACKUP2_DRIVE_ID") or "").strip()
SP_BACKUP2_FOLDER = (os.getenv("SP_BACKUP2_FOLDER") or "").strip("/")
SP_BACKUP2_CIERRES_DIR = (os.getenv("SP_BACKUP2_CIERRES_DIR") or "01_Cierres_Diarios").strip("/")

EXCLUIR_NOMBRES = {
    ".env",
}

EXCLUIR_EXT = {
    ".tmp",
    ".lock",
}


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
        carpeta=sp_join(SP_BACKUP_ROOT, SP_BACKUP_CIERRES_DIR, MES, FECHA),
        drive_id=None,
    )

    if not SP_BACKUP2_DRIVE_ID:
        raise RuntimeError("Falta SP_BACKUP2_DRIVE_ID en .env para la ruta secundaria.")

    if not SP_BACKUP2_FOLDER:
        raise RuntimeError("Falta SP_BACKUP2_FOLDER en .env para la ruta secundaria.")

    destino_secundario = DestinoSharePoint(
        nombre="SECUNDARIA_CONTROL_INTERNO",
        carpeta=sp_join(SP_BACKUP2_FOLDER, SP_BACKUP2_CIERRES_DIR, MES, FECHA),
        drive_id=SP_BACKUP2_DRIVE_ID,
    )

    return [destino_principal, destino_secundario]


def debe_subir(path: Path) -> bool:
    if not path.is_file():
        return False

    if path.name.lower() in EXCLUIR_NOMBRES:
        return False

    if path.suffix.lower() in EXCLUIR_EXT:
        return False

    return True


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
    print("SUBIDA CIERRE DIARIO A SHAREPOINT - FACTURAS JOYCO")
    print("=" * 100)
    print(f"Versión: {VERSION_UPLOAD}")
    print(f"Root: {ROOT}")
    print(f"Fecha: {FECHA}")
    print(f"Mes: {MES}")
    print(f"Carpeta local: {CIERRE_DIA_DIR}")
    print("-" * 100)

    if not CIERRE_DIA_DIR.exists():
        print("❌ No existe carpeta de cierre diario local.")
        return 1

    archivos = [p for p in sorted(CIERRE_DIA_DIR.iterdir()) if debe_subir(p)]

    if not archivos:
        print("❌ No hay archivos válidos para subir.")
        return 1

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
        print("✅ Subida de cierre diario terminada correctamente en AMBAS rutas.")
        print("=" * 100)
        return 0

    print("❌ Subida de cierre diario terminó con errores en una o más rutas.")
    print("⚠️ No borres ni archives localmente hasta revisar el error.")
    print("=" * 100)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
