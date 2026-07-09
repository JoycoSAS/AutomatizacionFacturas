# -*- coding: utf-8 -*-
"""
JOYCO - Facturas Procesador
Subida remota de cierre mensual V1 a repositorio unico de backups.

Politica:
- El cierre mensual NO se sube al SharePoint principal de operacion.
- Se sube unicamente al repositorio remoto de backups configurado por .env:
  BACKUP_DRIVE_ID / ONEDRIVE_BACKUP_DRIVE_ID / SP_BACKUP2_DRIVE_ID
  BACKUP_ROOT_FOLDER / ONEDRIVE_BACKUP_FOLDER / SP_BACKUP2_FOLDER
- La verificacion remota descarga cada archivo subido y compara:
  - Excel: datos internos hoja/celda.
  - CSV/JSON/TXT: SHA256 exacto.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import json
import os
import sys
import time
from pathlib import Path
from typing import Any, Optional
from urllib.parse import quote

import requests
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

try:
    import config  # noqa: F401
except Exception:
    pass

from services.m365.token import get_access_token

VERSION_UPLOAD = "2026-07-09-UPLOAD-CIERRE-MENSUAL-V1-REPO-BACKUPS-UNICO"
GRAPH = "https://graph.microsoft.com/v1.0"

DATA_DIR = ROOT / "data"
CIERRES_DIR = DATA_DIR / "cierres_diarios"
TMP_VERIFY_DIR = DATA_DIR / "_tmp_verificacion_cierre_mensual_v1"

EXCLUIR_NOMBRES = {".env", ".env.local", ".env.production"}
EXCLUIR_EXT = {".tmp", ".lock"}
EXCEL_EXT = {".xlsx", ".xlsm"}

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}


def ssl_verify() -> bool:
    return str(os.getenv("SSL_VERIFY", "true")).strip().lower() not in {"0", "false", "no", "off"}


def env_first(*names: str) -> str:
    for name in names:
        value = (os.getenv(name) or "").strip()
        if value:
            return value
    return ""


def encode_path(path: str) -> str:
    return "/".join(quote(part, safe="") for part in path.strip("/").split("/") if part)


def headers() -> dict[str, str]:
    token = get_access_token()
    return {"Authorization": f"Bearer {token}"}


def graph_get(url: str, **kwargs) -> requests.Response:
    return requests.get(url, headers=headers(), timeout=kwargs.pop("timeout", 120), verify=ssl_verify(), **kwargs)


def graph_put_content(drive_id: str, remote_path: str, local_file: Path) -> dict[str, Any]:
    url = f"{GRAPH}/drives/{quote(drive_id, safe='')}/root:/{encode_path(remote_path)}:/content"
    data = local_file.read_bytes()
    r = requests.put(url, headers=headers(), data=data, timeout=300, verify=ssl_verify())
    if r.status_code not in (200, 201):
        raise RuntimeError(f"PUT {r.status_code} {url} -> {r.text[:800]}")
    return r.json()


def graph_download(download_url: str, destino: Path) -> None:
    destino.parent.mkdir(parents=True, exist_ok=True)
    r = requests.get(download_url, timeout=300, verify=ssl_verify())
    if r.status_code != 200:
        raise RuntimeError(f"DOWNLOAD {r.status_code} -> {r.text[:500]}")
    destino.write_bytes(r.content)


def validar_drive(drive_id: str) -> dict[str, Any]:
    url = f"{GRAPH}/drives/{quote(drive_id, safe='')}?$select=id,name,webUrl,driveType"
    r = graph_get(url)
    if r.status_code != 200:
        raise RuntimeError(f"No se pudo validar drive {drive_id}: {r.status_code} {r.text[:500]}")
    data = r.json()
    print(f"Drive backup validado: {data.get('name')} | {data.get('id')}")
    print(f"Repositorio destino: {data.get('name')} | {data.get('webUrl')}")
    return data


def ensure_folder_recursive(drive_id: str, folder_path: str) -> None:
    parts = [p for p in folder_path.strip("/").split("/") if p]
    current = ""
    for part in parts:
        parent = current
        current = f"{current}/{part}".strip("/")
        check_url = f"{GRAPH}/drives/{quote(drive_id, safe='')}/root:/{encode_path(current)}"
        r = graph_get(check_url)
        if r.status_code == 200:
            continue
        if r.status_code != 404:
            raise RuntimeError(f"No se pudo verificar carpeta {current}: {r.status_code} {r.text[:500]}")
        if parent:
            create_url = f"{GRAPH}/drives/{quote(drive_id, safe='')}/root:/{encode_path(parent)}:/children"
        else:
            create_url = f"{GRAPH}/drives/{quote(drive_id, safe='')}/root/children"
        payload = {"name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "replace"}
        cr = requests.post(create_url, headers={**headers(), "Content-Type": "application/json"}, json=payload, timeout=120, verify=ssl_verify())
        if cr.status_code not in (200, 201):
            raise RuntimeError(f"No se pudo crear carpeta {current}: {cr.status_code} {cr.text[:500]}")


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def excel_digest(path: Path) -> dict[str, Any]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        digest = hashlib.sha256()
        hojas = []
        celdas = 0
        for ws in wb.worksheets:
            hojas.append(ws.title)
            digest.update(f"SHEET:{ws.title}\n".encode("utf-8"))
            for row in ws.iter_rows(values_only=True):
                vals = []
                for v in row:
                    if v is None:
                        vals.append("")
                    else:
                        vals.append(str(v))
                        celdas += 1
                digest.update(("\t".join(vals) + "\n").encode("utf-8", errors="replace"))
        return {"sha256_datos": digest.hexdigest(), "hojas": hojas, "celdas_no_vacias": celdas}
    finally:
        wb.close()


def verificar_archivo_subido(local: Path, item: dict[str, Any], rel_path: str) -> dict[str, Any]:
    download_url = item.get("@microsoft.graph.downloadUrl")
    if not download_url:
        raise RuntimeError(f"Graph no devolvio downloadUrl para {rel_path}")
    TMP_VERIFY_DIR.mkdir(parents=True, exist_ok=True)
    tmp = TMP_VERIFY_DIR / rel_path.replace("/", "__")
    graph_download(download_url, tmp)
    if local.suffix.lower() in EXCEL_EXT:
        local_digest = excel_digest(local)
        remote_digest = excel_digest(tmp)
        ok = local_digest["sha256_datos"] == remote_digest["sha256_datos"]
        if not ok:
            raise RuntimeError(f"Excel no coincide por datos: {rel_path}")
        print(f"Excel verificado por DATOS: {rel_path} | celdas_no_vacias={remote_digest['celdas_no_vacias']} | hojas={len(remote_digest['hojas'])}")
        return {"archivo": rel_path, "tipo": "excel_datos", "ok": True, "local": local_digest, "remoto": remote_digest}
    local_sha = sha256_file(local)
    remote_sha = sha256_file(tmp)
    if local_sha != remote_sha:
        raise RuntimeError(f"SHA256 no coincide para {rel_path}: local={local_sha}, remoto={remote_sha}")
    print(f"Archivo verificado por SHA256 exacto: {rel_path} ({local.stat().st_size} bytes)")
    return {"archivo": rel_path, "tipo": "sha256", "ok": True, "sha256": local_sha, "bytes": local.stat().st_size}


def parse_fecha(fecha: str) -> _dt.date:
    if fecha:
        return _dt.date.fromisoformat(fecha)
    return _dt.date.today()


def ultimo_dia_mes(fecha: _dt.date) -> _dt.date:
    if fecha.month == 12:
        siguiente = _dt.date(fecha.year + 1, 1, 1)
    else:
        siguiente = _dt.date(fecha.year, fecha.month + 1, 1)
    return siguiente - _dt.timedelta(days=1)


def rango_mes(fecha: _dt.date) -> tuple[_dt.date, _dt.date]:
    inicio = _dt.date(fecha.year, fecha.month, 1)
    return inicio, ultimo_dia_mes(fecha)


def rango_desde_args(args: argparse.Namespace) -> tuple[_dt.date, _dt.date]:
    if args.inicio or args.fin:
        if not args.inicio or not args.fin:
            raise RuntimeError("Si usas --inicio o --fin debes enviar ambos.")
        inicio = _dt.date.fromisoformat(args.inicio)
        fin = _dt.date.fromisoformat(args.fin)
        if fin < inicio:
            raise RuntimeError("El --fin no puede ser menor que --inicio.")
        return inicio, fin
    return rango_mes(parse_fecha(args.fecha))


def mes_carpeta(fecha: _dt.date) -> str:
    return f"{fecha:%Y-%m}_{MESES_ES.get(fecha.month, fecha.strftime('%B'))}"


def periodo_nombre(inicio: _dt.date, fin: _dt.date) -> str:
    if inicio.day == 1 and fin == ultimo_dia_mes(inicio):
        return inicio.strftime("%Y-%m")
    return f"{inicio.isoformat()}_a_{fin.isoformat()}"


def buscar_cierre_mensual_local(inicio: _dt.date, fin: _dt.date) -> Path:
    yyyy = inicio.strftime("%Y")
    mes = mes_carpeta(inicio)
    esperado = CIERRES_DIR / yyyy / mes / "Mensual"
    if esperado.exists():
        return esperado
    candidatos = list(CIERRES_DIR.glob(f"**/{mes}/Mensual"))
    candidatos = [p for p in candidatos if p.exists() and p.is_dir()]
    if not candidatos:
        raise RuntimeError(f"No se encontro cierre mensual local para {periodo_nombre(inicio, fin)} dentro de {CIERRES_DIR}.")
    candidatos = sorted(candidatos, key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0]


def debe_subir(path: Path) -> bool:
    if not path.is_file():
        return False
    if path.name in EXCLUIR_NOMBRES:
        return False
    if path.suffix.lower() in EXCLUIR_EXT:
        return False
    if any(part.startswith("_tmp") for part in path.parts):
        return False
    return True


def listar_archivos(base: Path) -> list[Path]:
    return [p for p in sorted(base.rglob("*")) if debe_subir(p)]


def validar_cierre_local_minimo(cierre_dir: Path, inicio: _dt.date, fin: _dt.date) -> None:
    periodo = periodo_nombre(inicio, fin)
    obligatorios = [
        cierre_dir / "01_Excel_Mensual" / f"facturas_mensual_{periodo}.xlsx",
        cierre_dir / "04_Manifest_Mensual" / f"manifest_mensual_{periodo}.json",
        cierre_dir / "04_Manifest_Mensual" / f"resumen_mensual_{periodo}.txt",
        cierre_dir / "05_Validaciones" / f"validacion_local_mensual_{periodo}.json",
    ]
    faltantes = [str(p) for p in obligatorios if not p.exists()]
    if faltantes:
        raise RuntimeError("El cierre mensual local no esta completo. Faltan: " + "; ".join(faltantes))
    validacion_path = cierre_dir / "05_Validaciones" / f"validacion_local_mensual_{periodo}.json"
    try:
        validacion = json.loads(validacion_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise RuntimeError(f"No se pudo leer validacion local: {validacion_path} | {exc}")
    if not validacion.get("ok"):
        raise RuntimeError(f"La validacion local del cierre mensual no esta OK: {validacion_path}")


def remote_base_para_cierre(cierre_dir: Path) -> str:
    root_folder = env_first("BACKUP_ROOT_FOLDER", "ONEDRIVE_BACKUP_FOLDER", "SP_BACKUP2_FOLDER")
    if not root_folder:
        raise RuntimeError("Falta BACKUP_ROOT_FOLDER/ONEDRIVE_BACKUP_FOLDER/SP_BACKUP2_FOLDER en .env")
    try:
        rel_cierre = cierre_dir.resolve().relative_to(CIERRES_DIR.resolve()).as_posix()
    except Exception:
        rel_cierre = cierre_dir.name
    return f"{root_folder}/{rel_cierre}".strip("/")


def escribir_validacion_remota(cierre_dir: Path, inicio: _dt.date, fin: _dt.date, drive: dict[str, Any], remote_base: str, resultados: list[dict[str, Any]]) -> Path:
    validaciones_dir = cierre_dir / "05_Validaciones"
    validaciones_dir.mkdir(parents=True, exist_ok=True)
    periodo = periodo_nombre(inicio, fin)
    path = validaciones_dir / f"validacion_remota_mensual_{periodo}.json"
    payload = {
        "tipo": "validacion_remota_cierre_mensual",
        "version": VERSION_UPLOAD,
        "periodo": periodo,
        "fecha_inicio": inicio.isoformat(),
        "fecha_fin": fin.isoformat(),
        "generado_en": _dt.datetime.now().isoformat(timespec="seconds"),
        "cierre_local": str(cierre_dir),
        "drive": {"id": drive.get("id"), "name": drive.get("name"), "webUrl": drive.get("webUrl"), "driveType": drive.get("driveType")},
        "remote_base": remote_base,
        "total_archivos_verificados": len(resultados),
        "ok": all(x.get("ok") for x in resultados),
        "resultados": resultados,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def main() -> int:
    parser = argparse.ArgumentParser(description="Sube cierre mensual V1 al repositorio remoto unico de backups.")
    parser.add_argument("--fecha", default="", help="Fecha dentro del mes, formato YYYY-MM-DD. Default: hoy.")
    parser.add_argument("--inicio", default="", help="Inicio de rango, formato YYYY-MM-DD. Opcional.")
    parser.add_argument("--fin", default="", help="Fin de rango, formato YYYY-MM-DD. Opcional.")
    parser.add_argument("--dry-run", action="store_true", help="Muestra archivos/ruta sin subir.")
    args = parser.parse_args()

    inicio, fin = rango_desde_args(args)
    periodo = periodo_nombre(inicio, fin)

    print("=" * 100)
    print("SUBIDA CIERRE MENSUAL V1 A REPOSITORIO DE BACKUPS - FACTURAS JOYCO")
    print("=" * 100)
    print(f"Version: {VERSION_UPLOAD}")
    print(f"Root: {ROOT}")
    print(f"Periodo: {periodo}")
    print(f"Rango: {inicio.isoformat()} a {fin.isoformat()}")
    print(f"Dry-run: {args.dry_run}")
    print("-" * 100)

    try:
        cierre_dir = buscar_cierre_mensual_local(inicio, fin)
        validar_cierre_local_minimo(cierre_dir, inicio, fin)
    except Exception as exc:
        print(f"Cierre mensual local no valido: {exc}")
        return 1

    try:
        remote_base = remote_base_para_cierre(cierre_dir)
    except Exception as exc:
        print(f"No se pudo calcular ruta remota: {exc}")
        return 1

    archivos = listar_archivos(cierre_dir)
    print(f"Cierre local: {cierre_dir}")
    print(f"Ruta remota base: {remote_base}")
    print(f"Archivos detectados: {len(archivos)}")
    for p in archivos:
        print(f"   - {p.relative_to(cierre_dir).as_posix()} ({p.stat().st_size} bytes)")

    if args.dry_run:
        print("-" * 100)
        print("DRY-RUN: no se llamo a Microsoft Graph ni se subio nada.")
        print("Si la ruta remota base es correcta, ejecuta sin --dry-run.")
        print("=" * 100)
        return 0

    drive_id = env_first("BACKUP_DRIVE_ID", "ONEDRIVE_BACKUP_DRIVE_ID", "SP_BACKUP2_DRIVE_ID")
    if not drive_id:
        print("Falta BACKUP_DRIVE_ID/ONEDRIVE_BACKUP_DRIVE_ID/SP_BACKUP2_DRIVE_ID en .env")
        return 1

    try:
        drive = validar_drive(drive_id)
        ensure_folder_recursive(drive_id, remote_base)
        print("Carpeta remota base verificada/creada.")
    except Exception as exc:
        print(f"Error preparando repositorio remoto de backups: {exc}")
        return 1

    resultados: list[dict[str, Any]] = []
    ok_todos = True
    try:
        if TMP_VERIFY_DIR.exists():
            shutil.rmtree(TMP_VERIFY_DIR, ignore_errors=True)
        TMP_VERIFY_DIR.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass

    for local in archivos:
        rel_path = local.relative_to(cierre_dir).as_posix()
        remote_path = f"{remote_base}/{rel_path}".strip("/")
        try:
            print("Subiendo:")
            print(f"   Local:  {local}")
            print(f"   Remoto: {remote_path}")
            item = graph_put_content(drive_id, remote_path, local)
            ultimo_error = None
            for intento in range(1, 5):
                try:
                    resultado = verificar_archivo_subido(local, item, rel_path)
                    resultados.append(resultado)
                    ultimo_error = None
                    break
                except Exception as exc:
                    ultimo_error = exc
                    print(f"Verificacion intento {intento}/4 fallo para {rel_path}: {exc}")
                    if intento < 4:
                        time.sleep(2 * intento)
            if ultimo_error is not None:
                raise ultimo_error
        except Exception as exc:
            ok_todos = False
            print(f"ERROR subiendo/verificando {rel_path}: {type(exc).__name__}: {exc}")

    validacion_remota = escribir_validacion_remota(cierre_dir, inicio, fin, drive, remote_base, resultados)
    rel_val = validacion_remota.relative_to(cierre_dir).as_posix()
    remote_val = f"{remote_base}/{rel_val}".strip("/")
    try:
        print("Subiendo validacion remota final:")
        print(f"   Local:  {validacion_remota}")
        print(f"   Remoto: {remote_val}")
        item = graph_put_content(drive_id, remote_val, validacion_remota)
        verificar_archivo_subido(validacion_remota, item, rel_val)
    except Exception as exc:
        ok_todos = False
        print(f"ERROR subiendo/verificando validacion remota final: {type(exc).__name__}: {exc}")

    print("-" * 100)
    print(f"Validacion remota local: {validacion_remota}")
    if ok_todos:
        print("Subida de cierre mensual V1 terminada correctamente en el repositorio unico de backups.")
        print("Verificacion aplicada archivo por archivo despues de descargar desde Graph.")
        print("=" * 100)
        return 0
    print("Subida de cierre mensual V1 termino con errores.")
    print("No borres ni archives localmente hasta revisar el error.")
    print("=" * 100)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
