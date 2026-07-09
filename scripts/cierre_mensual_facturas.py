# -*- coding: utf-8 -*-
"""
JOYCO - Facturas Procesador
Cierre mensual local seguro V1

Politica:
- El cierre mensual NO se construye desde cierres semanales ni diarios.
- Genera un Excel mensual cruzando el Excel operativo principal contra las auditorias reales del mes.
- Los cierres semanales, si existen, se copian solo como soporte/evidencia, no como fuente de calculo.
- No sube a OneDrive/SharePoint. La subida se hace con scripts/subir_cierre_mensual_sharepoint.py.
- No incluye .env real ni secretos.

Estructura generada:
data/cierres_diarios/YYYY/YYYY-MM_Mes/Mensual/
    01_Excel_Mensual/facturas_mensual_YYYY-MM.xlsx
    02_Auditorias_Mes/audit_*.csv
    03_Soporte_Semanas/SEMANA_YYYY-MM-DD_a_YYYY-MM-DD/...
    04_Manifest_Mensual/manifest_mensual_YYYY-MM.json
    04_Manifest_Mensual/resumen_mensual_YYYY-MM.txt
    05_Validaciones/validacion_local_mensual_YYYY-MM.json
    06_Logs_Mes/...
"""

from __future__ import annotations

import argparse
import csv
import datetime as _dt
import hashlib
import json
import os
import platform
import re
import shutil
import socket
import sys
from pathlib import Path
from typing import Any, Optional, Sequence

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.worksheet.table import Table, TableStyleInfo
except Exception as exc:  # pragma: no cover
    raise RuntimeError("Este script requiere openpyxl instalado.") from exc

ROOT = Path(__file__).resolve().parents[1]
try:
    sys.path.insert(0, str(ROOT))
    import config  # noqa: F401
except Exception:
    pass

VERSION_CIERRE = "2026-07-09-CIERRE-MENSUAL-V1-FUENTE-EXCEL-AUDITORIAS"

DATA_DIR = ROOT / "data"
AUDIT_DIR = DATA_DIR / "audit"
LOGS_DIR = ROOT / "logs"
DATA_LOGS_DIR = DATA_DIR / "logs"
CIERRES_DIR = DATA_DIR / "cierres_diarios"
LOCKS_DIR = DATA_DIR / "locks"

FACTURAS_PATH = Path(os.getenv("ARCHIVO_EXCEL_LOCAL", str(DATA_DIR / "facturas.xlsx")))
if not FACTURAS_PATH.is_absolute():
    FACTURAS_PATH = ROOT / FACTURAS_PATH

LOCK_FILE = LOCKS_DIR / "cierre_mensual_facturas.lock"
LOCK_TTL_SECONDS = int(os.getenv("LOCK_TTL_SECONDS", "3600") or "3600")

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

PALABRAS_SENSIBLES = (
    "SECRET", "PASSWORD", "PASS", "TOKEN", "KEY", "CLIENT_SECRET", "TENANT_ID", "CLIENT_ID",
    "CONTRASENA", "CONTRASEÑA", "PWD", "AUTH", "PRIVATE",
)
EXT_LOGS = {".log", ".txt", ".csv", ".json"}
EXCEL_SHEET_NAME = "Facturas"
EXCEL_TABLE_NAME = "TblFacturasMensual"


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Genera cierre mensual local seguro V1.")
    parser.add_argument("--fecha", default=_dt.date.today().isoformat(), help="Fecha dentro del mes, formato YYYY-MM-DD.")
    parser.add_argument("--inicio", default="", help="Inicio de rango manual, formato YYYY-MM-DD.")
    parser.add_argument("--fin", default="", help="Fin de rango manual, formato YYYY-MM-DD.")
    parser.add_argument("--dry-run", action="store_true", help="Calcula y muestra resumen sin generar archivos.")
    parser.add_argument("--permitir-vacio", action="store_true", help="Permite generar cierre aunque no haya filas mensuales.")
    return parser.parse_args()


def validar_fecha(fecha: str) -> _dt.date:
    try:
        return _dt.date.fromisoformat(fecha)
    except Exception as exc:
        raise RuntimeError(f"Fecha invalida {fecha!r}. Usa formato YYYY-MM-DD.") from exc


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
        inicio = validar_fecha(args.inicio)
        fin = validar_fecha(args.fin)
        if fin < inicio:
            raise RuntimeError("El --fin no puede ser menor que --inicio.")
        return inicio, fin
    return rango_mes(validar_fecha(args.fecha))


def fechas_en_rango(inicio: _dt.date, fin: _dt.date) -> list[_dt.date]:
    out: list[_dt.date] = []
    d = inicio
    while d <= fin:
        out.append(d)
        d += _dt.timedelta(days=1)
    return out


def mes_carpeta(fecha: _dt.date) -> str:
    return f"{fecha:%Y-%m}_{MESES_ES.get(fecha.month, fecha.strftime('%B'))}"


def periodo_nombre(inicio: _dt.date, fin: _dt.date) -> str:
    if inicio.day == 1 and fin == ultimo_dia_mes(inicio):
        return inicio.strftime("%Y-%m")
    return f"{inicio.isoformat()}_a_{fin.isoformat()}"


def semana_lunes_domingo(fecha: _dt.date) -> tuple[_dt.date, _dt.date]:
    inicio = fecha - _dt.timedelta(days=fecha.weekday())
    fin = inicio + _dt.timedelta(days=6)
    return inicio, fin


def semana_nombre(inicio: _dt.date, fin: _dt.date) -> str:
    return f"SEMANA_{inicio.isoformat()}_a_{fin.isoformat()}"


def semanas_en_rango(inicio: _dt.date, fin: _dt.date) -> list[tuple[_dt.date, _dt.date]]:
    semanas: list[tuple[_dt.date, _dt.date]] = []
    d = inicio
    vistos = set()
    while d <= fin:
        si, sf = semana_lunes_domingo(d)
        key = (si, sf)
        if key not in vistos:
            vistos.add(key)
            semanas.append(key)
        d += _dt.timedelta(days=1)
    return semanas


def rel(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(ROOT.resolve())).replace("\\", "/")
    except Exception:
        return str(path).replace("\\", "/")


def normalizar_texto(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    s = re.sub(r"\s+", " ", s)
    return s.upper()


def token_valido(s: str) -> bool:
    s = normalizar_texto(s)
    if not s:
        return False
    if len(s) < 4:
        return False
    if s in {"TRUE", "FALSE", "OK", "SI", "NO", "N/A", "NA", "NONE", "NULL"}:
        return False
    if re.fullmatch(r"\d{1,3}", s):
        return False
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return False
    if not any(ch.isdigit() for ch in s) and len(s) < 12:
        return False
    return True


def crear_lock() -> tuple[bool, str]:
    LOCKS_DIR.mkdir(parents=True, exist_ok=True)
    if LOCK_FILE.exists():
        try:
            stat = LOCK_FILE.stat()
            age = _dt.datetime.now().timestamp() - stat.st_mtime
            if age < LOCK_TTL_SECONDS:
                return False, f"Existe lock activo: {LOCK_FILE} | edad={age:.0f}s"
            LOCK_FILE.unlink(missing_ok=True)
        except Exception as exc:
            return False, f"No se pudo eliminar lock vencido: {LOCK_FILE} | {exc}"
    payload = {
        "script": "cierre_mensual_facturas.py",
        "version": VERSION_CIERRE,
        "pid": os.getpid(),
        "started_at": _dt.datetime.now().isoformat(timespec="seconds"),
        "host": socket.gethostname(),
    }
    LOCK_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return True, "lock creado"


def liberar_lock() -> None:
    try:
        LOCK_FILE.unlink(missing_ok=True)
    except Exception:
        pass


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def info_archivo(path: Path, origen: Optional[Path] = None, categoria: str = "evidencia") -> dict[str, Any]:
    return {
        "categoria": categoria,
        "archivo": rel(path),
        "origen": rel(origen) if origen else None,
        "nombre": path.name,
        "bytes": path.stat().st_size if path.exists() else 0,
        "sha256": sha256_file(path) if path.exists() and path.is_file() else None,
    }


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with path.open("r", encoding=enc, newline="") as f:
                sample = f.read(4096)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                except Exception:
                    dialect = csv.excel
                reader = csv.DictReader(f, dialect=dialect)
                return [{k or "": v or "" for k, v in row.items()} for row in reader]
        except UnicodeDecodeError:
            continue
        except Exception:
            return []
    return []


def recolectar_audits(inicio: _dt.date, fin: _dt.date) -> list[Path]:
    rutas: list[Path] = []
    fechas = [d.isoformat() for d in fechas_en_rango(inicio, fin)]
    for fecha in fechas:
        patrones = [
            f"audit_detalle_{fecha}.csv",
            f"audit_runs_{fecha}.csv",
            f"audit_*_{fecha}.csv",
            f"*{fecha}*.csv",
        ]
        for base in (AUDIT_DIR, DATA_DIR, ROOT):
            if not base.exists():
                continue
            for patron in patrones:
                for p in base.glob(patron):
                    if p.is_file() and "audit" in p.name.lower():
                        rutas.append(p)
    vistos = set()
    out = []
    for p in rutas:
        try:
            key = str(p.resolve()).lower()
        except Exception:
            key = str(p).lower()
        if key not in vistos:
            vistos.add(key)
            out.append(p)
    return sorted(out)


def log_corresponde_rango(path: Path, inicio: _dt.date, fin: _dt.date) -> bool:
    name = path.name.lower()
    for d in fechas_en_rango(inicio, fin):
        fecha = d.isoformat()
        if fecha in name or fecha.replace("-", "") in name:
            return True
    try:
        mtime = _dt.date.fromtimestamp(path.stat().st_mtime)
        return inicio <= mtime <= fin
    except Exception:
        return False


def recolectar_logs(inicio: _dt.date, fin: _dt.date) -> list[Path]:
    rutas: list[Path] = []
    for base in (LOGS_DIR, DATA_LOGS_DIR):
        if not base.exists():
            continue
        for p in base.rglob("*"):
            if not p.is_file():
                continue
            if p.suffix.lower() not in EXT_LOGS:
                continue
            if log_corresponde_rango(p, inicio, fin):
                rutas.append(p)
    vistos = set()
    out = []
    for p in rutas:
        try:
            key = str(p.resolve()).lower()
        except Exception:
            key = str(p).lower()
        if key not in vistos:
            vistos.add(key)
            out.append(p)
    return sorted(out)


def extraer_candidatos_auditoria(audit_files: list[Path]) -> set[str]:
    candidatos: set[str] = set()
    columnas_clave = (
        "cufe", "cude", "uuid", "numero", "factura", "prefijo", "radicado",
        "archivo", "file", "nombre", "invoice", "documento",
    )
    for path in audit_files:
        rows = read_csv_rows(path)
        for row in rows:
            for k, v in row.items():
                nk = normalizar_texto(k)
                if not any(c.upper() in nk for c in columnas_clave):
                    continue
                nv = normalizar_texto(v)
                if token_valido(nv):
                    candidatos.add(nv)
    return candidatos


def fila_match_candidatos(row: Sequence[Any], candidatos: set[str]) -> bool:
    if not candidatos:
        return False
    valores = {normalizar_texto(v) for v in row if token_valido(normalizar_texto(v))}
    if valores.intersection(candidatos):
        return True
    # Match parcial conservador para nombres de archivo/radicados embebidos.
    for v in valores:
        if len(v) < 6:
            continue
        for c in candidatos:
            if len(c) >= 8 and (c in v or v in c):
                return True
    return False


def obtener_filas_mensuales_desde_excel(
    inicio: _dt.date,
    fin: _dt.date,
    audit_files: list[Path],
) -> tuple[list[Any], list[list[Any]], dict[str, Any]]:
    meta: dict[str, Any] = {
        "metodo_seleccion": None,
        "archivo_excel": rel(FACTURAS_PATH),
        "auditorias_usadas": [rel(p) for p in audit_files],
        "candidatos": 0,
        "advertencias": [],
    }
    if not FACTURAS_PATH.exists():
        raise RuntimeError(f"No existe Excel operativo: {FACTURAS_PATH}")

    candidatos = extraer_candidatos_auditoria(audit_files)
    meta["candidatos"] = len(candidatos)

    print(f"Leyendo Excel operativo: {FACTURAS_PATH}")
    wb = load_workbook(FACTURAS_PATH, read_only=True, data_only=True)
    try:
        ws = wb[EXCEL_SHEET_NAME] if EXCEL_SHEET_NAME in wb.sheetnames else wb[wb.sheetnames[0]]
        print(f"Hoja usada: {ws.title} | max_row={ws.max_row} | max_column={ws.max_column}")
        rows_iter = ws.iter_rows(values_only=True)
        headers = list(next(rows_iter, []))
        rows_match: list[list[Any]] = []
        total = 0
        for row_tuple in rows_iter:
            total += 1
            row = list(row_tuple)
            if not any(v is not None and str(v).strip() for v in row):
                continue
            if fila_match_candidatos(row, candidatos):
                rows_match.append(row)
            if total % 1000 == 0:
                print(f"   ... cruce auditoria mensual: {total} | matches={len(rows_match)}")
        print(f"Lectura Excel terminada: filas leidas={total} | matches={len(rows_match)}")
        if not candidatos:
            meta["metodo_seleccion"] = "sin_candidatos_auditoria_mes"
            meta["advertencias"].append("No se encontraron candidatos suficientes en auditorias del mes.")
        else:
            meta["metodo_seleccion"] = "auditorias_mes_vs_excel_operativo"
        return headers, rows_match, meta
    finally:
        wb.close()


def ajustar_ancho_columnas(ws) -> None:
    for col_cells in ws.columns:
        max_len = 0
        col_letter = col_cells[0].column_letter
        for cell in col_cells[:200]:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 45)


def crear_excel_mensual(destino: Path, headers: list[Any], rows: list[list[Any]]) -> dict[str, Any]:
    destino.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = EXCEL_SHEET_NAME
    clean_headers = [str(h) if h is not None and str(h).strip() else f"Columna_{i+1}" for i, h in enumerate(headers)]
    ws.append(clean_headers)
    for row in rows:
        ws.append(row[:len(clean_headers)] + [None] * max(0, len(clean_headers) - len(row)))
    if clean_headers:
        end_col = ws.cell(row=1, column=len(clean_headers)).column_letter
        end_row = max(1, len(rows) + 1)
        table = Table(displayName=EXCEL_TABLE_NAME, ref=f"A1:{end_col}{end_row}")
        style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        table.tableStyleInfo = style
        ws.add_table(table)
    ajustar_ancho_columnas(ws)
    wb.save(destino)
    wb.close()
    return {
        "archivo": rel(destino),
        "filas_datos": len(rows),
        "columnas": len(clean_headers),
        "bytes": destino.stat().st_size,
        "sha256": sha256_file(destino),
    }


def copiar_archivo(origen: Path, destino: Path, categoria: str) -> Optional[dict[str, Any]]:
    if not origen.exists() or not origen.is_file():
        return None
    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(origen, destino)
    return info_archivo(destino, origen=origen, categoria=categoria)


def copiar_auditorias(audit_files: list[Path], destino_dir: Path) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    for p in audit_files:
        info = copiar_archivo(p, destino_dir / p.name, "auditoria_mes")
        if info:
            items.append(info)
    return items


def copiar_logs(log_files: list[Path], destino_dir: Path) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    for p in log_files:
        info = copiar_archivo(p, destino_dir / p.name, "log_mes")
        if info:
            items.append(info)
    return items


def buscar_cierres_semanales_soporte(inicio: _dt.date, fin: _dt.date, mes_dir: Path) -> list[Path]:
    out: list[Path] = []
    for si, sf in semanas_en_rango(inicio, fin):
        semana_dir = mes_dir / semana_nombre(si, sf)
        semanal_v2 = semana_dir / "Semanal"
        if semanal_v2.exists() and semanal_v2.is_dir():
            out.append(semanal_v2)
            continue
        rango = f"{si.isoformat()}_a_{sf.isoformat()}"
        semanal_v1 = semana_dir / "02_Cierre_Semanal" / f"cierre_semanal_{rango}"
        if semanal_v1.exists() and semanal_v1.is_dir():
            out.append(semanal_v1)
    return out


def copiar_soporte_semanas(cierres_semanales: list[Path], destino_dir: Path) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    destino_dir.mkdir(parents=True, exist_ok=True)
    for cierre in cierres_semanales:
        semana_folder = cierre.parent.name if cierre.name == "Semanal" else cierre.name.replace("cierre_semanal_", "SEMANA_")
        destino_base = destino_dir / semana_folder
        destino_base.mkdir(parents=True, exist_ok=True)
        candidatos = []
        candidatos.extend((cierre / "04_Manifest_Semanal").glob("manifest_semanal_*.json"))
        candidatos.extend((cierre / "04_Manifest_Semanal").glob("resumen_semanal_*.txt"))
        candidatos.extend((cierre / "05_Validaciones").glob("validacion_local_semanal_*.json"))
        candidatos.extend((cierre / "05_Validaciones").glob("validacion_remota_semanal_*.json"))
        for origen in candidatos:
            try:
                info = copiar_archivo(origen, destino_base / origen.name, "soporte_semana")
                if info:
                    items.append(info)
            except Exception as exc:
                print(f"ADVERTENCIA: no se pudo copiar soporte semanal {origen}: {type(exc).__name__}: {exc}")
    return items


def redactar_env_line(line: str) -> str:
    if not line or line.strip().startswith("#") or "=" not in line:
        return line
    key, value = line.split("=", 1)
    key_up = key.upper()
    if any(p in key_up for p in PALABRAS_SENSIBLES):
        if value.strip():
            return f"{key}=***REDACTADO***"
    return line


def crear_snapshot_env_redactado(destino_dir: Path, periodo: str) -> Optional[dict[str, Any]]:
    env_path = ROOT / ".env"
    if not env_path.exists():
        return None
    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / f"snapshot_env_redactado_mensual_{periodo}.txt"
    lineas = env_path.read_text(encoding="utf-8", errors="replace").splitlines()
    destino.write_text("\n".join(redactar_env_line(x) for x in lineas) + "\n", encoding="utf-8")
    return info_archivo(destino, origen=env_path, categoria="config_redactada")


def validar_excel_generado(path: Path, expected_columns: int) -> dict[str, Any]:
    out: dict[str, Any] = {"archivo": rel(path), "existe": path.exists(), "abre_ok": False, "hojas": [], "filas_datos": 0, "columnas": 0, "errores": []}
    if not path.exists():
        out["errores"].append("No existe Excel mensual.")
        return out
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            out["hojas"] = wb.sheetnames
            ws = wb[EXCEL_SHEET_NAME] if EXCEL_SHEET_NAME in wb.sheetnames else wb[wb.sheetnames[0]]
            out["abre_ok"] = True
            out["filas_datos"] = max(0, ws.max_row - 1)
            out["columnas"] = ws.max_column
            if expected_columns and ws.max_column != expected_columns:
                out["errores"].append(f"Columnas esperadas={expected_columns}, encontradas={ws.max_column}.")
        finally:
            wb.close()
    except Exception as exc:
        out["errores"].append(f"No se pudo abrir Excel mensual: {type(exc).__name__}: {exc}")
    return out


def listar_archivos_para_manifest(base: Path, excluir: Optional[set[Path]] = None) -> list[dict[str, Any]]:
    excluir_resolved = {p.resolve() for p in (excluir or set()) if p.exists()}
    items = []
    for p in sorted(base.rglob("*")):
        if not p.is_file():
            continue
        try:
            if p.resolve() in excluir_resolved:
                continue
        except Exception:
            pass
        items.append(info_archivo(p, categoria="evidencia"))
    return items


def generar_resumen_txt(path: Path, manifest: dict[str, Any]) -> None:
    lines = [
        "CIERRE MENSUAL LOCAL SEGURO V1 - FACTURAS JOYCO",
        "=" * 90,
        f"Version: {manifest['version']}",
        f"Periodo: {manifest['periodo']}",
        f"Rango: {manifest['fecha_inicio']} a {manifest['fecha_fin']}",
        f"Generado: {manifest['generado_en']}",
        f"Root: {manifest['root']}",
        f"Carpeta cierre: {manifest['carpeta_cierre']}",
        "",
        "Resumen:",
        f"- Filas en Excel mensual: {manifest['excel_mensual']['filas_datos']}",
        f"- Columnas en Excel mensual: {manifest['excel_mensual']['columnas']}",
        f"- Metodo seleccion: {manifest['seleccion_filas'].get('metodo_seleccion')}",
        f"- Auditorias copiadas: {manifest['conteos']['auditorias']}",
        f"- Logs copiados: {manifest['conteos']['logs']}",
        f"- Soportes semanales copiados: {manifest['conteos']['soporte_semanas']}",
        f"- Total archivos evidencia: {manifest['total_archivos']}",
        f"- Total bytes evidencia: {manifest['total_bytes']}",
        "",
        "Regla aplicada:",
        "- Este cierre mensual se calcula desde data/facturas.xlsx + auditorias reales del mes.",
        "- No consolida desde cierres semanales ni diarios para evitar arrastrar errores previos.",
        "- Los cierres semanales se copian solo como soporte/evidencia, si existen.",
        "- No se incluye .env real ni secretos.",
        "=" * 90,
    ]
    if manifest.get("advertencias"):
        lines.extend(["", "Advertencias:"])
        lines.extend([f"- {x}" for x in manifest["advertencias"]])
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    args = _parse_args()
    inicio, fin = rango_desde_args(args)
    anio = inicio.strftime("%Y")
    mes_nombre = mes_carpeta(inicio)
    periodo = periodo_nombre(inicio, fin)

    mes_dir = CIERRES_DIR / anio / mes_nombre
    cierre_dir = mes_dir / "Mensual"

    excel_dir = cierre_dir / "01_Excel_Mensual"
    auditoria_dir = cierre_dir / "02_Auditorias_Mes"
    soporte_semanas_dir = cierre_dir / "03_Soporte_Semanas"
    manifest_dir = cierre_dir / "04_Manifest_Mensual"
    validaciones_dir = cierre_dir / "05_Validaciones"
    logs_dir = cierre_dir / "06_Logs_Mes"

    excel_path = excel_dir / f"facturas_mensual_{periodo}.xlsx"
    manifest_path = manifest_dir / f"manifest_mensual_{periodo}.json"
    resumen_path = manifest_dir / f"resumen_mensual_{periodo}.txt"
    validacion_path = validaciones_dir / f"validacion_local_mensual_{periodo}.json"

    print("=" * 100)
    print("CIERRE MENSUAL LOCAL SEGURO V1 - FACTURAS JOYCO")
    print("=" * 100)
    print(f"Version: {VERSION_CIERRE}")
    print(f"Root: {ROOT}")
    print(f"Periodo: {periodo}")
    print(f"Rango: {inicio.isoformat()} a {fin.isoformat()}")
    print(f"Mes carpeta: {mes_nombre}")
    print(f"Carpeta local destino: {cierre_dir}")
    print(f"Dry-run: {args.dry_run}")
    print("-" * 100)

    ok_lock, msg_lock = crear_lock()
    if not ok_lock:
        print(f"ERROR lock: {msg_lock}")
        return 2

    try:
        audit_files = recolectar_audits(inicio, fin)
        log_files = recolectar_logs(inicio, fin)
        cierres_semanales = buscar_cierres_semanales_soporte(inicio, fin, mes_dir)
        headers, filas_mensuales, meta_seleccion = obtener_filas_mensuales_desde_excel(inicio, fin, audit_files)

        advertencias: list[str] = []
        advertencias.extend(meta_seleccion.get("advertencias", []))
        if not audit_files:
            advertencias.append("No se encontraron archivos de auditoria dentro del rango mensual.")
        if not log_files:
            advertencias.append("No se encontraron logs especificos del mes.")
        if not filas_mensuales:
            advertencias.append("El Excel mensual no tendra filas de datos para este periodo.")
        if not cierres_semanales:
            advertencias.append("No se encontraron cierres semanales previos como soporte. El calculo mensual no depende de ellos.")

        print(f"Auditorias detectadas: {len(audit_files)}")
        print(f"Logs detectados: {len(log_files)}")
        print(f"Cierres semanales soporte detectados: {len(cierres_semanales)}")
        print(f"Metodo seleccion filas: {meta_seleccion.get('metodo_seleccion')}")
        print(f"Filas mensuales detectadas: {len(filas_mensuales)}")

        if args.dry_run:
            print("-" * 100)
            print("DRY-RUN: no se generaron archivos.")
            if advertencias:
                print("Advertencias:")
                for adv in advertencias:
                    print(f"  - {adv}")
            return 0

        if cierre_dir.exists():
            print(f"Ya existe carpeta de cierre mensual. Se regenerara: {cierre_dir}")
            shutil.rmtree(cierre_dir, ignore_errors=True)

        for d in (excel_dir, auditoria_dir, soporte_semanas_dir, manifest_dir, validaciones_dir, logs_dir):
            d.mkdir(parents=True, exist_ok=True)

        excel_info = crear_excel_mensual(excel_path, list(headers), filas_mensuales)
        auditorias_info = copiar_auditorias(audit_files, auditoria_dir)
        logs_info = copiar_logs(log_files, logs_dir)
        soporte_semanas_info = copiar_soporte_semanas(cierres_semanales, soporte_semanas_dir)
        env_info = crear_snapshot_env_redactado(manifest_dir, periodo)

        validacion_excel = validar_excel_generado(excel_path, expected_columns=len(headers))
        validacion = {
            "tipo": "validacion_local_cierre_mensual",
            "version": VERSION_CIERRE,
            "periodo": periodo,
            "fecha_inicio": inicio.isoformat(),
            "fecha_fin": fin.isoformat(),
            "generado_en": _dt.datetime.now().isoformat(timespec="seconds"),
            "excel": validacion_excel,
            "auditorias_detectadas": len(audit_files),
            "auditorias_copiadas": len(auditorias_info),
            "logs_detectados": len(log_files),
            "logs_copiados": len(logs_info),
            "cierres_semanales_soporte_detectados": len(cierres_semanales),
            "soportes_semanales_copiados": len(soporte_semanas_info),
            "ok": bool(validacion_excel.get("abre_ok")) and not validacion_excel.get("errores"),
            "advertencias": advertencias,
        }
        validacion_path.write_text(json.dumps(validacion, ensure_ascii=False, indent=2), encoding="utf-8")

        manifest = {
            "tipo": "cierre_mensual_local_seguro_v1",
            "version": VERSION_CIERRE,
            "anio": anio,
            "mes_carpeta": mes_nombre,
            "periodo": periodo,
            "fecha_inicio": inicio.isoformat(),
            "fecha_fin": fin.isoformat(),
            "generado_en": _dt.datetime.now().isoformat(timespec="seconds"),
            "root": str(ROOT),
            "host": socket.gethostname(),
            "platform": platform.platform(),
            "python": sys.version.replace("\n", " "),
            "carpeta_cierre": rel(cierre_dir),
            "excel_mensual": excel_info,
            "seleccion_filas": meta_seleccion,
            "fuente_datos": {
                "principal": rel(FACTURAS_PATH),
                "auditorias": [rel(p) for p in audit_files],
                "nota": "Los cierres semanales/diarios son soporte/evidencia, no fuente de calculo.",
            },
            "conteos": {
                "filas_mensuales": len(filas_mensuales),
                "auditorias": len(auditorias_info),
                "logs": len(logs_info),
                "soporte_semanas": len(soporte_semanas_info),
                "config_redactada": 1 if env_info else 0,
            },
            "validacion_local": rel(validacion_path),
            "advertencias": advertencias,
            "nota": "Cierre mensual V1: fuente oficial Excel principal + auditorias reales del mes.",
        }

        archivos_pre = listar_archivos_para_manifest(cierre_dir, excluir={manifest_path})
        manifest["archivos"] = archivos_pre
        manifest["total_archivos"] = len(archivos_pre)
        manifest["total_bytes"] = sum(int(x.get("bytes", 0) or 0) for x in archivos_pre)
        generar_resumen_txt(resumen_path, manifest)

        archivos = listar_archivos_para_manifest(cierre_dir, excluir={manifest_path})
        manifest["archivos"] = archivos
        manifest["total_archivos"] = len(archivos) + 1
        manifest["total_bytes"] = sum(int(x.get("bytes", 0) or 0) for x in archivos)
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

        print(f"Excel mensual: {excel_path}")
        print(f"Filas Excel mensual: {excel_info['filas_datos']}")
        print(f"Auditorias copiadas: {len(auditorias_info)}")
        print(f"Logs copiados: {len(logs_info)}")
        print(f"Soportes semanales copiados: {len(soporte_semanas_info)}")
        print(f"Manifest: {manifest_path}")
        print(f"Resumen: {resumen_path}")
        print(f"Validacion local: {validacion_path}")
        print(f"Total archivos evidencia: {manifest['total_archivos']}")
        print(f"Total bytes evidencia: {manifest['total_bytes']}")
        if advertencias:
            print("Advertencias:")
            for adv in advertencias:
                print(f"  - {adv}")

        if not args.permitir_vacio and len(filas_mensuales) == 0:
            print("Cierre generado sin filas mensuales. Codigo 3 para revision controlada.")
            print("Si el periodo realmente no tuvo facturas nuevas, usa --permitir-vacio.")
            return 3

        print("=" * 100)
        print("Cierre mensual local V1 generado correctamente.")
        print("Siguiente paso: subir/validar en OneDrive con scripts\\subir_cierre_mensual_sharepoint.py")
        print("=" * 100)
        return 0

    except Exception as exc:
        print(f"ERROR generando cierre mensual local V1: {type(exc).__name__}: {exc}")
        print("No subas ni archives nada hasta revisar el error.")
        return 1
    finally:
        liberar_lock()


if __name__ == "__main__":
    raise SystemExit(main())
