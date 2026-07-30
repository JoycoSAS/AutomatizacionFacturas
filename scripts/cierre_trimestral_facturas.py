# -*- coding: utf-8 -*-
"""
JOYCO - Facturas Procesador
Preparación segura del cierre trimestral de facturas.

Este script ejecuta únicamente la FASE 1 del flujo trimestral:

1. Validar el estado trimestral y el Excel activo.
2. Comprobar que exista un cierre mensual previo validado en el
   repositorio productivo de backups.
3. Crear el Excel histórico del periodo.
4. Crear respaldos técnicos, un Excel limpio candidato, manifiesto,
   resumen y estado de preparación.
5. Mantener intactos:
   - data/facturas.xlsx
   - data/state/cierre_trimestral_state.json
   - Excel activo de SharePoint

Modos:
- --dry-run:
  diagnostica y muestra el plan. No modifica archivos.
- --preparar --confirmar PREPARAR_CIERRE_TRIMESTRAL:
  prepara el histórico local de forma segura y atómica.
- --real y --local:
  se conservan únicamente para bloquear comandos antiguos y evitar
  que alguien ejecute el flujo inseguro anterior.

La limpieza del Excel activo y la actualización del estado pertenecen
a una fase posterior, que solo podrá ejecutarse después de validar el
cierre histórico en los dos destinos remotos.
"""

from __future__ import annotations

import argparse
import calendar
import hashlib
import json
import os
import re
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, Optional

from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
STATE_PATH = DATA_DIR / "state" / "cierre_trimestral_state.json"
FACTURAS_PATH = DATA_DIR / "facturas.xlsx"
CIERRES_DIARIOS_DIR = DATA_DIR / "cierres_diarios"
CIERRES_TRIMESTRALES_DIR = DATA_DIR / "cierres_trimestrales"

VERSION = "2026-07-30-CIERRE-TRIMESTRAL-V5-PREPARACION-SEGURA"
CONFIRMACION_PREPARAR = "PREPARAR_CIERRE_TRIMESTRAL"

ESTADO_PREPARADO = "PREPARADO_PENDIENTE_VALIDACION_REMOTA"
PREFIJO_VALIDACION_REMOTA = "validacion_remota_cierre_trimestral_"

load_dotenv(ROOT / ".env")

BACKUP_ROOT_FOLDER = (
    os.getenv("BACKUP_ROOT_FOLDER")
    or os.getenv("SP_BACKUP2_FOLDER")
    or ""
).strip().strip("/")

BACKUP_DRIVE_ID = (
    os.getenv("BACKUP_DRIVE_ID")
    or os.getenv("SP_BACKUP2_DRIVE_ID")
    or ""
).strip()

HEADERS_ESPERADOS = [
    "Radicado",
    "ProyectoProceso",
    "Archivo",
    "Empresa emisora",
    "CUFE",
    "Ciudad emisora",
    "Código ciudad",
    "NIT",
    "Cliente",
    "Número de factura",
    "Año",
    "Mes",
    "Día",
    "Tipo de contribuyente",
    "Actividad económica",
    "DESCRIPCIÓN",
    "Concepto",
    "VALOR",
    "Estado_calidad",
]


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as archivo:
        for chunk in iter(lambda: archivo.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def escribir_json_atomico(path: Path, datos: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporal = path.with_suffix(path.suffix + ".tmp")
    temporal.write_text(
        json.dumps(datos, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    temporal.replace(path)


def escribir_texto_atomico(path: Path, contenido: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporal = path.with_suffix(path.suffix + ".tmp")
    temporal.write_text(contenido, encoding="utf-8")
    temporal.replace(path)


def parse_fecha(valor: str, campo: str) -> date:
    try:
        return datetime.strptime(valor, "%Y-%m-%d").date()
    except Exception as exc:
        raise RuntimeError(
            f"Fecha inválida en {campo}: {valor}. "
            "Formato esperado YYYY-MM-DD."
        ) from exc


def add_months(fecha: date, meses: int) -> date:
    mes_total = fecha.month - 1 + meses
    anio = fecha.year + mes_total // 12
    mes = mes_total % 12 + 1
    dia = min(fecha.day, calendar.monthrange(anio, mes)[1])
    return date(anio, mes, dia)


def calcular_siguiente_periodo(fecha_fin_actual: date) -> Dict[str, str]:
    """
    Calcula el siguiente ciclo OPERATIVO de tres meses.

    No se asume que el periodo cerrado sea un trimestre calendario.
    Ejemplo:
      cierre actual: 2026-07-31
      siguiente: 2026-08-01 a 2026-10-31
    """
    siguiente_inicio = fecha_fin_actual + timedelta(days=1)
    siguiente_fin = add_months(siguiente_inicio, 3) - timedelta(days=1)

    periodo = (
        f"{siguiente_inicio.year}-CICLO_"
        f"{siguiente_inicio.strftime('%Y%m%d')}_A_"
        f"{siguiente_fin.strftime('%Y%m%d')}"
    )

    return {
        "periodo_activo": periodo,
        "fecha_inicio_periodo_activo": siguiente_inicio.isoformat(),
        "proximo_cierre_estimado": siguiente_fin.isoformat(),
    }


def cargar_estado() -> dict:
    if not STATE_PATH.exists():
        raise RuntimeError(
            f"No existe el estado trimestral: {STATE_PATH}\n"
            "Primero debe inicializarse el estado del periodo operativo "
            "que termina el 2026-07-31."
        )

    with STATE_PATH.open("r", encoding="utf-8-sig") as archivo:
        estado = json.load(archivo)

    requeridos = [
        "periodo_activo",
        "fecha_inicio_periodo_activo",
        "proximo_cierre_estimado",
        "estado",
    ]

    faltantes = [campo for campo in requeridos if not estado.get(campo)]
    if faltantes:
        raise RuntimeError(
            f"Estado trimestral incompleto. Faltan campos: {faltantes}"
        )

    if str(estado.get("estado")).upper() != "ACTIVO":
        raise RuntimeError(
            f"Estado trimestral no está ACTIVO: {estado.get('estado')}"
        )

    inicio = parse_fecha(
        str(estado["fecha_inicio_periodo_activo"]),
        "fecha_inicio_periodo_activo",
    )
    fin = parse_fecha(
        str(estado["proximo_cierre_estimado"]),
        "proximo_cierre_estimado",
    )

    if inicio > fin:
        raise RuntimeError(
            "El estado trimestral es inválido: la fecha de inicio "
            "es posterior a la fecha de cierre."
        )

    return estado


def diagnosticar_excel(path: Path = FACTURAS_PATH) -> dict:
    if not path.exists():
        raise RuntimeError(f"No existe el Excel: {path}")

    wb = load_workbook(
        path,
        read_only=False,
        data_only=False,
        keep_links=False,
    )

    try:
        if "Facturas" not in wb.sheetnames:
            raise RuntimeError(
                "El Excel no tiene la hoja requerida: Facturas"
            )

        ws = wb["Facturas"]
        headers = [cell.value for cell in ws[1]]

        if headers != HEADERS_ESPERADOS:
            raise RuntimeError(
                "Los encabezados del Excel no coinciden con la "
                "estructura esperada.\n"
                f"Detectados: {headers}\n"
                f"Esperados: {HEADERS_ESPERADOS}"
            )

        tablas = list(ws.tables.keys())
        tabla_ref = (
            ws.tables["TblFacturas"].ref
            if "TblFacturas" in ws.tables
            else None
        )

        if ws.max_column != 19:
            raise RuntimeError(
                "Cantidad de columnas inválida. "
                f"Detectadas: {ws.max_column}. Esperadas: 19."
            )

        if tabla_ref is None:
            raise RuntimeError(
                "No existe la tabla requerida TblFacturas."
            )

        return {
            "archivo": str(path),
            "hojas": list(wb.sheetnames),
            "hoja_principal": ws.title,
            "filas": int(ws.max_row or 0),
            "columnas": int(ws.max_column or 0),
            "filas_datos": max(int(ws.max_row or 0) - 1, 0),
            "tablas": tablas,
            "tbl_facturas_ref": tabla_ref,
            "bytes": path.stat().st_size,
            "sha256": sha256_file(path),
        }
    finally:
        wb.close()


def crear_excel_limpio(origen: Path, destino: Path) -> None:
    """
    Crea un candidato limpio con la misma estructura del Excel activo.

    Nunca modifica el archivo origen.
    """
    wb = load_workbook(
        origen,
        read_only=False,
        data_only=False,
        keep_links=False,
    )

    try:
        if "Facturas" not in wb.sheetnames:
            raise RuntimeError(
                "No se puede crear el candidato limpio: "
                "falta la hoja Facturas."
            )

        ws = wb["Facturas"]

        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)

        ws.tables.clear()

        tabla = Table(displayName="TblFacturas", ref="A1:S1")
        tabla.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(tabla)

        destino.parent.mkdir(parents=True, exist_ok=True)
        wb.save(destino)
    finally:
        wb.close()


def _entero_no_negativo(valor: Any) -> int:
    if isinstance(valor, bool):
        return int(valor)
    if isinstance(valor, int):
        return max(valor, 0)
    if isinstance(valor, float):
        return max(int(valor), 0)
    if isinstance(valor, str):
        try:
            return max(int(valor.strip()), 0)
        except Exception:
            return 0
    if isinstance(valor, (list, tuple, set, dict)):
        return len(valor)
    return 0


def _total_verificados(datos: dict) -> int:
    candidatos = [
        datos.get("total_archivos_verificados"),
        datos.get("archivos_verificados"),
        datos.get("total_verificados"),
        datos.get("verificados"),
    ]

    for valor in candidatos:
        numero = _entero_no_negativo(valor)
        if numero > 0:
            return numero

    resultados = datos.get("resultados")
    if isinstance(resultados, dict):
        for clave in (
            "total_archivos_verificados",
            "archivos_verificados",
            "verificados",
            "correctos",
        ):
            numero = _entero_no_negativo(resultados.get(clave))
            if numero > 0:
                return numero

    return 0


def _total_fallidos(datos: dict) -> int:
    candidatos = [
        datos.get("total_archivos_fallidos"),
        datos.get("archivos_fallidos"),
        datos.get("fallidos"),
        datos.get("errores"),
    ]

    total = 0
    for valor in candidatos:
        total = max(total, _entero_no_negativo(valor))

    resultados = datos.get("resultados")
    if isinstance(resultados, dict):
        for clave in (
            "total_archivos_fallidos",
            "archivos_fallidos",
            "fallidos",
            "errores",
        ):
            total = max(
                total,
                _entero_no_negativo(resultados.get(clave)),
            )

    return total


def _drive_id_validacion(datos: dict) -> str:
    drive = datos.get("drive")
    if isinstance(drive, dict):
        return str(drive.get("id") or "").strip()

    for clave in (
        "drive_id",
        "backup_drive_id",
        "sp_backup2_drive_id",
    ):
        valor = str(datos.get(clave) or "").strip()
        if valor:
            return valor

    return ""


def _carpeta_mensual_de(path: Path) -> Optional[Path]:
    actual = path.parent

    while actual != actual.parent:
        if actual.name.lower() == "mensual":
            return actual
        actual = actual.parent

    return None


def _es_validacion_remota_mensual(path: Path, datos: dict) -> bool:
    nombre = path.name.lower()

    if "validacion" not in nombre or "remot" not in nombre:
        return False

    if datos.get("ok") is not True:
        return False

    if _total_verificados(datos) <= 0:
        return False

    if _total_fallidos(datos) > 0:
        return False

    remote_base = str(
        datos.get("remote_base")
        or datos.get("ruta_remota")
        or datos.get("carpeta_remota")
        or ""
    ).strip()

    if not remote_base:
        return False

    if (
        BACKUP_ROOT_FOLDER
        and not remote_base.startswith(BACKUP_ROOT_FOLDER)
    ):
        return False

    drive_id = _drive_id_validacion(datos)

    if (
        BACKUP_DRIVE_ID
        and drive_id
        and drive_id != BACKUP_DRIVE_ID
    ):
        return False

    return _carpeta_mensual_de(path) is not None


def _archivos_evidencia_mensual(carpeta_mensual: Path) -> list[Path]:
    """
    Selecciona evidencia suficiente sin copiar miles de archivos fuente.
    """
    palabras = (
        "manifest",
        "resumen",
        "validacion",
        "retencion",
        "paquete",
        "mensual",
    )
    extensiones = {
        ".json",
        ".txt",
        ".xlsx",
        ".xlsm",
        ".zip",
        ".csv",
    }

    seleccionados: list[Path] = []

    for archivo in carpeta_mensual.rglob("*"):
        if not archivo.is_file():
            continue

        nombre = archivo.name.lower()

        if archivo.suffix.lower() not in extensiones:
            continue

        if not any(palabra in nombre for palabra in palabras):
            continue

        seleccionados.append(archivo)

    return sorted(
        set(seleccionados),
        key=lambda path: path.as_posix().lower(),
    )


def buscar_respaldo_mensual_validado() -> dict:
    if not CIERRES_DIARIOS_DIR.exists():
        raise RuntimeError(
            "No existe la carpeta de cierres diarios/mensuales: "
            f"{CIERRES_DIARIOS_DIR}"
        )

    candidatos: list[tuple[float, Path, dict]] = []

    for path in CIERRES_DIARIOS_DIR.rglob("*.json"):
        if "mensual" not in {
            parte.lower() for parte in path.parts
        }:
            continue

        try:
            datos = json.loads(
                path.read_text(encoding="utf-8-sig")
            )
        except Exception:
            continue

        if _es_validacion_remota_mensual(path, datos):
            candidatos.append(
                (path.stat().st_mtime, path, datos)
            )

    if not candidatos:
        detalle_destino = (
            f" bajo {BACKUP_ROOT_FOLDER}"
            if BACKUP_ROOT_FOLDER
            else ""
        )
        raise RuntimeError(
            "No se encontró un cierre mensual previo con validación "
            f"remota OK{detalle_destino}. "
            "Se bloquea la preparación trimestral."
        )

    _, validacion_path, datos = max(
        candidatos,
        key=lambda item: item[0],
    )

    carpeta_mensual = _carpeta_mensual_de(validacion_path)
    if carpeta_mensual is None:
        raise RuntimeError(
            "No fue posible resolver la carpeta Mensual de la "
            "validación seleccionada."
        )

    archivos = _archivos_evidencia_mensual(carpeta_mensual)

    if validacion_path not in archivos:
        archivos.append(validacion_path)
        archivos.sort(key=lambda path: path.as_posix().lower())

    if not archivos:
        raise RuntimeError(
            "La validación mensual es correcta, pero no se encontró "
            "evidencia local para copiar al cierre trimestral."
        )

    drive_id = _drive_id_validacion(datos)
    remote_base = str(
        datos.get("remote_base")
        or datos.get("ruta_remota")
        or datos.get("carpeta_remota")
        or ""
    ).strip()

    return {
        "carpeta_mensual": carpeta_mensual,
        "validacion_remota": validacion_path,
        "datos_validacion": datos,
        "archivos_evidencia": archivos,
        "total_verificados": _total_verificados(datos),
        "total_fallidos": _total_fallidos(datos),
        "remote_base": remote_base,
        "drive_id": drive_id,
        "generado_en": datos.get("generado_en"),
    }


def _nombre_seguro_periodo(periodo: str) -> str:
    limpio = re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        periodo.strip(),
    ).strip("._-")

    if not limpio:
        raise RuntimeError(
            f"Periodo activo inválido para nombre de carpeta: {periodo!r}"
        )

    return limpio


def datos_periodo(estado: dict) -> Dict[str, str]:
    periodo = str(estado["periodo_activo"]).strip()
    fecha_inicio = str(
        estado["fecha_inicio_periodo_activo"]
    ).strip()
    fecha_fin = str(estado["proximo_cierre_estimado"]).strip()

    inicio = parse_fecha(
        fecha_inicio,
        "fecha_inicio_periodo_activo",
    )
    fin = parse_fecha(
        fecha_fin,
        "proximo_cierre_estimado",
    )

    if inicio > fin:
        raise RuntimeError(
            "Periodo inválido: fecha de inicio posterior a fecha fin."
        )

    nombre_periodo = _nombre_seguro_periodo(periodo)

    prefijo_anio = f"{inicio.year}-"
    if nombre_periodo.startswith(prefijo_anio):
        nombre_carpeta_periodo = nombre_periodo[len(prefijo_anio):]
    else:
        nombre_carpeta_periodo = nombre_periodo

    if not nombre_carpeta_periodo:
        nombre_carpeta_periodo = nombre_periodo

    nombre_archivo = (
        f"facturas_{fecha_inicio}_a_{fecha_fin}.xlsx"
    )

    carpeta_base = (
        CIERRES_TRIMESTRALES_DIR
        / str(inicio.year)
        / nombre_carpeta_periodo
    )
    carpeta_excel = carpeta_base / "01_Excel_Cierre"
    carpeta_soportes = carpeta_base / "02_Soportes_Tecnicos"
    carpeta_respaldo_mensual = (
        carpeta_base / "03_Respaldo_Mensual_Validado"
    )
    destino_local = carpeta_excel / nombre_archivo

    return {
        "periodo": periodo,
        "nombre_periodo_seguro": nombre_periodo,
        "anio": str(inicio.year),
        "nombre_carpeta_periodo": nombre_carpeta_periodo,
        "fecha_inicio": fecha_inicio,
        "fecha_fin": fecha_fin,
        "nombre_archivo": nombre_archivo,
        "carpeta_base": str(carpeta_base),
        "carpeta_excel": str(carpeta_excel),
        "carpeta_soportes": str(carpeta_soportes),
        "carpeta_respaldo_mensual": str(
            carpeta_respaldo_mensual
        ),
        "destino_local": str(destino_local),
    }


def _resumen_drive(drive_id: str) -> str:
    if not drive_id:
        return "(no registrado)"

    return (
        f"***{drive_id[-6:]} "
        f"(longitud={len(drive_id)})"
    )


def imprimir_plan(
    estado: dict,
    info_excel: dict,
    respaldo_mensual: dict,
    modo: str,
) -> Dict[str, str]:
    periodo = datos_periodo(estado)
    siguiente = calcular_siguiente_periodo(
        parse_fecha(periodo["fecha_fin"], "fecha_fin")
    )

    print("✅ Estado trimestral cargado correctamente.")
    print(f"Periodo operativo: {periodo['periodo']}")
    print(f"Fecha inicio registrada: {periodo['fecha_inicio']}")
    print(f"Fecha de cierre: {periodo['fecha_fin']}")
    print(f"Carpeta histórica: {periodo['carpeta_base']}")

    print("-" * 100)
    print("✅ Excel activo validado.")
    print(f"Archivo: {FACTURAS_PATH}")
    print(f"Filas totales: {info_excel['filas']}")
    print(f"Filas de datos: {info_excel['filas_datos']}")
    print(f"Columnas: {info_excel['columnas']}")
    print(f"TblFacturas ref: {info_excel['tbl_facturas_ref']}")
    print(f"Bytes: {info_excel['bytes']}")
    print(f"SHA256: {info_excel['sha256']}")

    print("-" * 100)
    print("✅ Cierre mensual previo validado.")
    print(
        "Carpeta mensual: "
        f"{respaldo_mensual['carpeta_mensual']}"
    )
    print(
        "Validación remota: "
        f"{respaldo_mensual['validacion_remota']}"
    )
    print(
        "Destino remoto: "
        f"{respaldo_mensual['remote_base']}"
    )
    print(
        "Drive: "
        f"{_resumen_drive(respaldo_mensual['drive_id'])}"
    )
    print(
        "Archivos verificados: "
        f"{respaldo_mensual['total_verificados']}"
    )
    print(
        "Archivos de evidencia a conservar: "
        f"{len(respaldo_mensual['archivos_evidencia'])}"
    )

    print("-" * 100)
    print("PLAN DE PREPARACIÓN SEGURA:")
    print(
        "1. Crear copia histórica exacta del Excel activo:"
    )
    print(f"   {periodo['destino_local']}")
    print(
        "2. Crear respaldo técnico adicional del Excel activo."
    )
    print(
        "3. Copiar el estado trimestral previo sin modificarlo."
    )
    print(
        "4. Crear y validar un Excel limpio CANDIDATO, "
        "sin activarlo."
    )
    print(
        "5. Conservar la evidencia del cierre mensual validado."
    )
    print(
        "6. Crear manifiesto, resumen y estado de preparación."
    )
    print(
        "7. Mantener intactos data/facturas.xlsx y el estado activo."
    )
    print(
        "8. Dejar la finalización bloqueada hasta validar "
        "los dos destinos remotos."
    )
    print("-" * 100)
    print("SIGUIENTE CICLO PREVISTO:")
    print(f"Periodo: {siguiente['periodo_activo']}")
    print(
        "Inicio: "
        f"{siguiente['fecha_inicio_periodo_activo']}"
    )
    print(
        "Fin: "
        f"{siguiente['proximo_cierre_estimado']}"
    )

    if modo == "DRY RUN":
        print("-" * 100)
        print(
            "✅ Modo diagnóstico: no se creará ni modificará "
            "ningún archivo."
        )

    return periodo


def validar_fecha_preparacion(
    estado: dict,
    confirmar: Optional[str],
) -> date:
    if confirmar != CONFIRMACION_PREPARAR:
        raise RuntimeError(
            "Preparación real bloqueada. Debe usarse:\n"
            "python scripts\\cierre_trimestral_facturas.py "
            "--preparar --confirmar "
            f"{CONFIRMACION_PREPARAR}"
        )

    fecha_fin = parse_fecha(
        str(estado["proximo_cierre_estimado"]),
        "proximo_cierre_estimado",
    )
    hoy = date.today()

    if hoy < fecha_fin:
        raise RuntimeError(
            "Preparación bloqueada por fecha.\n"
            f"Periodo activo: {estado['periodo_activo']}\n"
            f"Fecha de cierre: {fecha_fin.isoformat()}\n"
            f"Fecha actual: {hoy.isoformat()}\n"
            "No se permite preparar el cierre antes de la fecha "
            "oficial del periodo."
        )

    return hoy


def _copiar_evidencia_mensual(
    respaldo_mensual: dict,
    destino_base: Path,
) -> list[dict]:
    origen_base = Path(respaldo_mensual["carpeta_mensual"])
    resultado: list[dict] = []

    for origen in respaldo_mensual["archivos_evidencia"]:
        origen = Path(origen)

        try:
            relativa = origen.relative_to(origen_base)
        except ValueError:
            relativa = Path(origen.name)

        destino = destino_base / relativa
        destino.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(origen, destino)

        hash_origen = sha256_file(origen)
        hash_destino = sha256_file(destino)

        if hash_origen != hash_destino:
            raise RuntimeError(
                "La evidencia mensual copiada no coincide con "
                f"el origen: {origen}"
            )

        resultado.append(
            {
                "ruta_origen": str(origen),
                "ruta_relativa_cierre": (
                    Path("03_Respaldo_Mensual_Validado")
                    / relativa
                ).as_posix(),
                "bytes": destino.stat().st_size,
                "sha256": hash_destino,
            }
        )

    return resultado


def _validar_preparacion_existente(
    periodo: dict,
    info_excel: dict,
) -> Optional[dict]:
    carpeta_base = Path(periodo["carpeta_base"])

    if not carpeta_base.exists():
        return None

    estado_path = (
        Path(periodo["carpeta_soportes"])
        / (
            "estado_preparacion_cierre_trimestral_"
            f"{periodo['nombre_periodo_seguro']}.json"
        )
    )

    if not estado_path.exists():
        raise RuntimeError(
            "La carpeta del periodo ya existe, pero no contiene "
            "un estado de preparación válido. Se requiere revisión "
            f"manual antes de continuar: {carpeta_base}"
        )

    datos = json.loads(
        estado_path.read_text(encoding="utf-8-sig")
    )

    if datos.get("estado") != ESTADO_PREPARADO:
        raise RuntimeError(
            "La preparación existente no está en el estado esperado: "
            f"{datos.get('estado')}"
        )

    hash_registrado = str(
        datos.get("excel_activo_sha256_al_preparar") or ""
    )

    if hash_registrado != info_excel["sha256"]:
        raise RuntimeError(
            "Ya existe una preparación para este periodo, pero "
            "data/facturas.xlsx cambió después de prepararla.\n"
            f"SHA preparación: {hash_registrado}\n"
            f"SHA actual:      {info_excel['sha256']}\n"
            "No se debe sobrescribir ni finalizar hasta revisar "
            "la diferencia."
        )

    historico = Path(
        str(datos.get("excel_historico") or "")
    )

    if not historico.exists():
        raise RuntimeError(
            "La preparación registrada no contiene el Excel "
            f"histórico esperado: {historico}"
        )

    if sha256_file(historico) != hash_registrado:
        raise RuntimeError(
            "El Excel histórico de la preparación existente no "
            "coincide con el hash registrado."
        )

    return datos


def _crear_manifest_preparacion(
    *,
    periodo: dict,
    info_excel: dict,
    estado_activo: dict,
    archivo_historico_stage: Path,
    respaldo_activo_stage: Path,
    respaldo_state_stage: Path,
    candidato_limpio_stage: Path,
    info_candidato: dict,
    evidencia_mensual: list[dict],
    respaldo_mensual: dict,
    generado_en: str,
) -> dict:
    final_base = Path(periodo["carpeta_base"])
    final_soportes = Path(periodo["carpeta_soportes"])

    archivo_historico_final = Path(periodo["destino_local"])
    respaldo_activo_final = (
        final_soportes / respaldo_activo_stage.name
    )
    respaldo_state_final = (
        final_soportes / respaldo_state_stage.name
    )
    candidato_limpio_final = (
        final_soportes / candidato_limpio_stage.name
    )

    siguiente = calcular_siguiente_periodo(
        parse_fecha(periodo["fecha_fin"], "fecha_fin")
    )

    return {
        "tipo": "PREPARACION_CIERRE_TRIMESTRAL",
        "version_script": VERSION,
        "generado_en": generado_en,
        "estado": ESTADO_PREPARADO,
        "finalizacion_autorizada": False,
        "root": str(ROOT),
        "periodo": periodo["periodo"],
        "nombre_periodo_seguro": (
            periodo["nombre_periodo_seguro"]
        ),
        "fecha_inicio": periodo["fecha_inicio"],
        "fecha_fin": periodo["fecha_fin"],
        "carpeta_periodo": str(final_base),
        "excel_activo": str(FACTURAS_PATH),
        "excel_activo_bytes_al_preparar": (
            info_excel["bytes"]
        ),
        "excel_activo_sha256_al_preparar": (
            info_excel["sha256"]
        ),
        "excel_info_al_preparar": info_excel,
        "excel_historico": str(archivo_historico_final),
        "excel_historico_bytes": (
            archivo_historico_stage.stat().st_size
        ),
        "excel_historico_sha256": sha256_file(
            archivo_historico_stage
        ),
        "respaldo_excel_activo": str(
            respaldo_activo_final
        ),
        "respaldo_excel_activo_bytes": (
            respaldo_activo_stage.stat().st_size
        ),
        "respaldo_excel_activo_sha256": sha256_file(
            respaldo_activo_stage
        ),
        "respaldo_estado_activo": str(
            respaldo_state_final
        ),
        "respaldo_estado_activo_bytes": (
            respaldo_state_stage.stat().st_size
        ),
        "respaldo_estado_activo_sha256": sha256_file(
            respaldo_state_stage
        ),
        "excel_limpio_candidato": str(
            candidato_limpio_final
        ),
        "excel_limpio_candidato_bytes": (
            candidato_limpio_stage.stat().st_size
        ),
        "excel_limpio_candidato_sha256": sha256_file(
            candidato_limpio_stage
        ),
        "excel_limpio_candidato_info": info_candidato,
        "estado_trimestral_al_preparar": estado_activo,
        "respaldo_mensual_validado": {
            "carpeta_origen": str(
                respaldo_mensual["carpeta_mensual"]
            ),
            "validacion_remota_origen": str(
                respaldo_mensual["validacion_remota"]
            ),
            "remote_base": respaldo_mensual[
                "remote_base"
            ],
            "drive_id": respaldo_mensual["drive_id"],
            "total_archivos_verificados": (
                respaldo_mensual["total_verificados"]
            ),
            "total_archivos_fallidos": (
                respaldo_mensual["total_fallidos"]
            ),
            "generado_en": respaldo_mensual[
                "generado_en"
            ],
            "archivos_evidencia": evidencia_mensual,
        },
        "siguiente_periodo_previsto": siguiente,
        "controles": {
            "excel_activo_no_reemplazado": True,
            "estado_trimestral_no_actualizado": True,
            "sharepoint_no_modificado": True,
            "requiere_validacion_remota_doble": True,
            "requiere_finalizacion_separada": True,
        },
        "nota": (
            "La preparación histórica quedó generada, pero la "
            "finalización permanece bloqueada. Antes de limpiar "
            "data/facturas.xlsx o actualizar el estado, el uploader "
            "trimestral debe publicar y verificar el cierre en los "
            "dos destinos remotos y generar su evidencia final."
        ),
    }


def _crear_resumen_preparacion(manifest: dict) -> str:
    siguiente = manifest["siguiente_periodo_previsto"]
    respaldo = manifest["respaldo_mensual_validado"]

    lineas = [
        "PREPARACIÓN SEGURA DEL CIERRE TRIMESTRAL",
        "=" * 80,
        f"Versión: {manifest['version_script']}",
        f"Generado en: {manifest['generado_en']}",
        f"Estado: {manifest['estado']}",
        "",
        "Periodo preparado:",
        f"- Nombre: {manifest['periodo']}",
        f"- Inicio registrado: {manifest['fecha_inicio']}",
        f"- Fecha de cierre: {manifest['fecha_fin']}",
        "",
        "Excel activo al preparar:",
        f"- Ruta: {manifest['excel_activo']}",
        (
            "- Filas de datos: "
            f"{manifest['excel_info_al_preparar']['filas_datos']}"
        ),
        (
            "- SHA256: "
            f"{manifest['excel_activo_sha256_al_preparar']}"
        ),
        "",
        "Excel histórico:",
        f"- Ruta: {manifest['excel_historico']}",
        f"- SHA256: {manifest['excel_historico_sha256']}",
        "",
        "Excel limpio candidato:",
        f"- Ruta: {manifest['excel_limpio_candidato']}",
        (
            "- Filas: "
            f"{manifest['excel_limpio_candidato_info']['filas']}"
        ),
        (
            "- Tabla: "
            f"{manifest['excel_limpio_candidato_info']['tbl_facturas_ref']}"
        ),
        (
            "- SHA256: "
            f"{manifest['excel_limpio_candidato_sha256']}"
        ),
        "",
        "Cierre mensual previo validado:",
        f"- Origen: {respaldo['carpeta_origen']}",
        f"- Ruta remota: {respaldo['remote_base']}",
        (
            "- Archivos remotos verificados: "
            f"{respaldo['total_archivos_verificados']}"
        ),
        (
            "- Evidencias conservadas: "
            f"{len(respaldo['archivos_evidencia'])}"
        ),
        "",
        "Siguiente ciclo previsto:",
        f"- Periodo: {siguiente['periodo_activo']}",
        (
            "- Inicio: "
            f"{siguiente['fecha_inicio_periodo_activo']}"
        ),
        (
            "- Fin: "
            f"{siguiente['proximo_cierre_estimado']}"
        ),
        "",
        "CONTROLES DE SEGURIDAD:",
        "- data/facturas.xlsx NO fue reemplazado.",
        "- El estado trimestral activo NO fue actualizado.",
        "- SharePoint y OneDrive NO fueron modificados.",
        "- La finalización NO está autorizada todavía.",
        (
            "- Debe ejecutarse y validarse primero la subida "
            "histórica a los dos destinos."
        ),
        "=" * 80,
    ]

    return "\n".join(lineas) + "\n"


def ejecutar_preparacion(
    periodo: dict,
    info_excel: dict,
    respaldo_mensual: dict,
    estado: dict,
) -> int:
    existente = _validar_preparacion_existente(
        periodo,
        info_excel,
    )

    if existente is not None:
        print("-" * 100)
        print(
            "✅ El cierre trimestral ya estaba preparado "
            "con el mismo Excel activo."
        )
        print(
            "Estado: "
            f"{existente.get('estado')}"
        )
        print(
            "Excel histórico: "
            f"{existente.get('excel_historico')}"
        )
        print(
            "✅ No se creó ni modificó ningún archivo."
        )
        print("=" * 100)
        return 0

    final_base = Path(periodo["carpeta_base"])
    final_parent = final_base.parent
    final_parent.mkdir(parents=True, exist_ok=True)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stage = final_parent / (
        f".tmp_preparacion_"
        f"{periodo['nombre_carpeta_periodo']}_{stamp}"
    )

    if stage.exists():
        shutil.rmtree(stage)

    stage_excel = stage / "01_Excel_Cierre"
    stage_soportes = stage / "02_Soportes_Tecnicos"
    stage_mensual = stage / "03_Respaldo_Mensual_Validado"

    stage_excel.mkdir(parents=True, exist_ok=False)
    stage_soportes.mkdir(parents=True, exist_ok=False)
    stage_mensual.mkdir(parents=True, exist_ok=False)

    hash_activo_inicial = info_excel["sha256"]
    generado_en = datetime.now().isoformat(timespec="seconds")

    try:
        historico_stage = (
            stage_excel / periodo["nombre_archivo"]
        )
        shutil.copy2(FACTURAS_PATH, historico_stage)

        respaldo_activo_stage = stage_soportes / (
            "facturas_ACTIVO_RESPALDO_"
            f"{periodo['nombre_periodo_seguro']}_{stamp}.xlsx"
        )
        shutil.copy2(FACTURAS_PATH, respaldo_activo_stage)

        respaldo_state_stage = stage_soportes / (
            "cierre_trimestral_state_ANTES_"
            f"{periodo['nombre_periodo_seguro']}_{stamp}.json"
        )
        shutil.copy2(STATE_PATH, respaldo_state_stage)

        candidato_limpio_stage = stage_soportes / (
            "facturas_NUEVO_CANDIDATO_"
            f"{periodo['nombre_periodo_seguro']}_{stamp}.xlsx"
        )
        crear_excel_limpio(
            FACTURAS_PATH,
            candidato_limpio_stage,
        )

        info_candidato = diagnosticar_excel(
            candidato_limpio_stage
        )

        if (
            info_candidato["filas"] != 1
            or info_candidato["columnas"] != 19
            or info_candidato["tbl_facturas_ref"] != "A1:S1"
        ):
            raise RuntimeError(
                "Excel limpio candidato inválido: "
                f"{info_candidato}"
            )

        evidencia_mensual = _copiar_evidencia_mensual(
            respaldo_mensual,
            stage_mensual,
        )

        hash_activo_final = sha256_file(FACTURAS_PATH)

        if hash_activo_final != hash_activo_inicial:
            raise RuntimeError(
                "data/facturas.xlsx cambió durante la preparación.\n"
                f"SHA inicial: {hash_activo_inicial}\n"
                f"SHA final:   {hash_activo_final}\n"
                "Se cancela la preparación para evitar un cierre "
                "inconsistente."
            )

        for copia in (
            historico_stage,
            respaldo_activo_stage,
        ):
            if sha256_file(copia) != hash_activo_inicial:
                raise RuntimeError(
                    "Una copia del Excel activo no coincide con "
                    f"el original: {copia}"
                )

        manifest = _crear_manifest_preparacion(
            periodo=periodo,
            info_excel=info_excel,
            estado_activo=estado,
            archivo_historico_stage=historico_stage,
            respaldo_activo_stage=respaldo_activo_stage,
            respaldo_state_stage=respaldo_state_stage,
            candidato_limpio_stage=candidato_limpio_stage,
            info_candidato=info_candidato,
            evidencia_mensual=evidencia_mensual,
            respaldo_mensual=respaldo_mensual,
            generado_en=generado_en,
        )

        manifest_name = (
            "manifest_cierre_trimestral_"
            f"{periodo['nombre_periodo_seguro']}.json"
        )
        resumen_name = (
            "RESUMEN_CIERRE_TRIMESTRAL_"
            f"{periodo['nombre_periodo_seguro']}.txt"
        )
        estado_name = (
            "estado_preparacion_cierre_trimestral_"
            f"{periodo['nombre_periodo_seguro']}.json"
        )

        manifest_stage = stage_soportes / manifest_name
        resumen_stage = stage_soportes / resumen_name
        estado_stage = stage_soportes / estado_name

        escribir_json_atomico(manifest_stage, manifest)
        escribir_texto_atomico(
            resumen_stage,
            _crear_resumen_preparacion(manifest),
        )

        estado_preparacion = {
            "tipo": "ESTADO_PREPARACION_CIERRE_TRIMESTRAL",
            "version_script": VERSION,
            "generado_en": generado_en,
            "estado": ESTADO_PREPARADO,
            "finalizacion_autorizada": False,
            "periodo": periodo["periodo"],
            "fecha_inicio": periodo["fecha_inicio"],
            "fecha_fin": periodo["fecha_fin"],
            "carpeta_periodo": periodo["carpeta_base"],
            "manifest": str(
                Path(periodo["carpeta_soportes"])
                / manifest_name
            ),
            "resumen": str(
                Path(periodo["carpeta_soportes"])
                / resumen_name
            ),
            "excel_activo": str(FACTURAS_PATH),
            "excel_activo_sha256_al_preparar": (
                hash_activo_inicial
            ),
            "excel_historico": periodo["destino_local"],
            "excel_historico_sha256": sha256_file(
                historico_stage
            ),
            "excel_limpio_candidato": str(
                Path(periodo["carpeta_soportes"])
                / candidato_limpio_stage.name
            ),
            "excel_limpio_candidato_sha256": sha256_file(
                candidato_limpio_stage
            ),
            "validacion_remota_requerida": str(
                Path(periodo["carpeta_soportes"])
                / (
                    PREFIJO_VALIDACION_REMOTA
                    + periodo["nombre_carpeta_periodo"]
                    + ".json"
                )
            ),
            "nota": (
                "Preparación completa. La finalización sigue "
                "bloqueada hasta que exista una validación remota "
                "OK en ambos destinos y el Excel activo conserve "
                "el mismo SHA256."
            ),
        }

        escribir_json_atomico(
            estado_stage,
            estado_preparacion,
        )

        if final_base.exists():
            raise RuntimeError(
                "La carpeta final apareció durante la preparación. "
                f"No se sobrescribirá: {final_base}"
            )

        stage.replace(final_base)

        if sha256_file(FACTURAS_PATH) != hash_activo_inicial:
            raise RuntimeError(
                "El Excel activo cambió justo después de publicar "
                "la preparación. No debe continuarse con la subida."
            )

        estado_actual = cargar_estado()
        if estado_actual != estado:
            raise RuntimeError(
                "El estado trimestral cambió durante la preparación. "
                "No debe continuarse con la subida."
            )

        print("-" * 100)
        print("✅ PREPARACIÓN TRIMESTRAL SEGURA COMPLETADA.")
        print(f"Carpeta: {final_base}")
        print(
            "Excel histórico: "
            f"{Path(periodo['destino_local'])}"
        )
        print(
            "Manifest: "
            f"{Path(periodo['carpeta_soportes']) / manifest_name}"
        )
        print(
            "Estado de preparación: "
            f"{Path(periodo['carpeta_soportes']) / estado_name}"
        )
        print("-" * 100)
        print("✅ data/facturas.xlsx permanece intacto.")
        print(
            "✅ cierre_trimestral_state.json permanece intacto."
        )
        print("✅ No se tocó SharePoint ni OneDrive.")
        print(
            "⚠️ La finalización continúa bloqueada hasta validar "
            "la subida en ambos destinos."
        )
        print("=" * 100)
        return 0

    except Exception:
        if stage.exists():
            shutil.rmtree(stage, ignore_errors=True)
        raise


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Diagnostica sin modificar archivos.",
    )
    parser.add_argument(
        "--preparar",
        action="store_true",
        help=(
            "Prepara el histórico local sin limpiar el Excel activo "
            "ni actualizar el estado."
        ),
    )
    parser.add_argument(
        "--confirmar",
        default=None,
        help="Confirmación obligatoria para --preparar.",
    )
    parser.add_argument(
        "--real",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--local",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    args = parser.parse_args()

    if args.real or args.local:
        print("=" * 100)
        print("CIERRE TRIMESTRAL FACTURAS - COMANDO ANTIGUO BLOQUEADO")
        print("=" * 100)
        print(
            "❌ Los modos --real y --local fueron retirados porque "
            "pertenecían al flujo inseguro anterior."
        )
        print(
            "✅ Usa --dry-run o la fase segura "
            "--preparar --confirmar "
            f"{CONFIRMACION_PREPARAR}."
        )
        print(
            "⚠️ Este script nunca reemplaza data/facturas.xlsx "
            "ni actualiza el estado."
        )
        print("=" * 100)
        return 1

    if args.dry_run and args.preparar:
        print(
            "❌ Usa solo un modo: --dry-run o --preparar."
        )
        return 1

    modo = "PREPARAR" if args.preparar else "DRY RUN"

    print("=" * 100)
    print(f"CIERRE TRIMESTRAL FACTURAS - {modo}")
    print("=" * 100)
    print(f"Versión: {VERSION}")
    print(f"Root: {ROOT}")
    print("-" * 100)

    try:
        estado = cargar_estado()

        if args.preparar:
            validar_fecha_preparacion(
                estado,
                args.confirmar,
            )

        info_excel = diagnosticar_excel()
        respaldo_mensual = (
            buscar_respaldo_mensual_validado()
        )
        periodo = imprimir_plan(
            estado,
            info_excel,
            respaldo_mensual,
            modo,
        )

        if args.preparar:
            return ejecutar_preparacion(
                periodo,
                info_excel,
                respaldo_mensual,
                estado,
            )

        print("-" * 100)
        print(
            "✅ DRY RUN finalizado. "
            "No se modificó ningún archivo."
        )
        print("=" * 100)
        return 0

    except Exception as exc:
        print(f"❌ Error en cierre trimestral: {exc}")
        print(
            "No se debe continuar con la subida ni la "
            "finalización hasta revisar el error."
        )
        print("=" * 100)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
