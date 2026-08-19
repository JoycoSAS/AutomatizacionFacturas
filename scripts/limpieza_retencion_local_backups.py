# -*- coding: utf-8 -*-
"""
JOYCO - Limpieza segura de backups locales.

Funciones:
- Diagnostico de retencion local.
- Validacion de evidencia e integridad local.
- Validacion de evidencia remota almacenada.
- Revalidacion viva contra Microsoft Graph mediante GET/descarga.
- Eliminacion exclusivamente local, individual y controlada,
  solo cuando todas las barreras de seguridad se cumplen.
- Nunca elimina archivos de OneDrive/SharePoint.

Controles obligatorios para eliminacion:
1. Retencion vencida.
2. Validacion local correcta.
3. Evidencia remota almacenada correcta.
4. Revalidacion viva contra Graph correcta.
5. Identidad exacta del cierre.
6. Confirmacion explicita.
7. Variables de habilitacion y alertas activas.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import sys
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from dotenv import dotenv_values


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
BASE = ROOT / "data" / "cierres_diarios"
FACTURAS_ENV = Path(
    "/etc/joyco/facturas-procesador/facturas.env"
)

VERSION = (
    "2026-08-19-LIMPIEZA-RETENCION-LOCAL-"
    "V7-EJECUCION-LOCAL-CONTROLADA"
)

CONFIRMACION_EJECUCION = "ELIMINAR_SOLO_BACKUPS_LOCALES_VALIDADOS"
ENV_LIMPIEZA_HABILITADA = "JOYCO_LIMPIEZA_LOCAL_HABILITADA"
ENV_ALERTAS_HABILITADAS = "JOYCO_ALERTAS_LIMPIEZA_HABILITADAS"

ZONA_HORARIA = ZoneInfo("America/Bogota")

RETENCIONES = {
    "DIARIO": 7,
    "SEMANAL": 15,
    "MENSUAL": 20,
    "TRIMESTRAL": 30,
}

PREFIJOS_ESTADO = {
    "DIARIO": "estado_retencion_diario_",
    "SEMANAL": "estado_retencion_semanal_",
    "MENSUAL": "estado_retencion_mensual_",
    "TRIMESTRAL": "estado_retencion_trimestral_",
}


class ErrorValidacion(RuntimeError):
    pass


def leer_json(path: Path, descripcion: str) -> dict[str, Any]:
    if not path.is_file():
        raise ErrorValidacion(
            f"No existe {descripcion}: {path}"
        )

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise ErrorValidacion(
            f"No se pudo leer {descripcion}: "
            f"{type(exc).__name__}: {exc}"
        ) from exc

    if not isinstance(data, dict):
        raise ErrorValidacion(
            f"{descripcion} no contiene un objeto JSON."
        )

    return data


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()

    with path.open("rb") as f:
        for bloque in iter(lambda: f.read(1024 * 1024), b""):
            h.update(bloque)

    return h.hexdigest()


def dentro_de(path: Path, base: Path) -> bool:
    try:
        path.resolve().relative_to(base.resolve())
        return True
    except ValueError:
        return False


def parse_datetime(valor: str) -> dt.datetime:
    try:
        resultado = dt.datetime.fromisoformat(valor)
    except Exception as exc:
        raise ErrorValidacion(
            f"Fecha/hora invalida: {valor!r}"
        ) from exc

    if resultado.tzinfo is None:
        resultado = resultado.replace(
            tzinfo=ZONA_HORARIA
        )

    return resultado.astimezone(ZONA_HORARIA)


def resolver_archivo_manifest(
    cierre_dir: Path,
    entrada: dict[str, Any],
) -> Path:
    relativa = str(
        entrada.get("ruta_relativa") or ""
    ).strip()

    if relativa:
        candidato = cierre_dir / relativa
    else:
        destino = str(
            entrada.get("destino") or ""
        ).strip()

        if not destino:
            raise ErrorValidacion(
                "Entrada del manifest sin ruta_relativa "
                "ni destino."
            )

        candidato = Path(destino)

        if candidato.is_absolute():
            if not dentro_de(candidato, cierre_dir):
                partes = candidato.parts

                indices = [
                    i
                    for i, parte in enumerate(partes)
                    if parte == "cierres_diarios"
                ]

                if not indices:
                    raise ErrorValidacion(
                        "Ruta absoluta del manifest fuera "
                        "de la jerarquia cierres_diarios: "
                        f"{candidato}"
                    )

                indice = indices[-1]
                sufijo = partes[indice + 1:]

                if not sufijo:
                    raise ErrorValidacion(
                        "Ruta absoluta del manifest sin "
                        "archivo relativo a cierres_diarios."
                    )

                candidato = BASE.joinpath(*sufijo)
        else:
            candidato = ROOT / candidato

    if not dentro_de(candidato, cierre_dir):
        raise ErrorValidacion(
            "El manifest referencia un archivo fuera "
            f"del cierre real: {candidato}"
        )

    return candidato


def validar_archivos_manifest(
    cierre_dir: Path,
    manifest: dict[str, Any],
) -> tuple[bool, list[str], int]:
    errores: list[str] = []

    archivos = manifest.get("archivos")

    if not isinstance(archivos, list) or not archivos:
        return False, [
            "Manifest sin lista de archivos valida."
        ], 0

    verificados = 0

    for entrada in archivos:
        if not isinstance(entrada, dict):
            errores.append(
                "Entrada de manifest no es objeto."
            )
            continue

        try:
            archivo = resolver_archivo_manifest(
                cierre_dir,
                entrada,
            )

            if not archivo.is_file():
                errores.append(
                    f"No existe: {archivo}"
                )
                continue

            esperado_bytes = entrada.get("bytes")

            if esperado_bytes is not None:
                if archivo.stat().st_size != int(
                    esperado_bytes
                ):
                    errores.append(
                        f"Bytes distintos: {archivo}"
                    )
                    continue

            esperado_sha = str(
                entrada.get("sha256") or ""
            ).strip().lower()

            if esperado_sha:
                actual_sha = sha256_file(archivo)

                if actual_sha.lower() != esperado_sha:
                    errores.append(
                        f"SHA256 distinto: {archivo}"
                    )
                    continue

            verificados += 1

        except Exception as exc:
            errores.append(
                f"{type(exc).__name__}: {exc}"
            )

    return (
        len(errores) == 0
        and verificados == len(archivos),
        errores,
        verificados,
    )


def datos_candidato(
    nivel: str,
    estado_path: Path,
) -> dict[str, Any]:
    estado = leer_json(
        estado_path,
        "estado de retencion",
    )

    cierre_dir = estado_path.parent.parent

    if nivel == "DIARIO":
        fecha = str(estado.get("fecha") or "").strip()

        if not fecha:
            raise ErrorValidacion(
                "Estado diario sin fecha."
            )

        identidad = fecha

        manifest_path = (
            cierre_dir
            / "04_Manifest"
            / f"manifest_diario_{fecha}.json"
        )

        local_path = (
            cierre_dir
            / "05_Validaciones"
            / f"validacion_local_{fecha}.json"
        )

        remota_path = (
            cierre_dir
            / "05_Validaciones"
            / f"validacion_remota_{fecha}.json"
        )

    elif nivel == "SEMANAL":
        inicio = str(
            estado.get("semana_inicio") or ""
        ).strip()
        fin = str(
            estado.get("semana_fin") or ""
        ).strip()

        if not inicio or not fin:
            raise ErrorValidacion(
                "Estado semanal sin rango."
            )

        identidad = f"{inicio}_a_{fin}"

        manifest_path = (
            cierre_dir
            / "04_Manifest_Semanal"
            / f"manifest_semanal_{identidad}.json"
        )

        local_path = (
            cierre_dir
            / "05_Validaciones"
            / f"validacion_local_semanal_{identidad}.json"
        )

        remota_path = (
            cierre_dir
            / "05_Validaciones"
            / f"validacion_remota_semanal_{identidad}.json"
        )

    elif nivel == "MENSUAL":
        periodo = str(
            estado.get("periodo") or ""
        ).strip()

        if not periodo:
            raise ErrorValidacion(
                "Estado mensual sin periodo."
            )

        identidad = periodo

        manifest_path = (
            cierre_dir
            / "04_Manifest_Mensual"
            / f"manifest_mensual_{periodo}.json"
        )

        local_path = (
            cierre_dir
            / "05_Validaciones"
            / f"validacion_local_mensual_{periodo}.json"
        )

        remota_path = (
            cierre_dir
            / "05_Validaciones"
            / f"validacion_remota_mensual_{periodo}.json"
        )

    elif nivel == "TRIMESTRAL":
        prefijo = PREFIJOS_ESTADO["TRIMESTRAL"]
        nombre = estado_path.name

        if not (
            nombre.startswith(prefijo)
            and nombre.endswith(".json")
        ):
            raise ErrorValidacion(
                "Nombre de estado trimestral invalido."
            )

        identidad = nombre[
            len(prefijo):-5
        ]

        manifest_path = (
            cierre_dir
            / "02_Manifest"
            / f"manifest_cierre_trimestral_{identidad}.json"
        )

        local_path = (
            cierre_dir
            / "03_Validaciones"
            / f"validacion_local_cierre_trimestral_{identidad}.json"
        )

        remota_path = (
            cierre_dir
            / "03_Validaciones"
            / f"validacion_remota_cierre_trimestral_{identidad}.json"
        )

    else:
        raise ErrorValidacion(
            f"Nivel no soportado: {nivel}"
        )

    return {
        "nivel": nivel,
        "identidad": identidad,
        "cierre_dir": cierre_dir,
        "estado_path": estado_path,
        "estado": estado,
        "manifest_path": manifest_path,
        "local_path": local_path,
        "remota_path": remota_path,
    }


def validar_local(
    candidato: dict[str, Any],
) -> tuple[bool, list[str], dict[str, Any]]:
    errores: list[str] = []

    try:
        local = leer_json(
            candidato["local_path"],
            "validacion local",
        )
    except Exception as exc:
        return False, [str(exc)], {}

    if local.get("ok") is not True:
        errores.append(
            "validacion local tiene ok != true"
        )

    try:
        manifest = leer_json(
            candidato["manifest_path"],
            "manifest",
        )
    except Exception as exc:
        return False, errores + [str(exc)], local

    nivel = candidato["nivel"]

    if nivel in {
        "DIARIO",
        "SEMANAL",
        "MENSUAL",
    }:
        ok_manifest, errores_manifest, total = (
            validar_archivos_manifest(
                candidato["cierre_dir"],
                manifest,
            )
        )

        if not ok_manifest:
            errores.extend(errores_manifest)

        if nivel == "MENSUAL":
            integridad = local.get(
                "integridad_manifest"
            )

            if not isinstance(
                integridad,
                dict,
            ) or integridad.get("ok") is not True:
                errores.append(
                    "integridad_manifest mensual no OK"
                )

        detalle = {
            "archivos_manifest_verificados": total,
        }

    else:
        # El manifest trimestral contiene inventario de
        # toda la jerarquia inferior. Esa jerarquia NO
        # se usa como lista de eliminacion.
        excel = Path(
            str(
                manifest.get(
                    "excel_historico"
                ) or ""
            )
        )

        if not excel.is_absolute():
            excel = ROOT / excel

        if (
            not excel.is_file()
            or not dentro_de(
                excel,
                candidato["cierre_dir"],
            )
        ):
            errores.append(
                "Excel historico trimestral inexistente "
                "o fuera de Trimestral."
            )
        else:
            esperado = str(
                manifest.get(
                    "excel_historico_sha256"
                ) or ""
            ).lower()

            if (
                esperado
                and sha256_file(excel).lower()
                != esperado
            ):
                errores.append(
                    "SHA256 del Excel historico "
                    "trimestral no coincide."
                )

        inventario = Path(
            str(
                manifest.get(
                    "inventario_jerarquia"
                ) or ""
            )
        )

        if not inventario.is_absolute():
            inventario = ROOT / inventario

        if (
            not inventario.is_file()
            or not dentro_de(
                inventario,
                candidato["cierre_dir"],
            )
        ):
            errores.append(
                "Inventario trimestral inexistente "
                "o fuera de Trimestral."
            )
        else:
            esperado_inv = str(
                manifest.get(
                    "inventario_jerarquia_sha256"
                ) or ""
            ).lower()

            if (
                esperado_inv
                and sha256_file(inventario).lower()
                != esperado_inv
            ):
                errores.append(
                    "SHA256 del inventario trimestral "
                    "no coincide."
                )

        controles = manifest.get(
            "controles",
            {},
        )

        if controles.get(
            "cierres_inferiores_copiados"
        ) is not False:
            errores.append(
                "Control trimestral "
                "cierres_inferiores_copiados invalido."
            )

        if controles.get(
            "zip_trimestral_creado"
        ) is not False:
            errores.append(
                "Control trimestral ZIP invalido."
            )

        detalle = {
            "archivos_manifest_verificados":
                "CONTROL_TRIMESTRAL_ESPECIFICO",
        }

    return (
        len(errores) == 0,
        errores,
        {
            **detalle,
            "local": local,
            "manifest": manifest,
        },
    )


def validar_evidencia_remota_almacenada(
    candidato: dict[str, Any],
) -> tuple[bool, list[str], dict[str, Any]]:
    errores: list[str] = []

    try:
        remota = leer_json(
            candidato["remota_path"],
            "validacion remota almacenada",
        )
    except Exception as exc:
        return False, [str(exc)], {}

    nivel = candidato["nivel"]

    if remota.get("ok") is not True:
        errores.append(
            "validacion remota tiene ok != true"
        )

    if nivel in {
        "DIARIO",
        "SEMANAL",
        "MENSUAL",
    }:
        resultados = remota.get("resultados")

        if (
            not isinstance(resultados, list)
            or not resultados
        ):
            errores.append(
                "validacion remota sin resultados"
            )
        elif not all(
            isinstance(x, dict)
            and x.get("ok") is True
            for x in resultados
        ):
            errores.append(
                "existe resultado remoto no OK"
            )

        total = remota.get(
            "total_archivos_verificados"
        )

        if (
            isinstance(resultados, list)
            and total is not None
            and int(total) != len(resultados)
        ):
            errores.append(
                "total_archivos_verificados "
                "no coincide con resultados"
            )

    else:
        if remota.get("tipo") != (
            "VALIDACION_REMOTA_ESTRUCTURA_TRIMESTRAL"
        ):
            errores.append(
                "tipo de validacion remota "
                "trimestral invalido"
            )

        if remota.get("estado") != (
            "JERARQUIA_COMPLETA_VALIDADA_EN_DESTINO_UNICO"
        ):
            errores.append(
                "estado remoto trimestral invalido"
            )

        if remota.get(
            "destino_unico"
        ) is not True:
            errores.append(
                "destino_unico trimestral != true"
            )

    estado = candidato["estado"]

    if estado.get("ok") is not True:
        errores.append(
            "estado de retencion tiene ok != true"
        )

    if estado.get(
        "validacion_remota_publicada_y_verificada"
    ) is not True:
        errores.append(
            "estado no confirma publicacion/verificacion remota"
        )

    esperado_dias = RETENCIONES[nivel]

    if int(
        estado.get("retencion_local_dias") or -1
    ) != esperado_dias:
        errores.append(
            "dias de retencion no coinciden "
            f"con politica: {esperado_dias}"
        )

    remote_estado = str(
        estado.get("remote_base") or ""
    ).strip("/")

    remote_validacion = str(
        remota.get("remote_base") or ""
    ).strip("/")

    if (
        remote_estado
        and remote_validacion
        and remote_estado != remote_validacion
    ):
        errores.append(
            "remote_base de retencion y "
            "validacion remota no coinciden"
        )

    return (
        len(errores) == 0,
        errores,
        remota,
    )


def descubrir_estados(
    nivel: str,
) -> list[Path]:
    prefijo = PREFIJOS_ESTADO[nivel]

    return sorted(
        p
        for p in BASE.rglob(
            f"{prefijo}*.json"
        )
        if p.is_file()
    )


def evaluar_candidato(
    candidato: dict[str, Any],
    ahora: dt.datetime,
) -> dict[str, Any]:
    estado = candidato["estado"]

    errores: list[str] = []

    ok_local, err_local, info_local = (
        validar_local(candidato)
    )

    errores.extend(
        f"LOCAL: {x}"
        for x in err_local
    )

    ok_remota, err_remota, remota = (
        validar_evidencia_remota_almacenada(
            candidato
        )
    )

    errores.extend(
        f"REMOTA_ALMACENADA: {x}"
        for x in err_remota
    )

    vencimiento_raw = str(
        estado.get(
            "eliminacion_local_permitida_desde"
        ) or ""
    ).strip()

    vencida = False
    vencimiento = None

    if not vencimiento_raw:
        errores.append(
            "RETENCION: sin fecha de vencimiento"
        )
    else:
        try:
            vencimiento = parse_datetime(
                vencimiento_raw
            )
            vencida = ahora >= vencimiento
        except Exception as exc:
            errores.append(
                f"RETENCION: {exc}"
            )

    return {
        "nivel": candidato["nivel"],
        "identidad": candidato["identidad"],
        "cierre_dir": str(
            candidato["cierre_dir"]
        ),
        "ok_validacion_local": ok_local,
        "ok_validacion_remota_almacenada":
            ok_remota,
        "retencion_vencida": vencida,
        "vencimiento": (
            vencimiento.isoformat(
                timespec="seconds"
            )
            if vencimiento
            else None
        ),
        "preaprobado_para_validacion_viva": (
            ok_local
            and ok_remota
            and vencida
            and not errores
        ),
        "errores": errores,
        "detalle_local": {
            "archivos_manifest_verificados":
                info_local.get(
                    "archivos_manifest_verificados"
                )
        },
        "remote_base": (
            remota.get("remote_base")
            if isinstance(remota, dict)
            else None
        ),
    }


def cargar_modulo_graph_solo_lectura():
    """
    Carga utilidades Graph existentes.

    Esta funcion NO ejecuta escrituras.
    Las unicas funciones usadas posteriormente son:
    - validar_drive          -> GET
    - obtener_item_por_path -> GET
    - comparar_local_remoto -> GET/download
    """
    if not FACTURAS_ENV.is_file():
        raise ErrorValidacion(
            f"No existe entorno protegido: {FACTURAS_ENV}"
        )

    for clave, valor in dotenv_values(
        FACTURAS_ENV
    ).items():
        if valor is not None:
            os.environ[clave] = valor

    if str(SCRIPTS) not in sys.path:
        sys.path.insert(0, str(SCRIPTS))

    import subir_cierre_trimestral_sharepoint as graph_mod

    if not str(
        graph_mod.BACKUP_DRIVE_ID or ""
    ).strip():
        raise ErrorValidacion(
            "BACKUP_DRIVE_ID no esta configurado."
        )

    if graph_mod.BACKUP_ROOT_FOLDER != (
        "Backups_Facturas_Produccion"
    ):
        raise ErrorValidacion(
            "El destino configurado no coincide con "
            "Backups_Facturas_Produccion."
        )

    # GET de validacion del drive.
    drive = graph_mod.validar_drive()

    return graph_mod, drive


def archivos_locales_revalidables(
    candidato: dict[str, Any],
) -> list[Path]:
    cierre_dir = candidato["cierre_dir"]

    if not cierre_dir.is_dir():
        raise ErrorValidacion(
            f"No existe cierre local: {cierre_dir}"
        )

    archivos = sorted(
        p
        for p in cierre_dir.rglob("*")
        if p.is_file()
        and p.suffix.lower()
        not in {".tmp", ".lock", ".pyc"}
    )

    if not archivos:
        raise ErrorValidacion(
            "El cierre no contiene archivos para revalidar."
        )

    return archivos


def revalidar_vivo_con_graph(
    candidato: dict[str, Any],
    graph_mod,
) -> tuple[bool, list[str], int]:
    """
    Tercera validacion obligatoria.

    Comprueba EN VIVO que cada archivo que se pretende
    eliminar localmente sigue existiendo en OneDrive y
    coincide con la copia local.

    Excel:
        comparacion por datos/celdas.

    Otros:
        tamaño + SHA256 binario.

    NO contiene DELETE, PUT ni POST.
    """
    errores: list[str] = []

    remota = leer_json(
        candidato["remota_path"],
        "validacion remota almacenada",
    )

    estado = candidato["estado"]

    remote_base = str(
        remota.get("remote_base") or ""
    ).strip("/")

    if not remote_base:
        return False, [
            "Validacion remota sin remote_base."
        ], 0

    if not remote_base.startswith(
        "Backups_Facturas_Produccion/"
    ):
        return False, [
            "remote_base fuera del repositorio oficial."
        ], 0

    remote_estado = str(
        estado.get("remote_base") or ""
    ).strip("/")

    if remote_estado != remote_base:
        return False, [
            "remote_base del estado de retencion "
            "no coincide con la evidencia remota."
        ], 0

    archivos = archivos_locales_revalidables(
        candidato
    )

    cierre_dir = candidato["cierre_dir"]

    # Diario/Semanal/Mensual:
    # remote_base ya apunta al cierre concreto.
    #
    # Trimestral:
    # remote_base apunta al trimestre y la carpeta
    # local objetivo es .../Trimestral.
    if candidato["nivel"] == "TRIMESTRAL":
        base_relativa = cierre_dir.parent
    else:
        base_relativa = cierre_dir

    verificados = 0

    for archivo in archivos:
        try:
            if not dentro_de(
                archivo,
                cierre_dir,
            ):
                raise ErrorValidacion(
                    f"Archivo fuera del cierre: {archivo}"
                )

            relativa = archivo.relative_to(
                base_relativa
            ).as_posix()

            remote_path = (
                f"{remote_base}/{relativa}"
            ).strip("/")

            item = graph_mod.obtener_item_por_path(
                graph_mod.BACKUP_DRIVE_ID,
                remote_path,
            )

            if item is None:
                errores.append(
                    f"NO_EXISTE_REMOTO: {relativa}"
                )
                continue

            iguales, motivo = (
                graph_mod.comparar_local_remoto(
                    archivo,
                    relativa,
                    item,
                )
            )

            if not iguales:
                errores.append(
                    f"INTEGRIDAD_REMOTA: "
                    f"{relativa} -> {motivo}"
                )
                continue

            verificados += 1

        except Exception as exc:
            errores.append(
                f"{archivo.name}: "
                f"{type(exc).__name__}: {exc}"
            )

    return (
        len(errores) == 0
        and verificados == len(archivos),
        errores,
        verificados,
    )


def validar_objetivo_eliminacion_local(
    cierre_dir: Path,
    nivel: str,
) -> Path:
    """
    Permite eliminar solamente un cierre con la jerarquia
    oficial exacta bajo BASE.

    Tambien rechaza symlinks en cualquier componente.
    """
    base_lexica = BASE.absolute()
    objetivo_lexico = cierre_dir.absolute()

    try:
        relativa = objetivo_lexico.relative_to(
            base_lexica
        )
    except ValueError as exc:
        raise ErrorValidacion(
            "Objetivo de eliminacion fuera de BASE."
        ) from exc

    if not relativa.parts:
        raise ErrorValidacion(
            "Se intento eliminar la carpeta BASE completa."
        )

    actual = base_lexica

    for parte in relativa.parts:
        actual = actual / parte

        if actual.is_symlink():
            raise ErrorValidacion(
                f"Ruta de eliminacion contiene symlink: {actual}"
            )

    objetivo = objetivo_lexico.resolve()
    base = BASE.resolve()

    if not objetivo.is_dir():
        raise ErrorValidacion(
            f"Objetivo local inexistente: {objetivo}"
        )

    if not dentro_de(objetivo, base):
        raise ErrorValidacion(
            "Objetivo resuelto fuera de BASE."
        )

    patrones = {
        "ANIO": r"^\d{4}$",
        "TRIMESTRE": (
            r"^TRIMESTRE_"
            r"\d{4}-\d{2}-\d{2}_A_"
            r"\d{4}-\d{2}-\d{2}$"
        ),
        "MES": (
            r"^\d{4}-\d{2}_"
            r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+$"
        ),
        "SEMANA": (
            r"^SEMANA_"
            r"\d{4}-\d{2}-\d{2}_a_"
            r"\d{4}-\d{2}-\d{2}$"
        ),
        "DIARIO": r"^Diario_\d{4}-\d{2}-\d{2}$",
    }

    estructuras = {
        "DIARIO": (
            5,
            {
                0: patrones["ANIO"],
                1: patrones["TRIMESTRE"],
                2: patrones["MES"],
                3: patrones["SEMANA"],
                4: patrones["DIARIO"],
            },
        ),
        "SEMANAL": (
            5,
            {
                0: patrones["ANIO"],
                1: patrones["TRIMESTRE"],
                2: patrones["MES"],
                3: patrones["SEMANA"],
                4: r"^Semanal$",
            },
        ),
        "MENSUAL": (
            4,
            {
                0: patrones["ANIO"],
                1: patrones["TRIMESTRE"],
                2: patrones["MES"],
                3: r"^Mensual$",
            },
        ),
        "TRIMESTRAL": (
            3,
            {
                0: patrones["ANIO"],
                1: patrones["TRIMESTRE"],
                2: r"^Trimestral$",
            },
        ),
    }

    if nivel not in estructuras:
        raise ErrorValidacion(
            f"Nivel no soportado para eliminacion: {nivel}"
        )

    total_esperado, reglas = estructuras[nivel]
    partes = relativa.parts

    if len(partes) != total_esperado:
        raise ErrorValidacion(
            f"Jerarquia invalida para {nivel}: "
            f"{relativa.as_posix()}"
        )

    for indice, patron in reglas.items():
        if not re.fullmatch(
            patron,
            partes[indice],
        ):
            raise ErrorValidacion(
                f"Componente de jerarquia invalido para "
                f"{nivel}: {partes[indice]}"
            )

    return objetivo

def eliminar_directorio_local_seguro(
    cierre_dir: Path,
    nivel: str,
) -> None:
    """
    Elimina UNICAMENTE la carpeta local del cierre.

    No contiene ni invoca ninguna operacion Graph.
    """
    objetivo = validar_objetivo_eliminacion_local(
        cierre_dir,
        nivel,
    )

    shutil.rmtree(objetivo)

    if objetivo.exists():
        raise ErrorValidacion(
            f"La eliminacion local no se completo: {objetivo}"
        )


def prueba_borrado_sandbox() -> None:
    """
    Prueba la MISMA funcion de borrado que usaria un cierre,
    pero sobre una jerarquia artificial del año 2099.
    """
    base_prueba = BASE / "2099"

    if base_prueba.exists():
        raise ErrorValidacion(
            f"Sandbox ya existia antes de la prueba: {base_prueba}"
        )

    diario = (
        base_prueba
        / "TRIMESTRE_2099-01-01_A_2099-03-31"
        / "2099-01_Enero"
        / "SEMANA_2099-01-01_a_2099-01-07"
        / "Diario_2099-01-01"
    )

    try:
        diario.mkdir(parents=True)

        (diario / "archivo_prueba.txt").write_text(
            "PRUEBA BORRADO LOCAL AISLADO\n",
            encoding="utf-8",
        )

        eliminar_directorio_local_seguro(
            diario,
            "DIARIO",
        )

        if diario.exists():
            raise ErrorValidacion(
                "La funcion real no elimino el sandbox."
            )

        # Prueba negativa: un Diario directamente bajo el
        # año 2099 debe ser rechazado por jerarquia invalida.
        invalido = (
            base_prueba
            / "Diario_2099-01-02"
        )

        invalido.mkdir(parents=True)

        rechazo_ok = False

        try:
            eliminar_directorio_local_seguro(
                invalido,
                "DIARIO",
            )
        except ErrorValidacion:
            rechazo_ok = True

        if not rechazo_ok:
            raise ErrorValidacion(
                "La proteccion de jerarquia NO rechazo "
                "un objetivo invalido."
            )

    finally:
        shutil.rmtree(
            base_prueba,
            ignore_errors=True,
        )



def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Diagnostica candidatos de limpieza "
            "local por politica de retencion."
        )
    )

    parser.add_argument(
        "--nivel",
        choices=[
            "TODOS",
            "DIARIO",
            "SEMANAL",
            "MENSUAL",
            "TRIMESTRAL",
        ],
        default="TODOS",
    )

    parser.add_argument(
        "--ahora",
        help=(
            "Fecha/hora ISO opcional para pruebas "
            "de diagnostico. No elimina nada."
        ),
    )

    parser.add_argument(
        "--validar-vivo",
        action="store_true",
        help=(
            "Ejecuta la tercera validacion contra "
            "OneDrive usando exclusivamente GET/descarga. "
            "No elimina nada."
        ),
    )

    parser.add_argument(
        "--identidad",
        help=(
            "Limita la evaluacion a una identidad exacta, "
            "por ejemplo 2026-08-13."
        ),
    )

    parser.add_argument(
        "--ejecutar",
        action="store_true",
        help=(
            "Ejecuta la eliminacion exclusivamente local "
            "despues de superar todas las validaciones y "
            "barreras de seguridad."
        ),
    )

    parser.add_argument(
        "--confirmar",
        help=(
            "Confirmacion explicita requerida para la "
            "eliminacion local."
        ),
    )

    parser.add_argument(
        "--probar-borrado-sandbox",
        action="store_true",
        help=(
            "Prueba exclusivamente el borrado de una carpeta "
            "desechable creada bajo una jerarquia artificial y temporal bajo BASE/2099."
        ),
    )

    args = parser.parse_args()

    if args.probar_borrado_sandbox:
        if args.ejecutar:
            parser.error(
                "--probar-borrado-sandbox no puede combinarse "
                "con --ejecutar."
            )

        prueba_borrado_sandbox()

        print(
            "RESULTADO = "
            "PRUEBA_BORRADO_LOCAL_SANDBOX_OK"
        )
        return 0

    if args.ejecutar:
        if args.ahora:
            parser.error(
                "--ahora es exclusivamente de simulacion y "
                "NUNCA puede usarse junto con --ejecutar."
            )

        if args.confirmar != CONFIRMACION_EJECUCION:
            parser.error(
                "--ejecutar requiere --confirmar "
                f"{CONFIRMACION_EJECUCION}"
            )

        if os.getenv(
            ENV_LIMPIEZA_HABILITADA,
            "",
        ).strip().upper() != "SI":
            parser.error(
                "La eliminacion local esta bloqueada: "
                f"{ENV_LIMPIEZA_HABILITADA}=SI no esta habilitado."
            )

        if os.getenv(
            ENV_ALERTAS_HABILITADAS,
            "",
        ).strip().upper() != "SI":
            parser.error(
                "La eliminacion local esta bloqueada porque "
                "las alertas todavia no estan habilitadas: "
                f"{ENV_ALERTAS_HABILITADAS}=SI."
            )

        if not args.validar_vivo:
            parser.error(
                "--ejecutar requiere obligatoriamente "
                "--validar-vivo."
            )

        if args.nivel == "TODOS":
            parser.error(
                "La ejecucion no permite modo masivo con "
                "--nivel TODOS."
            )

        if not args.identidad:
            parser.error(
                "--ejecutar requiere una --identidad exacta."
            )

    ahora = (
        parse_datetime(args.ahora)
        if args.ahora
        else dt.datetime.now(
            ZONA_HORARIA
        )
    )

    niveles = (
        list(RETENCIONES)
        if args.nivel == "TODOS"
        else [args.nivel]
    )

    resultados: list[dict[str, Any]] = []

    graph_mod = None
    drive_graph = None

    print("=" * 100)
    print(
        "LIMPIEZA LOCAL DE BACKUPS - "
        + (
            "EJECUCION CONTROLADA"
            if args.ejecutar
            else "SOLO DIAGNOSTICO"
        )
    )
    print("=" * 100)
    print(f"Version: {VERSION}")
    print(f"Root: {ROOT}")
    print(f"Fecha evaluacion: {ahora.isoformat(timespec='seconds')}")
    print(
        "Graph: "
        + (
            "SOLO LECTURA - GET/DOWNLOAD"
            if args.validar_vivo
            else "NO"
        )
    )
    print(
        "Eliminaciones locales: "
        + (
            "HABILITADAS PARA UN CIERRE VALIDADO"
            if args.ejecutar
            else "NO"
        )
    )
    print("Eliminaciones remotas: IMPOSIBLES EN ESTA VERSION")
    print("-" * 100)

    for nivel in niveles:
        estados = descubrir_estados(nivel)

        print(
            f"{nivel}: estados detectados = "
            f"{len(estados)}"
        )

        for estado_path in estados:
            try:
                candidato = datos_candidato(
                    nivel,
                    estado_path,
                )

                if (
                    args.identidad
                    and candidato["identidad"]
                    != args.identidad
                ):
                    continue

                resultado = evaluar_candidato(
                    candidato,
                    ahora,
                )

                resultado[
                    "ok_validacion_viva"
                ] = None

                resultado[
                    "archivos_revalidados_vivo"
                ] = 0

                if (
                    args.validar_vivo
                    and resultado.get(
                        "preaprobado_para_validacion_viva"
                    )
                ):
                    try:
                        if graph_mod is None:
                            (
                                graph_mod,
                                drive_graph,
                            ) = (
                                cargar_modulo_graph_solo_lectura()
                            )

                            print(
                                "Drive Graph validado en "
                                "modo solo lectura: "
                                f"{drive_graph.get('name')}"
                            )

                        (
                            ok_viva,
                            errores_vivos,
                            total_vivo,
                        ) = revalidar_vivo_con_graph(
                            candidato,
                            graph_mod,
                        )

                        resultado[
                            "ok_validacion_viva"
                        ] = ok_viva

                        resultado[
                            "archivos_revalidados_vivo"
                        ] = total_vivo

                        if errores_vivos:
                            resultado[
                                "errores"
                            ].extend(
                                f"VALIDACION_VIVA: {x}"
                                for x in errores_vivos
                            )

                    except Exception as exc:
                        resultado[
                            "ok_validacion_viva"
                        ] = False
                        resultado[
                            "errores"
                        ].append(
                            "VALIDACION_VIVA: "
                            f"{type(exc).__name__}: {exc}"
                        )

                resultado[
                    "aprobado_para_eliminacion"
                ] = bool(
                    resultado.get(
                        "preaprobado_para_validacion_viva"
                    )
                    and resultado.get(
                        "ok_validacion_viva"
                    ) is True
                )

            except Exception as exc:
                resultado = {
                    "nivel": nivel,
                    "identidad": estado_path.name,
                    "preaprobado_para_validacion_viva": False,
                    "errores": [
                        f"{type(exc).__name__}: {exc}"
                    ],
                }

            resultados.append(resultado)

    print("-" * 100)

    preaprobados = [
        x for x in resultados
        if x.get(
            "preaprobado_para_validacion_viva"
        )
    ]

    con_errores = [
        x for x in resultados
        if x.get("errores")
    ]

    validaciones_vivas_ok = [
        x for x in resultados
        if x.get("ok_validacion_viva") is True
    ]

    aprobados_sin_eliminar = [
        x for x in resultados
        if x.get("aprobado_para_eliminacion") is True
    ]

    no_vencidos = [
        x for x in resultados
        if (
            x.get("retencion_vencida") is False
            and not x.get("errores")
        )
    ]

    print(
        f"CIERRES_EVALUADOS                 = "
        f"{len(resultados)}"
    )
    print(
        f"PREAPROBADOS_VALIDACION_VIVA      = "
        f"{len(preaprobados)}"
    )
    print(
        f"RETENCION_AUN_NO_VENCIDA          = "
        f"{len(no_vencidos)}"
    )
    print(
        f"CIERRES_CON_ERROR                 = "
        f"{len(con_errores)}"
    )
    print(
        f"VALIDACIONES_VIVAS_OK             = "
        f"{len(validaciones_vivas_ok)}"
    )
    etiqueta_aprobados = (
        "APROBADOS_PARA_ELIMINACION"
        if args.ejecutar
        else "APROBADOS_SIN_ELIMINAR"
    )

    print(
        f"{etiqueta_aprobados:<36} = "
        f"{len(aprobados_sin_eliminar)}"
    )

    if args.ejecutar:
        aprobados = [
            x
            for x in resultados
            if x.get("aprobado_para_eliminacion") is True
        ]

        if len(resultados) != 1:
            parser.error(
                "--ejecutar requiere que la identidad exacta "
                "resuelva un unico cierre."
            )

        if len(aprobados) != 1:
            parser.error(
                "El cierre solicitado NO supero todas las "
                "validaciones. No se elimino nada."
            )

        aprobado = aprobados[0]

        eliminar_directorio_local_seguro(
            Path(aprobado["cierre_dir"]),
            aprobado["nivel"],
        )

        aprobado["eliminado_localmente"] = True

        print()
        print(
            "ELIMINACION_LOCAL_COMPLETADA = "
            f"{aprobado['nivel']} "
            f"{aprobado['identidad']}"
        )
        print(
            "RUTA_ELIMINADA = "
            f"{aprobado['cierre_dir']}"
        )

    if con_errores:
        print()
        print("ERRORES:")
        for item in con_errores:
            print(
                f"- {item.get('nivel')} "
                f"{item.get('identidad')}"
            )
            for error in item.get(
                "errores",
                [],
            ):
                print(f"    * {error}")

    print()
    print(
        "RESULTADO = "
        + (
            "LIMPIEZA_RETENCION_LOCAL_COMPLETADA"
            if args.ejecutar
            else "DIAGNOSTICO_RETENCION_LOCAL_COMPLETADO"
        )
    )
    print("=" * 100)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
