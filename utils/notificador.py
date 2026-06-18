# -*- coding: utf-8 -*-
"""
JOYCO - Facturas Procesador
Módulo central de notificaciones y alertas.

Política definida:
- Procesos frecuentes: notificar solo fallas.
- Backup mensual: notificar éxito y falla.
- Cierre trimestral / rotación: notificar éxito y falla.
- No notificar cuando no hay facturas nuevas.
"""

from __future__ import annotations

import datetime
import json
import os
import sys
import traceback
from html import escape
from pathlib import Path
from typing import Any, Optional
from urllib.parse import quote

import requests
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

load_dotenv(ROOT / ".env")

VERSION_NOTIFICADOR = "2026-06-18-NOTIFICADOR-V1-GRAPH-EMAIL-CONTROLADO"
GRAPH = "https://graph.microsoft.com/v1.0"

ESTADOS_VALIDOS = {"OK", "ERROR", "INFO", "WARNING"}

PROCESOS_EXITO_INFORMATIVO = {
    "BACKUP_MENSUAL_LOCAL",
    "BACKUP_MENSUAL_SHAREPOINT",
    "CIERRE_TRIMESTRAL_LOCAL",
    "CIERRE_TRIMESTRAL_SHAREPOINT",
    "REEMPLAZO_EXCEL_ACTIVO_TRIMESTRAL",
}


def _bool_env(nombre: str, default: str = "0") -> bool:
    return str(os.getenv(nombre, default) or default).strip().lower() in {
        "1",
        "true",
        "yes",
        "y",
        "si",
        "sí",
        "on",
    }


def ssl_verify() -> bool:
    return not str(os.getenv("SSL_VERIFY", "true") or "true").strip().lower() in {
        "0",
        "false",
        "no",
        "off",
    }


def alertas_habilitadas() -> bool:
    return _bool_env("ALERTAS_HABILITADAS", "0")


def canal_alertas() -> str:
    return str(os.getenv("ALERTAS_CANAL", "CONSOLE") or "CONSOLE").strip().upper()


def nombre_sistema() -> str:
    return str(os.getenv("ALERTAS_NOMBRE_SISTEMA", "Automatización Facturas JOYCO") or "").strip()


def remitente_alertas() -> str:
    return str(os.getenv("ALERTAS_REMITENTE", "") or "").strip()


def destinatarios_alertas() -> list[str]:
    raw = str(os.getenv("ALERTAS_DESTINATARIOS", "") or "").strip()

    if not raw:
        return []

    partes = raw.replace(";", ",").replace("\n", ",").split(",")
    return [p.strip() for p in partes if p.strip()]


def _timestamp() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _normalizar_estado(estado: str) -> str:
    estado = str(estado or "").strip().upper()
    if estado not in ESTADOS_VALIDOS:
        return "INFO"
    return estado


def _prefijo_estado(estado: str) -> str:
    estado = _normalizar_estado(estado)

    if estado == "OK":
        return "✅ OK"
    if estado == "ERROR":
        return "❌ ERROR"
    if estado == "WARNING":
        return "⚠️ WARNING"
    return "ℹ️ INFO"


def _serializar_detalle(detalle: Any) -> str:
    if detalle is None:
        return ""

    if isinstance(detalle, str):
        return detalle.strip()

    try:
        return json.dumps(detalle, ensure_ascii=False, indent=2, default=str)
    except Exception:
        return str(detalle)


def _asunto(estado: str, proceso: str, asunto: str) -> str:
    estado = _normalizar_estado(estado)
    sistema = nombre_sistema()

    proceso = str(proceso or "PROCESO").strip().upper()
    asunto = str(asunto or "").strip()

    return f"[{sistema}] [{estado}] {proceso} - {asunto}"


def _contenido_texto(
    *,
    estado: str,
    proceso: str,
    asunto: str,
    mensaje: str,
    detalle: Any = None,
) -> str:
    detalle_txt = _serializar_detalle(detalle)

    partes = [
        f"{_prefijo_estado(estado)}",
        "",
        f"Sistema: {nombre_sistema()}",
        f"Proceso: {proceso}",
        f"Asunto: {asunto}",
        f"Fecha/hora: {_timestamp()}",
        "",
        "Mensaje:",
        str(mensaje or "").strip(),
    ]

    if detalle_txt:
        partes.extend(["", "Detalle técnico:", detalle_txt])

    return "\n".join(partes).strip()


def _contenido_html(
    *,
    estado: str,
    proceso: str,
    asunto: str,
    mensaje: str,
    detalle: Any = None,
) -> str:
    detalle_txt = _serializar_detalle(detalle)

    color = "#198754"
    if _normalizar_estado(estado) == "ERROR":
        color = "#dc3545"
    elif _normalizar_estado(estado) == "WARNING":
        color = "#fd7e14"

    html = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #222;">
        <h2 style="color: {color}; margin-bottom: 8px;">{escape(_prefijo_estado(estado))}</h2>
        <table style="border-collapse: collapse; margin-bottom: 16px;">
            <tr><td><b>Sistema:</b></td><td>{escape(nombre_sistema())}</td></tr>
            <tr><td><b>Proceso:</b></td><td>{escape(str(proceso))}</td></tr>
            <tr><td><b>Asunto:</b></td><td>{escape(str(asunto))}</td></tr>
            <tr><td><b>Fecha/hora:</b></td><td>{escape(_timestamp())}</td></tr>
        </table>

        <h3>Mensaje</h3>
        <p>{escape(str(mensaje or "")).replace(chr(10), "<br>")}</p>
    """

    if detalle_txt:
        html += f"""
        <h3>Detalle técnico</h3>
        <pre style="background:#f6f6f6; padding:12px; border:1px solid #ddd; white-space:pre-wrap;">
{escape(detalle_txt)}
        </pre>
        """

    html += """
    </body>
    </html>
    """

    return html.strip()


def _get_access_token() -> str:
    from services.m365.token import get_access_token

    return get_access_token()


def _enviar_graph_email(
    *,
    estado: str,
    proceso: str,
    asunto: str,
    mensaje: str,
    detalle: Any = None,
) -> dict:
    remitente = remitente_alertas()
    destinatarios = destinatarios_alertas()

    if not remitente:
        raise RuntimeError("Falta ALERTAS_REMITENTE en .env.")

    if not destinatarios:
        raise RuntimeError("Falta ALERTAS_DESTINATARIOS en .env.")

    token = _get_access_token()

    subject = _asunto(estado, proceso, asunto)
    html = _contenido_html(
        estado=estado,
        proceso=proceso,
        asunto=asunto,
        mensaje=mensaje,
        detalle=detalle,
    )

    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": html,
            },
            "toRecipients": [
                {"emailAddress": {"address": correo}}
                for correo in destinatarios
            ],
        },
        "saveToSentItems": True,
    }

    url = f"{GRAPH}/users/{quote(remitente, safe='')}/sendMail"

    r = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=60,
        verify=ssl_verify(),
    )

    if r.status_code not in (202,):
        raise RuntimeError(f"Graph sendMail falló: {r.status_code} -> {r.text[:800]}")

    return {
        "enviado": True,
        "canal": "GRAPH_EMAIL",
        "remitente": remitente,
        "destinatarios": destinatarios,
        "status_code": r.status_code,
        "subject": subject,
    }


def enviar_notificacion(
    *,
    proceso: str,
    estado: str,
    asunto: str,
    mensaje: str,
    detalle: Any = None,
    dry_run: bool = False,
    forzar: bool = False,
) -> dict:
    """
    Envía o simula una notificación.

    Si ALERTAS_HABILITADAS=0 y no es dry_run ni forzar, no envía.
    """

    estado = _normalizar_estado(estado)
    proceso = str(proceso or "PROCESO").strip().upper()
    asunto = str(asunto or "").strip()
    mensaje = str(mensaje or "").strip()

    contenido = _contenido_texto(
        estado=estado,
        proceso=proceso,
        asunto=asunto,
        mensaje=mensaje,
        detalle=detalle,
    )

    if dry_run:
        print("=" * 100)
        print("DRY RUN NOTIFICACIÓN - NO SE ENVÍA CORREO")
        print("=" * 100)
        print(contenido)
        print("=" * 100)
        return {
            "enviado": False,
            "dry_run": True,
            "canal": canal_alertas(),
            "proceso": proceso,
            "estado": estado,
        }

    if not alertas_habilitadas() and not forzar:
        print("🔕 Alertas deshabilitadas por ALERTAS_HABILITADAS=0.")
        return {
            "enviado": False,
            "motivo": "ALERTAS_DESHABILITADAS",
            "proceso": proceso,
            "estado": estado,
        }

    canal = canal_alertas()

    if canal in {"CONSOLE", "LOG", "NONE"}:
        print("=" * 100)
        print(f"NOTIFICACIÓN {canal} - NO SE ENVÍA CORREO")
        print("=" * 100)
        print(contenido)
        print("=" * 100)
        return {
            "enviado": False,
            "canal": canal,
            "proceso": proceso,
            "estado": estado,
        }

    if canal == "GRAPH_EMAIL":
        return _enviar_graph_email(
            estado=estado,
            proceso=proceso,
            asunto=asunto,
            mensaje=mensaje,
            detalle=detalle,
        )

    raise RuntimeError(f"ALERTAS_CANAL no soportado: {canal}")


def notificar_fallo(
    *,
    proceso: str,
    asunto: str,
    mensaje: str,
    detalle: Any = None,
    dry_run: bool = False,
) -> dict:
    return enviar_notificacion(
        proceso=proceso,
        estado="ERROR",
        asunto=asunto,
        mensaje=mensaje,
        detalle=detalle,
        dry_run=dry_run,
    )


def notificar_exito_control(
    *,
    proceso: str,
    asunto: str,
    mensaje: str,
    detalle: Any = None,
    dry_run: bool = False,
) -> dict:
    proceso_norm = str(proceso or "").strip().upper()

    if proceso_norm not in PROCESOS_EXITO_INFORMATIVO:
        return {
            "enviado": False,
            "motivo": "EXITO_NO_INFORMATIVO_POR_POLITICA",
            "proceso": proceso_norm,
        }

    return enviar_notificacion(
        proceso=proceso_norm,
        estado="OK",
        asunto=asunto,
        mensaje=mensaje,
        detalle=detalle,
        dry_run=dry_run,
    )


def detalle_excepcion(exc: BaseException) -> dict:
    return {
        "tipo": type(exc).__name__,
        "mensaje": str(exc),
        "traceback": traceback.format_exc(limit=8),
    }
