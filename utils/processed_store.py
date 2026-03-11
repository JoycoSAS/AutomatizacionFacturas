# utils/processed_store.py
import json
import os
import time
import random
from typing import Dict, Any, Optional


def _safe_write_json(path: str, data: Dict[str, Any], retries: int = 8, base_sleep: float = 0.12) -> None:
    """
    Escritura robusta en Windows:
    - Escribe a un archivo temporal único en el MISMO directorio.
    - flush + fsync para asegurar escritura.
    - os.replace con reintentos (por locks/antivirus/OneDrive).
    - Fallback: si no se puede reemplazar, intenta escribir directo (último recurso).
    """
    folder = os.path.dirname(path)
    if folder:
        os.makedirs(folder, exist_ok=True)

    # temp único (misma carpeta) para mantener atomicidad del replace
    uniq = f"{int(time.time() * 1000)}_{os.getpid()}_{random.randint(1000, 9999)}"
    tmp = f"{path}.{uniq}.tmp"

    # 1) escribir temp
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.flush()
        try:
            os.fsync(f.fileno())
        except Exception:
            # en algunos FS puede no estar disponible; no es crítico
            pass

    # 2) intentar replace con reintentos (Windows a veces bloquea)
    last_err: Optional[Exception] = None
    for i in range(retries):
        try:
            os.replace(tmp, path)  # atomic-ish en Windows si no está bloqueado
            return
        except PermissionError as e:
            last_err = e
            # Espera exponencial con jitter (evita colisiones)
            time.sleep(base_sleep * (2 ** i) + random.random() * 0.05)
        except OSError as e:
            last_err = e
            time.sleep(base_sleep * (2 ** i) + random.random() * 0.05)

    # 3) si no se pudo, intentar fallback:
    #    - borrar tmp si quedó
    #    - escribir directo sobre el destino (no atómico, pero evita crashear el proceso)
    try:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass
    except Exception:
        pass

    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            try:
                os.fsync(f.fileno())
            except Exception:
                pass
        return
    except Exception as e2:
        # Si también falla, re-lanzamos el error original para diagnóstico
        raise last_err or e2


class ProcessedStore:
    """
    Guarda estado local de mensajes procesados (Graph messageId)
    para evitar reprocesos aunque no puedas marcar como leído (403).

    Estructura:
    {
      "version": 1,
      "updated_at": 1234567890,
      "items": {
        "<message_id>": {"ts": 1234567890, "meta": {...}}
      }
    }
    """

    def __init__(self, path: str, ttl_days: int = 30):
        self.path = path
        self.ttl_seconds = int(ttl_days) * 24 * 3600

    def _load(self) -> Dict[str, Any]:
        if not os.path.exists(self.path):
            return {"version": 1, "updated_at": int(time.time()), "items": {}}
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)

            if "items" not in data or not isinstance(data["items"], dict):
                data["items"] = {}
            if "updated_at" not in data:
                data["updated_at"] = int(time.time())
            if "version" not in data:
                data["version"] = 1
            return data
        except Exception:
            # Si el JSON se dañó por algo raro, lo reiniciamos (mejor que crashear)
            return {"version": 1, "updated_at": int(time.time()), "items": {}}

    def _prune(self, data: Dict[str, Any]) -> Dict[str, Any]:
        now = int(time.time())
        items = data.get("items", {})
        if not items:
            data["updated_at"] = now
            return data

        to_del = []
        for mid, obj in items.items():
            try:
                ts = int(obj.get("ts", 0) or 0)
            except Exception:
                ts = 0
            if ts and (now - ts) > self.ttl_seconds:
                to_del.append(mid)

        for mid in to_del:
            items.pop(mid, None)

        data["items"] = items
        data["updated_at"] = now
        return data

    def is_processed(self, message_id: str) -> bool:
        if not message_id:
            return False
        data = self._prune(self._load())
        return message_id in data.get("items", {})

    def mark_processed(self, message_id: str, meta: Optional[Dict[str, Any]] = None) -> None:
        if not message_id:
            return
        data = self._prune(self._load())

        data["items"][message_id] = {
            "ts": int(time.time()),
            "meta": meta or {},
        }
        data["updated_at"] = int(time.time())

        _safe_write_json(self.path, data)