# utils/processed_store.py
import json
import os
import time
from typing import Dict, Any, Optional


def _safe_write_json(path: str, data: Dict[str, Any]) -> None:
    folder = os.path.dirname(path)
    if folder:
        os.makedirs(folder, exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


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
