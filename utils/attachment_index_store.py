# utils/attachment_index_store.py
import json
import os
import time
import datetime
from typing import Any, Dict, List, Optional, Tuple

from utils.text_normalizer import normalize_text


def _utc_now_ts() -> int:
    return int(time.time())


def _parse_iso_to_ts(iso_str: str) -> Optional[int]:
    """
    Convierte '2025-11-24T15:30:10Z' o con offset a epoch ts.
    """
    if not iso_str:
        return None
    s = iso_str.strip()
    try:
        # Graph suele traer Z
        if s.endswith("Z"):
            dt = datetime.datetime.fromisoformat(s.replace("Z", "+00:00"))
        else:
            dt = datetime.datetime.fromisoformat(s)
        return int(dt.timestamp())
    except Exception:
        return None


def _safe_read_json(path: str) -> Dict[str, Any]:
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def _safe_write_json(path: str, data: Dict[str, Any]) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _norm_cufe(cufe: str) -> str:
    if not cufe:
        return ""
    # cufe ya viene en hex normalmente; pero por si acaso:
    s = normalize_text(cufe)
    # dejamos solo 0-9a-f
    out = []
    for ch in s:
        if ch in "0123456789abcdef":
            out.append(ch)
    return "".join(out)


def _norm_num(n: str) -> str:
    """
    Normalización de número de factura para usarlo como clave:
    - sin espacios
    - mayúsculas
    - deja letras/números y separadores comunes (- / .)
    - también guardamos una "compacta" sin separadores para búsquedas flexibles
    """
    if not n:
        return ""
    s = str(n).strip().upper()
    s = s.replace("–", "-").replace("—", "-")
    # colapsa espacios
    s = " ".join(s.split())
    return s


def _num_variants(n: str) -> List[str]:
    s = _norm_num(n)
    if not s:
        return []
    compact = "".join(ch for ch in s if ch.isalnum())
    # variant con separadores normalizados (sin espacios alrededor)
    s2 = s.replace(" ", "")
    out = [s2]
    if compact and compact != s2:
        out.append(compact)
    # sin puntos
    no_dots = s2.replace(".", "")
    if no_dots != s2:
        out.append(no_dots)
    # sin guiones
    no_dash = s2.replace("-", "")
    if no_dash != s2:
        out.append(no_dash)
    # únicos manteniendo orden
    seen = set()
    uniq = []
    for x in out:
        if x and x not in seen:
            seen.add(x)
            uniq.append(x)
    return uniq


def _norm_fecha(fecha: str) -> str:
    """
    Espera fechas ya normalizadas a 'YYYY-MM-DD' en tu flujo.
    Si viene 'YYYY-MM', la deja igual. Si viene vacía, ''.
    """
    if not fecha:
        return ""
    return str(fecha).strip()


class AttachmentIndexStore:
    """
    Un solo JSON con secciones:
      - zip_by_cufe: { cufe: entry }
      - zip_by_nf: { "NUM|FECHA": [entry, entry, ...] }
      - conta15_by_cufe: { cufe: entry }
    """

    def __init__(
        self,
        path: str,
        ttl_days: int = 365,
        max_nf_per_key: int = 10,
    ):
        self.path = path
        self.ttl_days = max(1, int(ttl_days))
        self.max_nf_per_key = max(1, int(max_nf_per_key))
        self._data = _safe_read_json(self.path)
        self._ensure_shape()

    def _ensure_shape(self) -> None:
        if not isinstance(self._data, dict):
            self._data = {}
        self._data.setdefault("meta", {})
        self._data.setdefault("zip_by_cufe", {})
        self._data.setdefault("zip_by_nf", {})
        self._data.setdefault("conta15_by_cufe", {})

    def _save(self) -> None:
        _safe_write_json(self.path, self._data)

    def purge(self) -> int:
        """
        Borra entradas viejas basado en ts (received_ts o added_ts).
        Retorna cantidad borrada.
        """
        cutoff = _utc_now_ts() - int(self.ttl_days * 86400)
        removed = 0

        def is_old(entry: Dict[str, Any]) -> bool:
            ts = entry.get("received_ts") or entry.get("added_ts") or 0
            try:
                ts = int(ts)
            except Exception:
                ts = 0
            return ts and ts < cutoff

        # zip_by_cufe
        zc = self._data.get("zip_by_cufe", {})
        for k in list(zc.keys()):
            if is_old(zc.get(k, {})):
                del zc[k]
                removed += 1

        # zip_by_nf (listas)
        znf = self._data.get("zip_by_nf", {})
        for k in list(znf.keys()):
            arr = znf.get(k) or []
            if not isinstance(arr, list):
                arr = []
            new_arr = [e for e in arr if not is_old(e)]
            if not new_arr:
                if k in znf:
                    del znf[k]
                removed += 1
            else:
                znf[k] = new_arr

        # conta15_by_cufe
        c15 = self._data.get("conta15_by_cufe", {})
        for k in list(c15.keys()):
            if is_old(c15.get(k, {})):
                del c15[k]
                removed += 1

        if removed:
            self._save()
        return removed

    # -------------------
    # ZIP (XML) indexing
    # -------------------

    def upsert_zip(
        self,
        cufe: str,
        numero: str,
        fecha: str,
        msg_id: str,
        att_id: str,
        att_name: str,
        received_dt_iso: str = "",
    ) -> None:
        cufe_n = _norm_cufe(cufe)
        fecha_n = _norm_fecha(fecha)
        numero_vars = _num_variants(numero)

        entry = {
            "msg_id": msg_id,
            "att_id": att_id,
            "att_name": att_name or "",
            "received_dt": received_dt_iso or "",
            "received_ts": _parse_iso_to_ts(received_dt_iso) or 0,
            "added_ts": _utc_now_ts(),
            "cufe": cufe_n,
            "numero": _norm_num(numero),
            "fecha": fecha_n,
        }

        # por CUFE (1 a 1)
        if cufe_n:
            self._data["zip_by_cufe"][cufe_n] = entry

        # por NUM+FECHA (pueden existir varios: guardamos lista corta)
        if fecha_n and numero_vars:
            for nv in numero_vars:
                key = f"{nv}|{fecha_n}"
                arr = self._data["zip_by_nf"].get(key)
                if not isinstance(arr, list):
                    arr = []

                # evita duplicar mismo msg_id+att_id
                exists = any((e.get("msg_id") == msg_id and e.get("att_id") == att_id) for e in arr)
                if not exists:
                    arr.insert(0, entry)  # lo más reciente primero

                # recorta
                if len(arr) > self.max_nf_per_key:
                    arr = arr[: self.max_nf_per_key]

                self._data["zip_by_nf"][key] = arr

        self._save()

    def find_zip_by_cufe(self, cufe: str) -> Optional[Dict[str, Any]]:
        cufe_n = _norm_cufe(cufe)
        if not cufe_n:
            return None
        return self._data.get("zip_by_cufe", {}).get(cufe_n)

    def find_zip_by_num_fecha(self, numero: str, fecha: str) -> Optional[Dict[str, Any]]:
        fecha_n = _norm_fecha(fecha)
        if not fecha_n:
            return None
        for nv in _num_variants(numero):
            key = f"{nv}|{fecha_n}"
            arr = self._data.get("zip_by_nf", {}).get(key)
            if isinstance(arr, list) and arr:
                # devolver el más reciente (arr[0])
                return arr[0]
        return None

    # -------------------------
    # CONTA15 (PDF DIAN) index
    # -------------------------

    def upsert_conta15_pdf(
        self,
        cufe: str,
        msg_id: str,
        att_id: str,
        att_name: str,
        subject: str = "",
        received_dt_iso: str = "",
    ) -> None:
        cufe_n = _norm_cufe(cufe)
        if not cufe_n:
            return

        entry = {
            "msg_id": msg_id,
            "att_id": att_id,
            "att_name": att_name or "",
            "subject": subject or "",
            "subject_norm": normalize_text(subject),
            "received_dt": received_dt_iso or "",
            "received_ts": _parse_iso_to_ts(received_dt_iso) or 0,
            "added_ts": _utc_now_ts(),
            "cufe": cufe_n,
        }

        self._data["conta15_by_cufe"][cufe_n] = entry
        self._save()

    def find_conta15_by_cufe(self, cufe: str) -> Optional[Dict[str, Any]]:
        cufe_n = _norm_cufe(cufe)
        if not cufe_n:
            return None
        return self._data.get("conta15_by_cufe", {}).get(cufe_n)
