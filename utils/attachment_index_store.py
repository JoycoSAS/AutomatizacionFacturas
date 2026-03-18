import json
import os
import time
import datetime
import re
from typing import Any, Dict, List, Optional

from utils.text_normalizer import normalize_text


def _utc_now_ts() -> int:
    return int(time.time())


def _parse_iso_to_ts(iso_str: str) -> Optional[int]:
    if not iso_str:
        return None

    s = str(iso_str).strip()

    try:
        if s.endswith("Z"):
            dt = datetime.datetime.fromisoformat(s.replace("Z", "+00:00"))
        else:
            dt = datetime.datetime.fromisoformat(s)

        return int(dt.timestamp())

    except Exception:
        return None


def _safe_read_json(path: str) -> Dict[str, Any]:

    if not path or not os.path.exists(path):
        return {}

    try:

        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

            if isinstance(data, dict):
                return data

    except Exception:
        pass

    return {}


def _safe_write_json(path: str, data: Dict[str, Any]):

    folder = os.path.dirname(path)

    if folder:
        os.makedirs(folder, exist_ok=True)

    tmp = path + ".tmp"

    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    os.replace(tmp, path)


def _norm_cufe(cufe: str) -> str:

    if not cufe:
        return ""

    s = normalize_text(str(cufe))

    out = []

    for ch in s:
        if ch in "0123456789abcdef":
            out.append(ch)

    return "".join(out)


def _norm_num(n: str) -> str:

    if not n:
        return ""

    s = str(n).strip().upper()

    s = s.replace("–", "-")
    s = s.replace("—", "-")
    s = s.replace("_", "-")

    s = re.sub(r"\s+", " ", s).strip()

    s = re.sub(r"[^A-Z0-9\-/. ]", "", s)

    s = re.sub(r"\s*-\s*", "-", s)
    s = re.sub(r"\s*/\s*", "/", s)
    s = re.sub(r"\s*\.\s*", ".", s)

    s = re.sub(r"\s+", " ", s).strip()

    return s


def _solo_alnum(s: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", (s or "").upper())


def _num_variants(n: str) -> List[str]:

    base = _norm_num(n)

    if not base:
        return []

    out: List[str] = []

    def add(v: str):

        v = (v or "").strip()

        if v and v not in out:
            out.append(v)

    add(base)

    add(base.replace(" ", ""))

    add(base.replace(" ", "").replace(".", ""))

    add(base.replace(" ", "").replace("-", ""))

    add(_solo_alnum(base))

    compact = _solo_alnum(base)

    m = re.match(r"^([A-Z]+)(\d+)$", compact)

    if m:

        pref, dig = m.groups()

        add(f"{pref}{dig}")
        add(f"{pref}-{dig}")
        add(f"{pref} {dig}")

    m2 = re.match(r"^([A-Z]+)[\s\-/.]*(\d+)$", base)

    if m2:

        pref, dig = m2.groups()

        add(f"{pref}{dig}")
        add(f"{pref}-{dig}")
        add(f"{pref} {dig}")

    uniq: List[str] = []

    seen = set()

    for value in out:

        if value and value not in seen:

            seen.add(value)

            uniq.append(value)

    return uniq


def _norm_fecha(fecha: str) -> str:

    if not fecha:
        return ""

    return str(fecha).strip()


class AttachmentIndexStore:

    def __init__(
        self,
        path: str,
        ttl_days: int = 365,
        max_nf_per_key: int = 10,
        max_num_per_key: int = 10,
    ):

        self.path = path

        self.ttl_days = max(1, int(ttl_days))

        self.max_nf_per_key = max(1, int(max_nf_per_key))

        self.max_num_per_key = max(1, int(max_num_per_key))

        self._data = _safe_read_json(self.path)

        self._ensure_shape()

    def _ensure_shape(self):

        if not isinstance(self._data, dict):
            self._data = {}

        self._data.setdefault("meta", {})

        self._data.setdefault("zip_by_cufe", {})
        self._data.setdefault("zip_by_num", {})
        self._data.setdefault("zip_by_nf", {})
        self._data.setdefault("conta15_by_cufe", {})

    def _save(self):

        self._data["meta"]["updated_ts"] = _utc_now_ts()

        _safe_write_json(self.path, self._data)

    def purge(self) -> int:

        cutoff = _utc_now_ts() - int(self.ttl_days * 86400)

        removed = 0

        def is_old(entry):

            ts = entry.get("received_ts") or entry.get("added_ts") or 0

            try:
                ts = int(ts)
            except Exception:
                ts = 0

            return bool(ts) and ts < cutoff

        for key in list(self._data["zip_by_cufe"].keys()):

            if is_old(self._data["zip_by_cufe"][key]):

                del self._data["zip_by_cufe"][key]

                removed += 1

        for table in ["zip_by_num", "zip_by_nf"]:

            data = self._data.get(table, {})

            for key in list(data.keys()):

                arr = data.get(key) or []

                new_arr = [entry for entry in arr if not is_old(entry)]

                if not new_arr:

                    del data[key]

                    removed += 1

                else:

                    data[key] = new_arr

        if removed:
            self._save()

        return removed

    def upsert_zip(
        self,
        cufe: str,
        numero: str,
        fecha: str,
        msg_id: str,
        att_id: str,
        att_name: str,
        received_dt_iso: str = "",
    ):

        cufe_n = _norm_cufe(cufe)

        numero_n = _norm_num(numero)

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
            "numero": numero_n,
            "fecha": fecha_n,
        }

        if cufe_n:
            self._data["zip_by_cufe"][cufe_n] = entry

        for nv in numero_vars:

            arr = self._data["zip_by_num"].get(nv, [])

            exists = any(
                e["msg_id"] == msg_id and e["att_id"] == att_id
                for e in arr
            )

            if not exists:

                arr.insert(0, entry)

            if len(arr) > self.max_num_per_key:

                arr = arr[: self.max_num_per_key]

            self._data["zip_by_num"][nv] = arr

        if fecha_n:

            for nv in numero_vars:

                key = f"{nv}|{fecha_n}"

                arr = self._data["zip_by_nf"].get(key, [])

                exists = any(
                    e["msg_id"] == msg_id and e["att_id"] == att_id
                    for e in arr
                )

                if not exists:
                    arr.insert(0, entry)

                if len(arr) > self.max_nf_per_key:
                    arr = arr[: self.max_nf_per_key]

                self._data["zip_by_nf"][key] = arr

        self._save()

    def find_zip_by_cufe(self, cufe: str) -> Optional[Dict[str, Any]]:

        cufe_n = _norm_cufe(cufe)

        if not cufe_n:
            return None

        return self._data["zip_by_cufe"].get(cufe_n)

    def find_zip_by_numero(self, numero: str) -> Optional[Dict[str, Any]]:

        for nv in _num_variants(numero):

            arr = self._data["zip_by_num"].get(nv)

            if arr:
                return arr[0]

        target = _solo_alnum(numero)

        for key, arr in self._data["zip_by_num"].items():

            if _solo_alnum(key) == target and arr:
                return arr[0]

        return None

    def find_zip_by_num_fecha(self, numero: str, fecha: str) -> Optional[Dict[str, Any]]:

        fecha_n = _norm_fecha(fecha)

        if not fecha_n:
            return None

        for nv in _num_variants(numero):

            key = f"{nv}|{fecha_n}"

            arr = self._data["zip_by_nf"].get(key)

            if arr:
                return arr[0]

        return None