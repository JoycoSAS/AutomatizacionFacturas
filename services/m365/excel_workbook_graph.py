# services/m365/excel_workbook_graph.py
import time
import math
import requests
from decimal import Decimal, InvalidOperation
from typing import List, Dict, Set, Tuple, Optional, Any

from .sp_graph import get_item_by_path, DRIVE_ID, SSL_VERIFY, TIMEOUT, GRAPH
from .token import get_access_token

_SESSION = requests.Session()


def _h(session_id: Optional[str] = None) -> Dict[str, str]:
    h = {
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    if session_id:
        h["workbook-session-id"] = session_id
    return h


# Columnas que Excel Web NO debe interpretar como número.
# Si Excel las interpreta como número, las convierte a notación científica
# o pierde precisión.
TEXT_COLUMNS_FORCE = {
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
    "Nombre contrato",
    "Unidad económica",
    "DESCRIPCIÓN",
    "Concepto",
    "Estado_calidad",
}


def _plain_number_text(value: Any) -> str:
    """
    Convierte valores numéricos a texto plano evitando notación científica.
    Se usa solo para columnas que deben guardarse como texto.
    """
    if value is None:
        return ""

    if isinstance(value, bool):
        return "true" if value else "false"

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return format(value, "f").rstrip("0").rstrip(".")

    s = str(value).strip()
    if not s:
        return ""

    if s.startswith("'"):
        return s

    s_num = s.replace(",", ".")
    if "e+" in s_num.lower() or "e-" in s_num.lower():
        try:
            d = Decimal(s_num)
            if d == d.to_integral_value():
                return format(d.quantize(Decimal(1)), "f")
            return format(d, "f").rstrip("0").rstrip(".")
        except (InvalidOperation, ValueError):
            return s

    return s


def _excel_text(value: Any) -> str:
    """
    Fuerza texto en Excel. El apóstrofo inicial evita que Excel Online
    convierta CUFE/NIT/Número de factura a notación científica.
    """
    s = _plain_number_text(value).strip()
    if not s:
        return ""

    if s.startswith("'"):
        return s

    return "'" + s


def _normalizar_valor_para_excel_web(col: str, value: Any) -> Any:
    """
    Normaliza cada valor antes de enviarlo al Workbook API.
    """
    col = str(col or "").strip()

    if col in TEXT_COLUMNS_FORCE:
        return _excel_text(value)

    if value is None:
        return ""

    return value



class ExcelWorkbookGraph:
    """
    Escribe en un Excel en SharePoint usando Graph Workbook API (sin reemplazar el archivo).
    ✅ Permite que el Excel esté abierto en Excel Online (evita 423 locked del /content).
    ✅ Soporta drive_id opcional (por defecto usa SP_DRIVE_ID -> DRIVE_ID).
    ✅ Fuerza texto en columnas críticas para evitar notación científica y pérdida de precisión.

    IMPORTANTÍSIMO:
    - Para /tables/{table}/... debe existir una TABLA real en el Excel (ej: TblFacturas).
    """

    def __init__(self, sp_excel_rel_path: str, drive_id: Optional[str] = None):
        self.drive_id = (drive_id or DRIVE_ID or "").strip()
        if not self.drive_id:
            raise RuntimeError("No hay DRIVE_ID disponible para Workbook API (SP_DRIVE_ID).")

        self.sp_excel_rel_path = sp_excel_rel_path.strip().strip("/")

        item = get_item_by_path(self.sp_excel_rel_path, drive_id=self.drive_id)
        self.item_id = item["id"]

        self.base = f"{GRAPH}/drives/{self.drive_id}/items/{self.item_id}/workbook"

    # ---------------------------
    # Sesión
    # ---------------------------
    def create_session(self, persist_changes: bool = True) -> str:
        url = f"{self.base}/createSession"
        r = _SESSION.post(
            url,
            headers=_h(),
            json={"persistChanges": persist_changes},
            timeout=TIMEOUT,
            verify=SSL_VERIFY,
        )
        r.raise_for_status()
        data = r.json()
        sid = data.get("id")
        if not sid:
            raise RuntimeError(f"[Workbook] createSession sin id. Resp: {data}")
        return sid

    def close_session(self, session_id: str) -> None:
        try:
            url = f"{self.base}/closeSession"
            _SESSION.post(
                url,
                headers=_h(session_id),
                json={},
                timeout=TIMEOUT,
                verify=SSL_VERIFY,
            )
        except Exception:
            pass

    # ---------------------------
    # Utilidades Workbook
    # ---------------------------
    def list_tables(self, session_id: str) -> List[str]:
        url = f"{self.base}/tables"
        r = _SESSION.get(
            url,
            headers=_h(session_id),
            timeout=TIMEOUT,
            verify=SSL_VERIFY,
        )
        r.raise_for_status()
        data = r.json() or {}
        vals = data.get("value") or []
        names = []
        for t in vals:
            n = (t.get("name") or "").strip()
            if n:
                names.append(n)
        return names

    def table_exists(self, session_id: str, table_name: str) -> bool:
        try:
            names = self.list_tables(session_id)
            return table_name in names
        except Exception:
            return False

    def get_table_range_values(self, session_id: str, table_name: str) -> List[List[Any]]:
        """
        Devuelve values del range de la tabla:
        - values[0] = header
        - values[1:] = filas
        """
        url = f"{self.base}/tables/{table_name}/range"
        r = _SESSION.get(
            url,
            headers=_h(session_id),
            timeout=TIMEOUT,
            verify=SSL_VERIFY,
        )
        r.raise_for_status()
        data = r.json() or {}
        return data.get("values") or []

    def add_rows(self, session_id: str, table_name: str, rows: List[List[Any]]) -> None:
        if not rows:
            return
        url = f"{self.base}/tables/{table_name}/rows/add"
        payload = {"index": None, "values": rows}
        r = _SESSION.post(
            url,
            headers=_h(session_id),
            json=payload,
            timeout=TIMEOUT,
            verify=SSL_VERIFY,
        )
        r.raise_for_status()

    # ---------------------------
    # Dedupe helpers
    # ---------------------------
    @staticmethod
    def _clean_key_value(value: Any) -> str:
        """
        Normaliza valores de llave tanto si vienen con apóstrofo como si vienen normales.
        """
        s = _plain_number_text(value).strip()
        if s.startswith("'"):
            s = s[1:].strip()
        return s

    @classmethod
    def _build_existing_keys(
        cls, table_values: List[List[Any]], key_cols: Tuple[str, ...]
    ) -> Set[Tuple[str, ...]]:
        if not table_values or len(table_values) < 2:
            return set()

        header = [str(x).strip() for x in (table_values[0] or [])]

        idxs = []
        for col in key_cols:
            try:
                idxs.append(header.index(col))
            except ValueError:
                return set()

        existing: Set[Tuple[str, ...]] = set()
        for row in table_values[1:]:
            vals = []
            ok = True
            for idx in idxs:
                v = cls._clean_key_value(row[idx]) if idx < len(row) and row[idx] is not None else ""
                if not v:
                    ok = False
                    break
                vals.append(v)
            if ok:
                existing.add(tuple(vals))
        return existing

    @staticmethod
    def _align_rows_to_table(header: List[str], rows_dicts: List[Dict[str, Any]]) -> List[List[Any]]:
        """
        Alinea dicts al orden real de columnas de la tabla.
        Aquí se fuerza texto para columnas críticas antes de enviar al Workbook API.
        """
        rows: List[List[Any]] = []

        for d in rows_dicts:
            row = []
            for col in header:
                row.append(_normalizar_valor_para_excel_web(col, d.get(col, "")))
            rows.append(row)

        return rows

    # ---------------------------
    # Public API
    # ---------------------------
    def append_rows_dedup(
        self,
        table_name: str,
        rows_dicts: List[Dict[str, Any]],
        key_cols: Tuple[str, ...] = ("Radicado", "Archivo", "Concepto"),
        require_table: bool = True,
        retries: int = 2,
        retry_sleep: float = 1.0,
    ) -> int:
        """
        1) Abre sesión workbook
        2) Verifica que exista la tabla (si require_table=True)
        3) Lee tabla
        4) Dedupe por key_cols
        5) Inserta solo nuevas
        """
        if not rows_dicts:
            print("[Workbook] append_rows_dedup: rows_dicts vacío")
            return 0

        session_id = self.create_session(persist_changes=True)
        try:
            if require_table and not self.table_exists(session_id, table_name):
                tables = self.list_tables(session_id)
                raise RuntimeError(
                    f"[Workbook] No se encontró la tabla '{table_name}'. "
                    f"Tablas disponibles: {tables or 'Ninguna'}"
                )

            table_values = self.get_table_range_values(session_id, table_name)
            if not table_values:
                raise RuntimeError(f"[Workbook] Tabla '{table_name}' sin datos (range vacío).")

            header = [str(x).strip() for x in (table_values[0] or [])]
            if not header:
                raise RuntimeError(f"[Workbook] Tabla '{table_name}' sin encabezados.")

            existing = self._build_existing_keys(table_values, key_cols)

            # Si solo existe el encabezado, no hay filas reales
            if len(table_values) <= 1:
                existing = set()

            # Sanitizar antes de insertar:
            # - NO dejar pasar filas sin Concepto.
            # - NO dejar pasar filas con llave incompleta.
            # Esto evita la fila adicional/fantasma en Excel Web.
            sanitized: List[Dict[str, Any]] = []
            descartadas_sin_concepto = 0
            descartadas_sin_llave = 0

            for raw in rows_dicts:
                if not isinstance(raw, dict):
                    continue

                d = dict(raw)

                concepto = self._clean_key_value(d.get("Concepto", ""))
                if not concepto:
                    descartadas_sin_concepto += 1
                    continue
                d["Concepto"] = concepto

                for col in key_cols:
                    if col in d and d[col] is not None:
                        d[col] = self._clean_key_value(d[col])

                k = tuple(self._clean_key_value(d.get(col, "")) for col in key_cols)
                if not all(k):
                    descartadas_sin_llave += 1
                    continue

                sanitized.append(d)

            filtered: List[Dict[str, Any]] = []
            seen_new: Set[Tuple[str, ...]] = set()

            for d in sanitized:
                k = tuple(self._clean_key_value(d.get(col, "")) for col in key_cols)

                if k in seen_new:
                    continue

                if k not in existing:
                    filtered.append(d)
                    seen_new.add(k)

            print(
                f"[Workbook] Tabla={table_name} | "
                f"rows_dicts={len(rows_dicts)} | "
                f"sanitized={len(sanitized)} | "
                f"table_rows={max(len(table_values) - 1, 0)} | "
                f"existing_keys={len(existing)} | "
                f"filtered={len(filtered)} | "
                f"desc_sin_concepto={descartadas_sin_concepto} | "
                f"desc_sin_llave={descartadas_sin_llave}"
            )

            if not filtered:
                print("[Workbook] No hay filas nuevas para insertar.")
                return 0

            rows_values = self._align_rows_to_table(header, filtered)

            last_err: Optional[Exception] = None
            for attempt in range(retries + 1):
                try:
                    self.add_rows(session_id, table_name, rows_values)
                    print(f"[Workbook] Insertadas correctamente: {len(filtered)} fila(s)")
                    return len(filtered)
                except Exception as e:
                    last_err = e
                    print(f"[Workbook] Error insert attempt={attempt + 1}: {e}")
                    if attempt < retries:
                        time.sleep(retry_sleep * (2 ** attempt))
                        continue
                    raise

            if last_err:
                raise last_err
            return 0
        finally:
            self.close_session(session_id)