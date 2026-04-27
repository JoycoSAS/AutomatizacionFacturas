# services/m365/excel_workbook_graph.py
import time
import requests
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


class ExcelWorkbookGraph:
    """
    Escribe en un Excel en SharePoint usando Graph Workbook API (sin reemplazar el archivo).
    ✅ Permite que el Excel esté abierto en Excel Online (evita 423 locked del /content).
    ✅ Soporta drive_id opcional (por defecto usa SP_DRIVE_ID -> DRIVE_ID).

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
    def _build_existing_keys(
        table_values: List[List[Any]], key_cols: Tuple[str, ...]
    ) -> Set[Tuple[str, ...]]:
        """
        Construye las llaves existentes usando una cantidad variable de columnas.

        Importante:
        - Antes esta función solo soportaba 2 columnas.
        - Ahora soporta 3 o más, por ejemplo:
          ("Radicado", "Archivo", "Concepto")
        """
        if not table_values or len(table_values) < 2:
            return set()

        header = [str(x).strip() for x in (table_values[0] or [])]

        idxs = []
        for col in key_cols:
            try:
                idxs.append(header.index(col))
            except ValueError:
                print(f"[Workbook] Columna de llave no encontrada en tabla: {col}")
                return set()

        existing: Set[Tuple[str, ...]] = set()

        for row in table_values[1:]:
            vals = []
            ok = True

            for idx in idxs:
                v = str(row[idx]).strip() if idx < len(row) and row[idx] is not None else ""
                if not v:
                    ok = False
                    break
                vals.append(v)

            if ok:
                existing.add(tuple(vals))

        return existing

    @staticmethod
    def _align_rows_to_table(header: List[str], rows_dicts: List[Dict[str, Any]]) -> List[List[Any]]:
        return [[d.get(col, "") for col in header] for d in rows_dicts]

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
        Reglas críticas:
        - NO insertar filas sin Concepto
        - NO insertar filas con llave incompleta
        - dedupe por (Radicado, Archivo, Concepto)
        """
        if not rows_dicts:
            print("[Workbook] append_rows_dedup: rows_dicts vacío")
            return 0

        sanitized: List[Dict[str, Any]] = []
        descartadas_sin_concepto = 0
        descartadas_sin_llave = 0

        for raw in rows_dicts:
            if not isinstance(raw, dict):
                continue

            d = dict(raw)

            concepto = str(d.get("Concepto", "")).strip()
            if not concepto:
                descartadas_sin_concepto += 1
                continue

            for col in key_cols:
                if col in d and d[col] is not None:
                    d[col] = str(d[col]).strip()

            k = tuple(str(d.get(col, "")).strip() for col in key_cols)
            if not all(k):
                descartadas_sin_llave += 1
                continue

            sanitized.append(d)

        if not sanitized:
            print(
                "[Workbook] No hay filas válidas para insertar. "
                f"descartadas_sin_concepto={descartadas_sin_concepto} | "
                f"descartadas_sin_llave={descartadas_sin_llave}"
            )
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
            if len(table_values) <= 1:
                existing = set()

            filtered: List[Dict[str, Any]] = []
            seen_new: Set[Tuple[str, ...]] = set()

            for d in sanitized:
                k = tuple(str(d.get(col, "")).strip() for col in key_cols)
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