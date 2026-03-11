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
        # Cerrar no es obligatorio; Graph a veces la expira igual.
        # Pero lo dejamos por orden.
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
        table_values: List[List[Any]], key_cols: Tuple[str, str]
    ) -> Set[Tuple[str, str]]:
        if not table_values or len(table_values) < 2:
            return set()

        header = [str(x).strip() for x in (table_values[0] or [])]
        try:
            i1 = header.index(key_cols[0])
            i2 = header.index(key_cols[1])
        except ValueError:
            # Si no están las columnas, no podemos dedupe => devolvemos vacío
            return set()

        existing: Set[Tuple[str, str]] = set()
        for row in table_values[1:]:
            v1 = str(row[i1]).strip() if i1 < len(row) and row[i1] is not None else ""
            v2 = str(row[i2]).strip() if i2 < len(row) and row[i2] is not None else ""
            if v1 and v2:
                existing.add((v1, v2))
        return existing

    @staticmethod
    def _align_rows_to_table(header: List[str], rows_dicts: List[Dict[str, Any]]) -> List[List[Any]]:
        # Alinea a columnas existentes de la tabla; las faltantes van vacío
        return [[d.get(col, "") for col in header] for d in rows_dicts]

    # ---------------------------
    # Public API
    # ---------------------------
    def append_rows_dedup(
        self,
        table_name: str,
        rows_dicts: List[Dict[str, Any]],
        key_cols: Tuple[str, str] = ("Archivo", "Concepto"),
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

            filtered = []
            for d in rows_dicts:
                k = (
                    str(d.get(key_cols[0], "")).strip(),
                    str(d.get(key_cols[1], "")).strip(),
                )
                if k[0] and k[1] and k not in existing:
                    filtered.append(d)

            if not filtered:
                return 0

            rows_values = self._align_rows_to_table(header, filtered)

            # Insert con reintentos (a veces Graph devuelve 429/5xx)
            last_err: Optional[Exception] = None
            for attempt in range(retries + 1):
                try:
                    self.add_rows(session_id, table_name, rows_values)
                    return len(filtered)
                except Exception as e:
                    last_err = e
                    if attempt < retries:
                        time.sleep(retry_sleep * (2 ** attempt))
                        continue
                    raise

            if last_err:
                raise last_err
            return 0
        finally:
            self.close_session(session_id)