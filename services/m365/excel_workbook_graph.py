# services/m365/excel_workbook_graph.py
import requests
from typing import List, Dict, Set, Tuple, Optional

from .sp_graph import get_item_by_path, DRIVE_ID, SSL_VERIFY, TIMEOUT, GRAPH
from .token import get_access_token

_SESSION = requests.Session()


def _h(session_id: Optional[str] = None) -> Dict[str, str]:
    h = {
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type": "application/json",
    }
    if session_id:
        h["workbook-session-id"] = session_id
    return h


class ExcelWorkbookGraph:
    """
    Escribe en un Excel en SharePoint usando Graph Workbook API (sin reemplazar el archivo).
    Ideal para que el Excel esté abierto en Excel Online y no crashee el flujo.
    """

    def __init__(self, sp_excel_rel_path: str):
        """
        sp_excel_rel_path ejemplo:
          "Innovacion/08. Pruebas proyectos/autoFacturas/excel/facturas.xlsx"
        """
        self.sp_excel_rel_path = sp_excel_rel_path.strip().strip("/")
        item = get_item_by_path(self.sp_excel_rel_path)
        self.item_id = item["id"]
        self.base = f"{GRAPH}/drives/{DRIVE_ID}/items/{self.item_id}/workbook"

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
        return r.json()["id"]

    def get_table_values(self, session_id: str, table_name: str) -> List[List]:
        url = f"{self.base}/tables/{table_name}/range"
        r = _SESSION.get(
            url,
            headers=_h(session_id),
            timeout=TIMEOUT,
            verify=SSL_VERIFY,
        )
        r.raise_for_status()
        data = r.json()
        return data.get("values") or []

    def add_rows(self, session_id: str, table_name: str, rows: List[List]) -> None:
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

    @staticmethod
    def _build_existing_keys(
        table_values: List[List], key_cols: Tuple[str, str]
    ) -> Set[Tuple[str, str]]:
        if not table_values or len(table_values) < 2:
            return set()

        header = [str(x).strip() for x in table_values[0]]
        try:
            i1 = header.index(key_cols[0])
            i2 = header.index(key_cols[1])
        except ValueError:
            return set()

        existing = set()
        for row in table_values[1:]:
            v1 = str(row[i1]).strip() if i1 < len(row) and row[i1] is not None else ""
            v2 = str(row[i2]).strip() if i2 < len(row) and row[i2] is not None else ""
            if v1 and v2:
                existing.add((v1, v2))
        return existing

    @staticmethod
    def _align_rows_to_table(header: List[str], rows_dicts: List[Dict]) -> List[List]:
        out = []
        for d in rows_dicts:
            out.append([d.get(col, "") for col in header])
        return out

    def append_rows_dedup(
        self,
        table_name: str,
        rows_dicts: List[Dict],
        key_cols: Tuple[str, str] = ("Archivo", "Concepto"),
    ) -> int:
        """
        1) Lee tabla en nube
        2) Dedupe por key_cols
        3) Inserta solo nuevas
        """
        if not rows_dicts:
            return 0

        session_id = self.create_session(persist_changes=True)
        table_values = self.get_table_values(session_id, table_name)
        if not table_values:
            raise RuntimeError(f"[Workbook] Tabla {table_name} sin datos (range vacío).")

        header = [str(x).strip() for x in table_values[0]]
        existing = self._build_existing_keys(table_values, key_cols)

        filtered = []
        for d in rows_dicts:
            k = (str(d.get(key_cols[0], "")).strip(), str(d.get(key_cols[1], "")).strip())
            if k[0] and k[1] and k not in existing:
                filtered.append(d)

        if not filtered:
            return 0

        rows_values = self._align_rows_to_table(header, filtered)
        self.add_rows(session_id, table_name, rows_values)
        return len(filtered)
