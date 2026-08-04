"""Carrier directory service: list / merge (SHIP-101)."""

from __future__ import annotations

import sqlite3

from app.schemas.logistics import (
    CarrierItem,
    CarrierListResponse,
    CarrierMergeResponse,
)
from core.kp_db_common import _connect
from core.kp_db_schema import ensure_schema


class CarrierError(ValueError):
    """Domain validation error for carrier operations (maps to 422)."""

    def __init__(self, message: str, *, code: str = "carrier_error") -> None:
        super().__init__(message)
        self.code = code


class CarrierService:
    def __init__(self, *, db_path: str) -> None:
        self.db_path = db_path

    def list_carriers(
        self,
        *,
        q: str | None = None,
        active: bool = True,
        limit: int = 100,
    ) -> CarrierListResponse:
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            clauses: list[str] = []
            params: list = []
            if active:
                clauses.append("c.active = 1")
            if q and q.strip():
                # LIKE в SQLite case-insensitive только для ASCII — ищем через casefold().
                clauses.append("(casefold(c.name) LIKE ? OR c.name_normalized LIKE ?)")
                needle = f"%{q.strip().casefold()}%"
                params.extend([needle, needle])
            where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
            cur.execute(
                f"""
                SELECT
                    c.id, c.name, c.source_sheet, c.note, c.active, c.merged_into_id,
                    (SELECT COUNT(*) FROM shipments s WHERE s.carrier_id = c.id) AS shipments_count
                FROM carriers c
                {where}
                ORDER BY c.name
                LIMIT ?
                """,
                (*params, int(limit)),
            )
            items = [
                CarrierItem(
                    id=int(row["id"]),
                    name=str(row["name"]),
                    source_sheet=row["source_sheet"],
                    note=row["note"],
                    active=bool(row["active"]),
                    merged_into_id=row["merged_into_id"],
                    shipments_count=int(row["shipments_count"] or 0),
                )
                for row in cur.fetchall()
            ]
            return CarrierListResponse(items=items, count=len(items))
        finally:
            conn.close()

    def merge(
        self,
        carrier_id: int,
        into_id: int,
        *,
        actor: str | None = None,
    ) -> CarrierMergeResponse:
        if int(carrier_id) == int(into_id):
            raise CarrierError(
                "Нельзя слить перевозчика самого в себя",
                code="carrier_merge_conflict",
            )
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            cur = conn.cursor()
            source = self._fetch_carrier(cur, carrier_id)
            target = self._fetch_carrier(cur, into_id)
            if source is None:
                raise CarrierError(
                    f"Перевозчик #{carrier_id} не найден",
                    code="carrier_not_found",
                )
            if target is None:
                raise CarrierError(
                    f"Целевой перевозчик #{into_id} не найден",
                    code="carrier_not_found",
                )
            if not target["active"] or target["merged_into_id"] is not None:
                raise CarrierError(
                    f"Нельзя слить в неактивного перевозчика «{target['name']}»",
                    code="carrier_merge_conflict",
                )
            cur.execute(
                "UPDATE shipments SET carrier_id = ? WHERE carrier_id = ?",
                (int(into_id), int(carrier_id)),
            )
            moved = int(cur.rowcount or 0)
            cur.execute(
                "UPDATE carriers SET merged_into_id = ?, active = 0 WHERE id = ?",
                (int(into_id), int(carrier_id)),
            )
            conn.commit()
            return CarrierMergeResponse(
                ok=True,
                carrier_id=int(carrier_id),
                into_id=int(into_id),
                moved_shipments=moved,
                message=(
                    f"Перевозчик «{source['name']}» слит в «{target['name']}», "
                    f"перенесено рейсов: {moved}"
                ),
            )
        except CarrierError:
            conn.rollback()
            raise
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    @staticmethod
    def _fetch_carrier(cur: sqlite3.Cursor, carrier_id: int) -> dict | None:
        cur.execute(
            "SELECT id, name, active, merged_into_id FROM carriers WHERE id = ?",
            (int(carrier_id),),
        )
        row = cur.fetchone()
        if row is None:
            return None
        return {
            "id": int(row[0]),
            "name": str(row[1]),
            "active": bool(row[2]),
            "merged_into_id": row[3],
        }
