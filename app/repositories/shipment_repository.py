"""SQL persistence for shipments (CRUD fetch/insert/replace).

Qty reservation helpers stay in ``core/kp_db_shipments``; this module owns
shipment / orders / items row access used by the facade and completion service.
"""

from __future__ import annotations

import sqlite3
from typing import Any

from app.schemas.logistics import (
    ShipmentAvailableByKp,
    ShipmentAvailableSgpRow,
    ShipmentItem,
    ShipmentOrderItem,
    ShipmentOrderPatch,
)
from app.services.shipment_errors import ShipmentError
from core.kp_db_common import _connect
from core.kp_db_schema import ensure_schema
from core.kp_db_shipments import available_qty
from core.kp_plate_weight import resolve_kp_line_weight_kg


class ShipmentRepository:
    def __init__(self, *, db_path: str) -> None:
        self.db_path = db_path

    def connect(self) -> sqlite3.Connection:
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def insert_shipment_with_orders(
        self,
        cur: sqlite3.Cursor,
        *,
        shipment_date: str,
        delivery_type: str,
        kp_ids: list[int],
        actor: str | None,
    ) -> int:
        cur.execute(
            "INSERT INTO shipments (shipment_date, delivery_type, actor) VALUES (?, ?, ?)",
            (shipment_date, delivery_type, actor),
        )
        shipment_id = int(cur.lastrowid)
        for kp_id in kp_ids:
            cur.execute(
                "INSERT INTO shipment_orders (shipment_id, kp_id, ya_order_no) VALUES (?, ?, ?)",
                (shipment_id, kp_id, self.prefill_ya_order_no(cur, kp_id)),
            )
        return shipment_id

    def copy_transport_fields(
        self,
        cur: sqlite3.Cursor,
        *,
        source: sqlite3.Row,
        shipment_id: int,
    ) -> None:
        cur.execute(
            """
            UPDATE shipments SET
                carrier_id = ?,
                driver_name = ?,
                vehicle_text = ?,
                vehicle_class = ?,
                proxy_no = ?
            WHERE id = ?
            """,
            (
                source["carrier_id"],
                source["driver_name"],
                source["vehicle_text"],
                source["vehicle_class"],
                source["proxy_no"],
                shipment_id,
            ),
        )

    def update_shipment_fields(
        self,
        cur: sqlite3.Cursor,
        shipment_id: int,
        fields: dict[str, Any],
    ) -> None:
        if not fields:
            return
        assignments = ", ".join(f"{name} = ?" for name in fields)
        cur.execute(
            f"UPDATE shipments SET {assignments} WHERE id = ?",
            (*fields.values(), int(shipment_id)),
        )

    def replace_items(
        self,
        cur: sqlite3.Cursor,
        shipment_id: int,
        prepared: list[dict[str, Any]],
    ) -> None:
        cur.execute("DELETE FROM shipment_items WHERE shipment_id = ?", (int(shipment_id),))
        for record in prepared:
            cur.execute(
                """
                INSERT INTO shipment_items (
                    shipment_id, item_type, completed_plate_id, kp_id, mark,
                    qty, unit_weight_kg, weight_kg, sort_order, note
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    int(shipment_id),
                    record["item_type"],
                    record["completed_plate_id"],
                    record["kp_id"],
                    record["mark"],
                    record["qty"],
                    record["unit_weight_kg"],
                    record["weight_kg"],
                    record["sort_order"],
                    record["note"],
                ),
            )

    def mark_done(
        self,
        cur: sqlite3.Cursor,
        shipment_id: int,
        *,
        completed_at: str,
        actor: str | None,
        status: str,
    ) -> None:
        cur.execute(
            "UPDATE shipments SET status = ?, completed_at = ?, actor = ? WHERE id = ?",
            (status, completed_at, actor, int(shipment_id)),
        )

    def delete_shipment(self, cur: sqlite3.Cursor, shipment_id: int) -> None:
        cur.execute("DELETE FROM shipments WHERE id = ?", (int(shipment_id),))

    def update_propose_snapshot(
        self, cur: sqlite3.Cursor, shipment_id: int, snapshot_json: str
    ) -> None:
        cur.execute(
            "UPDATE shipments SET propose_snapshot = ? WHERE id = ?",
            (snapshot_json, int(shipment_id)),
        )

    def list_shipment_rows(
        self,
        cur: sqlite3.Cursor,
        *,
        clauses: list[str],
        params: list[Any],
        limit: int,
        offset: int,
    ) -> list[sqlite3.Row]:
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        cur.execute(
            f"""
            SELECT s.*, c.name AS carrier_name
            FROM shipments s
            LEFT JOIN carriers c ON c.id = s.carrier_id
            {where}
            ORDER BY s.shipment_date DESC, s.id DESC
            LIMIT ? OFFSET ?
            """,
            (*params, int(limit), int(offset)),
        )
        return list(cur.fetchall())

    @staticmethod
    def fetch_shipment_row(cur: sqlite3.Cursor, shipment_id: int) -> sqlite3.Row:
        cur.execute(
            """
            SELECT s.*, c.name AS carrier_name
            FROM shipments s
            LEFT JOIN carriers c ON c.id = s.carrier_id
            WHERE s.id = ?
            """,
            (int(shipment_id),),
        )
        row = cur.fetchone()
        if row is None:
            raise ShipmentError(
                f"Рейс #{shipment_id} не найден",
                code="shipment_not_found",
            )
        return row

    @staticmethod
    def fetch_orders(
        cur: sqlite3.Cursor, shipment_ids: list[int]
    ) -> dict[int, list[ShipmentOrderItem]]:
        if not shipment_ids:
            return {}
        placeholders = ",".join("?" * len(shipment_ids))
        cur.execute(
            f"""
            SELECT so.id, so.shipment_id, so.kp_id, so.ya_order_no, o.customer_name
            FROM shipment_orders so
            LEFT JOIN KP_offers o ON o.kp_id = so.kp_id
            WHERE so.shipment_id IN ({placeholders})
            ORDER BY so.id
            """,
            [int(sid) for sid in shipment_ids],
        )
        result: dict[int, list[ShipmentOrderItem]] = {}
        for row in cur.fetchall():
            result.setdefault(int(row["shipment_id"]), []).append(
                ShipmentOrderItem(
                    id=int(row["id"]),
                    kp_id=row["kp_id"],
                    ya_order_no=row["ya_order_no"],
                    customer_name=row["customer_name"],
                )
            )
        return result

    @staticmethod
    def fetch_items(cur: sqlite3.Cursor, shipment_id: int) -> list[ShipmentItem]:
        cur.execute(
            """
            SELECT si.id, si.item_type, si.completed_plate_id, si.kp_id, si.mark,
                   si.qty, si.unit_weight_kg, si.weight_kg, si.sort_order, si.note,
                   cp.plate_name, cp.length_m, cp.width_m, cp.load_class
            FROM shipment_items si
            LEFT JOIN completed_plates cp ON cp.id = si.completed_plate_id
            WHERE si.shipment_id = ?
            ORDER BY si.sort_order, si.id
            """,
            (int(shipment_id),),
        )
        return [
            ShipmentItem(
                id=int(row["id"]),
                item_type=str(row["item_type"]),
                completed_plate_id=row["completed_plate_id"],
                kp_id=row["kp_id"],
                mark=row["mark"],
                plate_name=row["plate_name"],
                length_m=row["length_m"],
                width_m=row["width_m"],
                load_class=row["load_class"],
                qty=int(row["qty"]),
                unit_weight_kg=row["unit_weight_kg"],
                weight_kg=row["weight_kg"],
                sort_order=int(row["sort_order"] or 0),
                note=row["note"],
            )
            for row in cur.fetchall()
        ]

    @staticmethod
    def fetch_weight_totals(
        cur: sqlite3.Cursor, shipment_ids: list[int]
    ) -> dict[int, float]:
        if not shipment_ids:
            return {}
        placeholders = ",".join("?" * len(shipment_ids))
        cur.execute(
            f"""
            SELECT shipment_id, COALESCE(SUM(weight_kg), 0)
            FROM shipment_items
            WHERE shipment_id IN ({placeholders})
            GROUP BY shipment_id
            """,
            [int(sid) for sid in shipment_ids],
        )
        return {int(row[0]): float(row[1] or 0.0) for row in cur.fetchall()}

    def available_by_kp(
        self, cur: sqlite3.Cursor, orders: list[ShipmentOrderItem]
    ) -> list[ShipmentAvailableByKp]:
        groups: list[ShipmentAvailableByKp] = []
        for order in orders:
            if order.kp_id is None:
                continue
            cur.execute(
                """
                SELECT id, kp_id, plate_name, length_m, width_m, load_class,
                       qty, completed_date
                FROM completed_plates
                WHERE kp_id = ? AND qty > 0
                ORDER BY completed_date, id
                """,
                (order.kp_id,),
            )
            plates: list[ShipmentAvailableSgpRow] = []
            for cp in cur.fetchall():
                available = available_qty(cur, int(cp["id"]))
                if available <= 0:
                    continue
                unit, _ = resolve_kp_line_weight_kg(
                    {"length_m": cp["length_m"], "width_m": cp["width_m"], "qty": 1}
                )
                plates.append(
                    ShipmentAvailableSgpRow(
                        completed_plate_id=int(cp["id"]),
                        kp_id=int(cp["kp_id"]),
                        plate_name=str(cp["plate_name"] or ""),
                        length_m=cp["length_m"],
                        width_m=cp["width_m"],
                        load_class=cp["load_class"],
                        qty=int(cp["qty"]),
                        available_qty=available,
                        unit_weight_kg=unit,
                        completed_date=cp["completed_date"],
                    )
                )
            groups.append(ShipmentAvailableByKp(kp_id=order.kp_id, plates=plates))
        return groups

    def replace_orders(
        self,
        cur: sqlite3.Cursor,
        shipment_id: int,
        orders: list[ShipmentOrderPatch],
    ) -> None:
        kp_ids = list(dict.fromkeys(int(order.kp_id) for order in orders))
        if not kp_ids:
            raise ShipmentError(
                "Рейс требует хотя бы один заказ (КП)",
                code="shipment_no_orders",
            )
        self.assert_kp_exists(cur, kp_ids)
        cur.execute("DELETE FROM shipment_orders WHERE shipment_id = ?", (int(shipment_id),))
        ya_by_kp = {
            int(order.kp_id): (order.ya_order_no or "").strip() or None for order in orders
        }
        for kp_id in kp_ids:
            ya = ya_by_kp.get(kp_id) or self.prefill_ya_order_no(cur, kp_id)
            cur.execute(
                "INSERT INTO shipment_orders (shipment_id, kp_id, ya_order_no) VALUES (?, ?, ?)",
                (int(shipment_id), kp_id, ya),
            )

    @staticmethod
    def prefill_ya_order_no(cur: sqlite3.Cursor, kp_id: int) -> str | None:
        cur.execute(
            """
            SELECT so.ya_order_no
            FROM shipment_orders so
            JOIN shipments s ON s.id = so.shipment_id
            WHERE so.kp_id = ?
              AND so.ya_order_no IS NOT NULL
              AND so.ya_order_no != ''
            ORDER BY s.created_at DESC, so.id DESC
            LIMIT 1
            """,
            (int(kp_id),),
        )
        row = cur.fetchone()
        return str(row[0]) if row else None

    @staticmethod
    def assert_kp_exists(cur: sqlite3.Cursor, kp_ids: list[int]) -> None:
        for kp_id in kp_ids:
            cur.execute("SELECT kp_id FROM KP_offers WHERE kp_id = ?", (kp_id,))
            if cur.fetchone() is None:
                raise ShipmentError(
                    f"КП #{kp_id} не найдено",
                    code="shipment_kp_not_found",
                )

    @staticmethod
    def assert_carrier_exists(cur: sqlite3.Cursor, carrier_id: int) -> None:
        cur.execute("SELECT id FROM carriers WHERE id = ?", (carrier_id,))
        if cur.fetchone() is None:
            raise ShipmentError(
                f"Перевозчик #{carrier_id} не найден",
                code="shipment_carrier_not_found",
            )

    @staticmethod
    def assert_kp_in_shipment_orders(
        cur: sqlite3.Cursor,
        shipment_id: int,
        kp_id: int,
        *,
        code: str,
        detail: str,
    ) -> None:
        cur.execute(
            "SELECT 1 FROM shipment_orders WHERE shipment_id = ? AND kp_id = ?",
            (int(shipment_id), int(kp_id)),
        )
        if cur.fetchone() is None:
            raise ShipmentError(detail, code=code)

    @staticmethod
    def pile_weight_for_mark(cur: sqlite3.Cursor, mark: str | None) -> float | None:
        if not mark:
            return None
        cur.execute("SELECT weight_kg FROM pile_catalog WHERE mark = ?", (mark,))
        row = cur.fetchone()
        return float(row[0]) if row else None

    @staticmethod
    def fetch_completed_plate(cur: sqlite3.Cursor, completed_plate_id: int) -> sqlite3.Row | None:
        cur.execute(
            """
            SELECT id, kp_id, plate_name, length_m, width_m, qty
            FROM completed_plates WHERE id = ?
            """,
            (int(completed_plate_id),),
        )
        return cur.fetchone()

    @staticmethod
    def fetch_completed_plate_for_ship(
        cur: sqlite3.Cursor, completed_plate_id: int
    ) -> sqlite3.Row | None:
        cur.execute(
            """
            SELECT id, kp_id, plate_name, qty, production_day, plan_id
            FROM completed_plates WHERE id = ?
            """,
            (int(completed_plate_id),),
        )
        return cur.fetchone()
