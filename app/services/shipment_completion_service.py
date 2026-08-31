"""Shipment completion lifecycle: put_items / complete / cancel (+ S3 KP membership)."""

from __future__ import annotations

import sqlite3
from datetime import datetime
from typing import Any

from app.domain.enums import (
    KpStatus,
    PlateStatus,
    PlateTransitionReason,
    ShipmentItemType,
    ShipmentStatus,
)
from app.repositories.shipment_repository import ShipmentRepository
from app.schemas.logistics import ShipmentItem, ShipmentItemInput, ShipmentMutationResponse
from app.services.shipment_errors import ShipmentError
from core.kp_db_audit import audit_append
from core.kp_db_plates_completion import (
    check_and_update_kp_completion,
    freeze_ordered_qty_if_needed,
)
from core.kp_db_schema import ensure_schema
from core.kp_db_shipments import available_qty, shipped_qty_for_kp
from core.kp_plate_weight import resolve_kp_line_weight_kg


class ShipmentCompletionService:
    def __init__(
        self,
        *,
        db_path: str,
        repo: ShipmentRepository | None = None,
        maybe_write_event=None,
    ) -> None:
        self.db_path = db_path
        self._repo = repo or ShipmentRepository(db_path=db_path)
        self._maybe_write_event = maybe_write_event

    def put_items(
        self,
        shipment_id: int,
        items: list[ShipmentItemInput],
        *,
        actor: str | None = None,
    ) -> None:
        ensure_schema(self.db_path)
        conn = self._repo.connect()
        try:
            cur = conn.cursor()
            row = self._repo.fetch_shipment_row(cur, shipment_id)
            self._assert_in_work(row, shipment_id)
            prepared: list[dict[str, Any]] = []
            claimed: dict[int, int] = {}
            for idx, item in enumerate(items):
                record = self._prepare_item(cur, shipment_id, idx, item, claimed)
                prepared.append(record)
            self._repo.replace_items(cur, shipment_id, prepared)
            conn.commit()
        except ShipmentError:
            conn.rollback()
            raise
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _prepare_item(
        self,
        cur: sqlite3.Cursor,
        shipment_id: int,
        index: int,
        item: ShipmentItemInput,
        claimed: dict[int, int],
    ) -> dict[str, Any]:
        record: dict[str, Any] = {
            "item_type": item.item_type,
            "completed_plate_id": None,
            "kp_id": None,
            "mark": None,
            "qty": int(item.qty),
            "unit_weight_kg": None,
            "weight_kg": None,
            "sort_order": item.sort_order if item.sort_order is not None else index,
            "note": item.note,
        }
        if item.item_type == ShipmentItemType.PLATE.value:
            if item.completed_plate_id is None:
                raise ShipmentError(
                    "Для plate-строки нужна ссылка на плиту СГП (completed_plate_id)",
                    code="shipment_invalid_item",
                )
            cp = self._repo.fetch_completed_plate(cur, int(item.completed_plate_id))
            if cp is None:
                raise ShipmentError(
                    f"Плита СГП #{item.completed_plate_id} не найдена",
                    code="shipment_plate_not_found",
                )
            if cp["kp_id"] is None:
                raise ShipmentError(
                    f"Плита СГП #{item.completed_plate_id} не привязана к КП",
                    code="shipment_plate_unlinked",
                )
            self._repo.assert_kp_in_shipment_orders(
                cur,
                shipment_id,
                int(cp["kp_id"]),
                code="shipment_plate_kp_mismatch",
                detail=(
                    f"Плита СГП #{item.completed_plate_id} принадлежит КП #{cp['kp_id']}, "
                    f"которого нет в заказах рейса #{shipment_id}"
                ),
            )
            available = available_qty(
                cur, int(cp["id"]), exclude_shipment_id=shipment_id
            )
            cp_key = int(cp["id"])
            claimed[cp_key] = claimed.get(cp_key, 0) + int(item.qty)
            if claimed[cp_key] > available:
                raise ShipmentError(
                    f"Недостаточно «{cp['plate_name']}» на СГП: "
                    f"свободно {available}, требуется {claimed[cp_key]}",
                    code="shipment_no_availability",
                )
            unit, total = resolve_kp_line_weight_kg(
                {"length_m": cp["length_m"], "width_m": cp["width_m"], "qty": item.qty}
            )
            record["completed_plate_id"] = int(cp["id"])
            record["kp_id"] = int(cp["kp_id"])
            record["unit_weight_kg"] = unit
            record["weight_kg"] = item.weight_kg if item.weight_kg is not None else total
        else:
            mark = (item.mark or "").strip() or None
            record["mark"] = mark
            record["kp_id"] = int(item.kp_id) if item.kp_id is not None else None
            if record["kp_id"] is not None:
                self._repo.assert_kp_in_shipment_orders(
                    cur,
                    shipment_id,
                    int(record["kp_id"]),
                    code="shipment_pile_kp_mismatch",
                    detail=(
                        f"Свая с КП #{record['kp_id']} не входит в заказы рейса #{shipment_id}"
                    ),
                )
            catalog_weight = self._repo.pile_weight_for_mark(cur, mark)
            record["unit_weight_kg"] = catalog_weight
            if item.weight_kg is not None:
                record["weight_kg"] = item.weight_kg
            elif catalog_weight is not None:
                record["weight_kg"] = catalog_weight * item.qty
        return record

    def complete(
        self,
        shipment_id: int,
        *,
        actor: str | None = None,
        events_enabled: bool | None = None,
        export_dir: str | None = None,
    ) -> ShipmentMutationResponse:
        ensure_schema(self.db_path)
        conn = self._repo.connect()
        items: list[ShipmentItem] = []
        try:
            cur = conn.cursor()
            row = self._repo.fetch_shipment_row(cur, shipment_id)
            self._assert_in_work(row, shipment_id)
            orders = self._repo.fetch_orders(cur, [shipment_id]).get(shipment_id, [])
            if not orders:
                raise ShipmentError(
                    "У рейса нет заказов (КП) — отгрузка невозможна",
                    code="shipment_no_orders",
                )
            for order in orders:
                if not (order.ya_order_no or "").strip():
                    raise ShipmentError(
                        f"Не заполнен номер заказа (ЯР) для КП #{order.kp_id}",
                        code="shipment_missing_ya_order",
                    )
            items = self._repo.fetch_items(cur, shipment_id)
            if not items:
                raise ShipmentError(
                    "Состав рейса пуст — нечего отгружать",
                    code="shipment_no_items",
                )
            for item in items:
                if item.kp_id is None:
                    continue
                if item.item_type == ShipmentItemType.PLATE.value:
                    self._repo.assert_kp_in_shipment_orders(
                        cur,
                        shipment_id,
                        int(item.kp_id),
                        code="shipment_plate_kp_mismatch",
                        detail=(
                            f"Плита СГП #{item.completed_plate_id} принадлежит КП #{item.kp_id}, "
                            f"которого нет в заказах рейса #{shipment_id}"
                        ),
                    )
                else:
                    self._repo.assert_kp_in_shipment_orders(
                        cur,
                        shipment_id,
                        int(item.kp_id),
                        code="shipment_pile_kp_mismatch",
                        detail=(
                            f"Свая с КП #{item.kp_id} не входит в заказы рейса #{shipment_id}"
                        ),
                    )
            plate_items = [
                item for item in items if item.item_type == ShipmentItemType.PLATE.value
            ]
            cp_rows = self._preflight_availability(cur, shipment_id, plate_items)

            for item in plate_items:
                self._ship_plate_item(
                    cur,
                    item=item,
                    cp_row=cp_rows[int(item.completed_plate_id)],
                    shipment_id=shipment_id,
                    actor=actor,
                )

            completed_at = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
            self._repo.mark_done(
                cur,
                shipment_id,
                completed_at=completed_at,
                actor=actor,
                status=ShipmentStatus.DONE.value,
            )

            kp_ids = {order.kp_id for order in orders if order.kp_id is not None}
            kp_ids |= {item.kp_id for item in plate_items if item.kp_id is not None}
            for kp_id in sorted(kp_ids):
                if not self._check_kp_done_on_cursor(cur, int(kp_id)):
                    check_and_update_kp_completion(
                        int(kp_id), self.db_path, _external_conn=conn
                    )
            conn.commit()
        except ShipmentError:
            conn.rollback()
            raise
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

        if self._maybe_write_event is not None:
            self._maybe_write_event(
                shipment_id, events_enabled=events_enabled, export_dir=export_dir
            )
        return ShipmentMutationResponse(
            ok=True,
            shipment_id=int(shipment_id),
            status=ShipmentStatus.DONE.value,
            message=f"Рейс #{shipment_id} обработан: СГП списан, отгружено {len(items)} позиций",
        )

    def _preflight_availability(
        self,
        cur: sqlite3.Cursor,
        shipment_id: int,
        plate_items: list[ShipmentItem],
    ) -> dict[int, sqlite3.Row]:
        cp_rows: dict[int, sqlite3.Row] = {}
        claimed: dict[int, int] = {}
        for item in plate_items:
            cp_id = int(item.completed_plate_id)
            if cp_id not in cp_rows:
                cp_row = self._repo.fetch_completed_plate_for_ship(cur, cp_id)
                if cp_row is None:
                    raise ShipmentError(
                        f"Плита СГП #{cp_id} не найдена на складе",
                        code="shipment_no_availability",
                    )
                cp_rows[cp_id] = cp_row
            available = available_qty(cur, cp_id, exclude_shipment_id=shipment_id)
            claimed[cp_id] = claimed.get(cp_id, 0) + item.qty
            if claimed[cp_id] > available:
                raise ShipmentError(
                    f"Недостаточно «{cp_rows[cp_id]['plate_name']}» на СГП: "
                    f"свободно {available}, требуется {claimed[cp_id]}",
                    code="shipment_no_availability",
                )
        return cp_rows

    @staticmethod
    def _ship_plate_item(
        cur: sqlite3.Cursor,
        *,
        item: ShipmentItem,
        cp_row: sqlite3.Row,
        shipment_id: int,
        actor: str | None,
    ) -> None:
        cur.execute(
            "UPDATE completed_plates SET qty = qty - ? WHERE id = ?",
            (item.qty, int(cp_row["id"])),
        )
        audit_append(
            cur,
            plate_id=int(cp_row["id"]),
            kp_id=int(item.kp_id),
            plate_name=str(cp_row["plate_name"] or ""),
            plan_id=cp_row["plan_id"],
            day_number=cp_row["production_day"],
            from_status=PlateStatus.ON_SGP.value,
            to_status=PlateStatus.SHIPPED.value,
            qty=item.qty,
            reason=PlateTransitionReason.SGP_SHIP.value,
            actor=actor,
            shipment_id=int(shipment_id),
        )

    @staticmethod
    def _check_kp_done_on_cursor(cur: sqlite3.Cursor, kp_id: int) -> bool:
        freeze_ordered_qty_if_needed(cur, kp_id)
        cur.execute("SELECT ordered_qty FROM kp_meta WHERE kp_id = ?", (kp_id,))
        meta = cur.fetchone()
        if meta is None or meta[0] is None:
            return False
        ordered = int(meta[0])
        if ordered <= 0:
            return False
        if shipped_qty_for_kp(cur, kp_id) < ordered:
            return False
        cur.execute("SELECT COALESCE(SUM(qty), 0) FROM kp_plates WHERE kp_id = ?", (kp_id,))
        if int(cur.fetchone()[0] or 0) != 0:
            return False
        cur.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM completed_plates WHERE kp_id = ?",
            (kp_id,),
        )
        if int(cur.fetchone()[0] or 0) != 0:
            return False
        cur.execute(
            "UPDATE kp_meta SET status = ? WHERE kp_id = ?",
            (KpStatus.DONE.value, kp_id),
        )
        return True

    def cancel(
        self,
        shipment_id: int,
        *,
        actor: str | None = None,
    ) -> ShipmentMutationResponse:
        ensure_schema(self.db_path)
        conn = self._repo.connect()
        try:
            cur = conn.cursor()
            row = self._repo.fetch_shipment_row(cur, shipment_id)
            self._assert_in_work(row, shipment_id)
            self._repo.delete_shipment(cur, shipment_id)
            conn.commit()
            return ShipmentMutationResponse(
                ok=True,
                shipment_id=int(shipment_id),
                status="cancelled",
                message=f"Рейс #{shipment_id} отменён, состав освобождён",
            )
        except ShipmentError:
            conn.rollback()
            raise
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    @staticmethod
    def _assert_in_work(row: sqlite3.Row, shipment_id: int) -> None:
        if row["status"] != ShipmentStatus.IN_WORK.value:
            raise ShipmentError(
                f"Рейс #{shipment_id} уже обработан — изменения недоступны",
                code="shipment_not_in_work",
            )
