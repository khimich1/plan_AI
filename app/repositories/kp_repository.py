from __future__ import annotations

import sqlite3
from collections import defaultdict
from collections.abc import Sequence
from datetime import datetime

from app.core.settings import get_settings
from app.domain.enums import PlateStatus
from app.repositories.kp_offers_repository import KpOffersRepository
from core.kp import offers_write


class KpRepository:
    def __init__(self, db_path: str | None = None) -> None:
        settings = get_settings()
        self.db_path = db_path or str(settings.plita_db_path)
        self._offers = KpOffersRepository(self.db_path)

    def save_offer(
        self,
        *,
        customer_name: str,
        manager_name: str,
        creation_date: str | None = None,
        discount_percent: float = 0.0,
        logistics_cost: float = 0.0,
        delivery_conditions: str = "",
        payment_conditions: str = "",
        execution_terms: str = "",
        status: str = "в работе",
        order_data: Sequence[dict] | None = None,
        xlsx_path: str | None = None,
        owner_user_id: int | None = None,
        product_type: str = "plates",
        pile_logistics_cost: float = 0.0,
        pile_trip_overrides: dict | None = None,
    ) -> int:
        return offers_write.save_kp_to_db(
            creation_date=creation_date or datetime.now().strftime("%d.%m.%Y"),
            order_data=list(order_data or []),
            xlsx_file_path=xlsx_path,
            customer_name=customer_name,
            manager_name=manager_name,
            discount_percent=discount_percent,
            logistics_cost=logistics_cost,
            delivery_conditions=delivery_conditions,
            payment_conditions=payment_conditions,
            execution_terms=execution_terms,
            status=status,
            owner_user_id=owner_user_id,
            product_type=product_type,
            db_path=self.db_path,
            pile_logistics_cost=pile_logistics_cost,
            pile_trip_overrides=pile_trip_overrides,
        )

    def update_offer_from_order_data(
        self,
        kp_id: int,
        order_data: Sequence[dict] | None = None,
        *,
        customer_name: str | None = None,
        manager_name: str | None = None,
        discount_percent: float | None = None,
        logistics_cost: float | None = None,
        delivery_conditions: str | None = None,
        payment_conditions: str | None = None,
        execution_terms: str | None = None,
        xlsx_path: str | None = None,
        product_type: str = "plates",
        pile_logistics_cost: float | None = None,
        pile_trip_overrides: dict | None = None,
    ) -> int:
        """Append/update existing KP by ``line_id`` (same ``kp_id``)."""
        return offers_write.update_kp_from_order_data(
            kp_id,
            list(order_data or []),
            xlsx_file_path=xlsx_path,
            customer_name=customer_name,
            manager_name=manager_name,
            discount_percent=discount_percent,
            logistics_cost=logistics_cost,
            delivery_conditions=delivery_conditions,
            payment_conditions=payment_conditions,
            execution_terms=execution_terms,
            product_type=product_type,
            db_path=self.db_path,
            pile_logistics_cost=pile_logistics_cost,
            pile_trip_overrides=pile_trip_overrides,
        )

    def list_offers_grouped(self, **list_filters) -> dict[str, list[dict]]:
        return self._offers.list_grouped(**list_filters)

    def list_offers(self, limit: int = 100) -> list[dict]:
        query = """
        SELECT o.kp_id, o.creation_date, o.customer_name, o.manager_name,
               o.discount_percent, o.subtotal, o.vat_amount, o.total_amount,
               COALESCE(m.status, 'в работе') AS status
        FROM KP_offers o
        LEFT JOIN kp_meta m ON m.kp_id = o.kp_id
        ORDER BY o.kp_id DESC
        LIMIT ?
        """
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query, (limit,))
            return [dict(row) for row in cursor.fetchall()]

    def get_offer(self, kp_id: int) -> dict | None:
        return self._offers.get_by_id(kp_id)

    def update_offer_discount(self, kp_id: int, discount_percent: float) -> bool:
        return offers_write.update_kp_discount(kp_id, discount_percent, self.db_path)

    def update_offer_logistics_cost(self, kp_id: int, logistics_cost: float) -> bool:
        return offers_write.update_kp_logistics_cost(kp_id, logistics_cost, self.db_path)

    def update_offer_status(self, kp_id: int, status: str) -> bool:
        return offers_write.update_kp_status(kp_id, status, self.db_path)

    def update_offer_execution_date(self, kp_id: int, execution_date: str) -> bool:
        return offers_write.update_kp_execution_date(kp_id, execution_date, self.db_path)

    def delete_offer(self, kp_id: int) -> bool:
        return offers_write.delete_kp_by_id(kp_id, self.db_path)

    def get_completion_percentage(self, kp_id: int) -> dict:
        return self._offers.get_completion_percentage(kp_id)

    def list_production_candidates(self, limit: int = 500) -> list[dict]:
        query = """
        SELECT kp_id, id, plate_name, length_m, width_m, load_class, qty,
               status, plan_id, length_dm_raw, nomenclature_id
        FROM kp_plates
        WHERE status IN ('в производстве', 'в плане')
        ORDER BY kp_id DESC, id ASC
        LIMIT ?
        """
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query, (limit,))
            return [dict(row) for row in cursor.fetchall()]

    def list_kps_in_production(self) -> list[dict]:
        """Возвращает список КП со статусом 'в работе' с метриками выполнения."""
        # Include mono plates and mixed-with-plates; exclude non-plate KPs
        # (piles/FBS/etc.) via presence of kp_plates rows, not product_type alone.
        query = """
        SELECT o.kp_id, o.customer_name, o.creation_date, o.execution_terms
        FROM KP_offers o
        JOIN kp_meta m ON m.kp_id = o.kp_id
        WHERE m.status = 'в работе'
          AND EXISTS (
              SELECT 1 FROM kp_plates p WHERE p.kp_id = o.kp_id
          )
        ORDER BY o.kp_id ASC
        """
        result: list[dict] = []
        plates_by_kp: dict[int, list[dict]] = defaultdict(list)
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query)
            rows = [dict(row) for row in cursor.fetchall()]

            if rows:
                kp_ids = [int(r["kp_id"]) for r in rows]
                placeholders = ",".join("?" * len(kp_ids))
                plates_query = f"""
                SELECT kp_id, id, plate_name, length_m, width_m, load_class, qty
                FROM kp_plates
                WHERE status = ? AND kp_id IN ({placeholders})
                ORDER BY kp_id ASC, position_number, id
                """
                cursor.execute(
                    plates_query,
                    (PlateStatus.IN_PRODUCTION.value, *kp_ids),
                )
                for plate_row in cursor.fetchall():
                    plate = dict(plate_row)
                    plates_by_kp[int(plate["kp_id"])].append(
                        {
                            "id": int(plate["id"]),
                            "plate_name": plate.get("plate_name") or "",
                            "length_m": float(plate.get("length_m") or 0.0),
                            "width_m": float(plate.get("width_m") or 0.0),
                            "load_class": (
                                int(plate["load_class"])
                                if plate.get("load_class") is not None
                                else None
                            ),
                            "qty": int(plate.get("qty") or 0),
                        }
                    )

        for row in rows:
            kp_id = int(row["kp_id"])
            completion = self._offers.get_completion_percentage(kp_id)
            in_plan = self._offers.get_plates_in_plan_percentage(kp_id)
            total_length_m = self._offers.get_total_length(kp_id)

            result.append(
                {
                    "kp_id": kp_id,
                    "customer_name": row.get("customer_name") or "",
                    "creation_date": row.get("creation_date") or "",
                    "execution_terms": row.get("execution_terms") or "",
                    "total_plates": int(completion.get("total_plates", 0)),
                    "completed_plates": int(completion.get("completed_plates", 0)),
                    "completion_pct": float(completion.get("percentage", 0.0)),
                    "in_plan_pct": float(in_plan.get("percentage", 0.0)),
                    "total_length_m": round(float(total_length_m), 2),
                    "plates": plates_by_kp.get(kp_id, []),
                }
            )
        return result

    def get_plate_qty_remaining(self, plate_id: int) -> int:
        """Unplanned remaining qty for a single ``kp_plates`` row.

        Formula (per-row)::

            qty − Σ(qty WHERE id=plate_id AND status='в плане' AND plan_id IS NOT NULL)

        After the split model this is the row ``qty`` when the plate is still in
        production or stuck ``в плане`` without ``plan_id``; actively planned
        rows (``в плане`` + ``plan_id``) yield ``0``. Missing ``plate_id``
        returns ``0``.
        """
        query = """
        SELECT
            p.qty - COALESCE((
                SELECT SUM(q.qty)
                FROM kp_plates q
                WHERE q.id = p.id
                  AND q.status = ?
                  AND q.plan_id IS NOT NULL
            ), 0) AS qty_remaining
        FROM kp_plates p
        WHERE p.id = ?
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(query, (PlateStatus.IN_PLAN.value, plate_id))
            row = cursor.fetchone()
            if row is None:
                return 0
            return int(row[0] or 0)

    def list_delivery_batch_items_for_in_production_plates(self) -> list[dict]:
        """Cross-KP delivery batch lines for plates still ``в производстве``.

        Returns dicts with keys ``plate_id``, ``produce_by``, ``qty``,
        ``batch_name``. SQL stays in the repository so services stay I/O-thin.
        """
        query = """
        SELECT
            i.plate_id AS plate_id,
            b.produce_by AS produce_by,
            i.qty AS qty,
            b.name AS batch_name
        FROM delivery_batch_item i
        JOIN delivery_batch b ON b.id = i.batch_id
        JOIN kp_plates p ON p.id = i.plate_id
        WHERE p.status = ?
        ORDER BY i.plate_id ASC, b.produce_by ASC, i.id ASC
        """
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query, (PlateStatus.IN_PRODUCTION.value,))
            return [
                {
                    "plate_id": int(row["plate_id"]),
                    "produce_by": row["produce_by"],
                    "qty": int(row["qty"]),
                    "batch_name": row["batch_name"],
                }
                for row in cursor.fetchall()
            ]
