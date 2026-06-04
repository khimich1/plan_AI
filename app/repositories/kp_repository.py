from __future__ import annotations

import sqlite3
from collections import defaultdict
from collections.abc import Sequence
from datetime import datetime

from app.core.settings import get_settings
from app.domain.enums import PlateStatus
from core import kp_db_offers


class KpRepository:
    def __init__(self, db_path: str | None = None) -> None:
        settings = get_settings()
        self.db_path = db_path or str(settings.plita_db_path)

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
    ) -> int:
        return kp_db_offers.save_kp_to_db(
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
            db_path=self.db_path,
        )

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

    def list_offers_grouped(self) -> dict[str, list[dict]]:
        return kp_db_offers.get_all_kp_list(self.db_path)

    def get_offer(self, kp_id: int) -> dict | None:
        return kp_db_offers.get_kp_by_id(kp_id, self.db_path)

    def update_offer_discount(self, kp_id: int, discount_percent: float) -> bool:
        return kp_db_offers.update_kp_discount(kp_id, discount_percent, self.db_path)

    def update_offer_logistics_cost(self, kp_id: int, logistics_cost: float) -> bool:
        return kp_db_offers.update_kp_logistics_cost(kp_id, logistics_cost, self.db_path)

    def update_offer_status(self, kp_id: int, status: str) -> bool:
        return kp_db_offers.update_kp_status(kp_id, status, self.db_path)

    def update_offer_execution_date(self, kp_id: int, execution_date: str) -> bool:
        return kp_db_offers.update_kp_execution_date(kp_id, execution_date, self.db_path)

    def delete_offer(self, kp_id: int) -> bool:
        return kp_db_offers.delete_kp_by_id(kp_id, self.db_path)

    def get_completion_percentage(self, kp_id: int) -> dict:
        return kp_db_offers.get_kp_completion_percentage(kp_id, self.db_path)

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
        query = """
        SELECT o.kp_id, o.customer_name, o.creation_date, o.execution_terms
        FROM KP_offers o
        JOIN kp_meta m ON m.kp_id = o.kp_id
        WHERE m.status = 'в работе'
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
            completion = kp_db_offers.get_kp_completion_percentage(kp_id, self.db_path)
            in_plan = kp_db_offers.get_kp_plates_in_plan_percentage(kp_id, self.db_path)
            total_length_m = kp_db_offers.get_kp_total_length(kp_id, self.db_path)

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
