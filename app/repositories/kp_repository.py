from __future__ import annotations

import sqlite3
from collections.abc import Sequence
from datetime import datetime

from app.core.settings import get_settings
from core import kp_db


class KpRepository:
    def __init__(self, db_path: str | None = None) -> None:
        settings = get_settings()
        self.db_path = db_path or str(settings.plita_db_path)
        kp_db.init_schema(self.db_path)

    def save_offer(
        self,
        *,
        customer_name: str,
        manager_name: str,
        creation_date: str | None = None,
        discount_percent: float = 0.0,
        delivery_conditions: str = "",
        payment_conditions: str = "",
        execution_terms: str = "",
        status: str = "в работе",
        order_data: Sequence[dict] | None = None,
        xlsx_path: str | None = None,
    ) -> int:
        return kp_db.save_kp_to_db(
            creation_date=creation_date or datetime.now().strftime("%d.%m.%Y"),
            order_data=list(order_data or []),
            xlsx_file_path=xlsx_path,
            customer_name=customer_name,
            manager_name=manager_name,
            discount_percent=discount_percent,
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

