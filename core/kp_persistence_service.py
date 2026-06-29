"""Commercial offer (KP) save orchestration (A2 offers slice)."""

from __future__ import annotations

from typing import Dict, List, Optional

from core.kp_db_common import DEFAULT_DB, _connect


class KpPersistenceService:
    """Persists a new commercial offer and its plate lines to SQLite."""

    @staticmethod
    def save_kp_to_db(
        creation_date: str,
        order_data: List[Dict],
        xlsx_file_path: Optional[str] = None,
        customer_name: Optional[str] = None,
        manager_name: Optional[str] = None,
        discount_percent: float = 0,
        delivery_conditions: Optional[str] = None,
        payment_conditions: Optional[str] = None,
        execution_terms: Optional[str] = None,
        status: str = "в работе",
        logistics_cost: float = 0.0,
        owner_user_id: int | None = None,
        db_path: str = DEFAULT_DB,
    ) -> int:
        trip_logistics = max(0.0, float(logistics_cost or 0.0))
        try:
            from core.commercial_offer_xlsx import calculate_total_cost

            totals = calculate_total_cost(
                order_data, discount_percent, logistics_cost=trip_logistics
            )
            subtotal = totals["subtotal"]
            vat_amount = totals["vat_amount"]
            total_amount = totals["total_with_vat"]
        except ImportError:
            from core.commercial_pricing import VAT_RATE

            subtotal = 0.0
            for item in order_data:
                qty = item.get("qty", 0)
                unit_price = item.get("unit_price", 0.0)
                discounted_price = unit_price * (1 - discount_percent / 100)
                subtotal += discounted_price * qty
            vat_amount = round(subtotal * VAT_RATE, 2)
            total_amount = round(subtotal + vat_amount, 2)

        conn = _connect(db_path)
        try:
            conn.execute("PRAGMA foreign_keys = ON")
            cur = conn.cursor()

            cur.execute(
                """
                INSERT INTO KP_offers (
                    creation_date, customer_name, manager_name, discount_percent,
                    subtotal, vat_amount, total_amount,
                    delivery_conditions, payment_conditions, execution_terms, logistics_cost
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    creation_date,
                    customer_name,
                    manager_name,
                    discount_percent,
                    subtotal,
                    vat_amount,
                    total_amount,
                    delivery_conditions,
                    payment_conditions,
                    execution_terms,
                    trip_logistics,
                ),
            )
            kp_id = cur.lastrowid

            from core.concrete_grade_resolver import resolve_concrete_grade_from_order
            from core.db_config import PB_DB_PATH

            for idx, item in enumerate(order_data, start=1):
                qty = item.get("qty", 0)
                unit_price = item.get("unit_price", 0.0)
                discounted_price = unit_price * (1 - discount_percent / 100)
                weight = item.get("weight", 0.0)
                unit_weight = weight / qty if qty > 0 else 0.0
                plate_name = item.get("name", "")
                nomenclature_id = item.get("nomenclature_id", None)
                concrete_grade = resolve_concrete_grade_from_order(
                    {
                        "concrete_grade": item.get("concrete_grade"),
                        "plate_name": plate_name,
                        "length": item.get("length_m", 0),
                        "load_code": item.get("load_class", 800),
                    },
                    db_path=PB_DB_PATH,
                )
                cur.execute(
                    """
                    INSERT INTO kp_plates (
                        kp_id, position_number, plate_name,
                        length_m, width_m, load_class,
                        qty, unit_weight, total_weight, discounted_price, unit_price,
                        length_dm_raw, nomenclature_id, concrete_grade
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        kp_id,
                        idx,
                        plate_name,
                        item.get("length_m", 0),
                        item.get("width_m", 0),
                        item.get("load_class", 800),
                        qty,
                        unit_weight,
                        weight,
                        discounted_price,
                        unit_price,
                        item.get("length_dm_raw", "") or "",
                        nomenclature_id,
                        concrete_grade,
                    ),
                )

            if xlsx_file_path:
                from core.kp_file_paths import resolve_kp_xlsx_path_for_read

                resolved_xlsx = resolve_kp_xlsx_path_for_read(xlsx_file_path)
                if resolved_xlsx is not None:
                    with open(resolved_xlsx, "rb") as f:
                        xlsx_blob = f.read()
                    cur.execute(
                        """
                        INSERT INTO kp_files (kp_id, xlsx_file, file_path)
                        VALUES (?, ?, ?)
                        """,
                        (kp_id, xlsx_blob, str(resolved_xlsx)),
                    )

            cur.execute(
                "INSERT INTO kp_meta (kp_id, status, owner_user_id) VALUES (?, ?, ?)",
                (kp_id, status, owner_user_id),
            )
            conn.commit()
            return kp_id
        finally:
            conn.close()
