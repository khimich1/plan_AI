"""Commercial offer (KP) save orchestration (A2 offers slice)."""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from core.domain.enums import KpStatus, PlateStatus
from core.kp_db_common import DEFAULT_DB, _connect
from core.pile_trip_pricing import coerce_pile_trip_overrides, dumps_pile_trip_overrides

_VALID_PRODUCT_TYPES = frozenset(
    {
        "plates",
        "piles",
        "steps",
        "marches",
        "bridge_piles",
        "fbs",
    }
)

_PRODUCT_KIND_TO_TYPE = {
    "pile": "piles",
    "step": "steps",
    "march": "marches",
    "bridge_pile": "bridge_piles",
    "fbs": "fbs",
}

_LINE_TABLE_BY_TYPE = {
    "plates": "kp_plates",
    "piles": "kp_piles",
    "steps": "kp_steps",
    "marches": "kp_marches",
    "bridge_piles": "kp_bridge_piles",
    "fbs": "kp_fbs",
}

_STATUS_IN_WORK = KpStatus.IN_WORK.value
_PLATE_IN_PLAN = PlateStatus.IN_PLAN.value
_PLATE_IN_PRODUCTION = PlateStatus.IN_PRODUCTION.value
_PROTECTED_PLATE_STATUSES = frozenset({_PLATE_IN_PLAN, _PLATE_IN_PRODUCTION})


def _normalize_product_type_param(product_type: str | None) -> str:
    normalized = (product_type or "plates").strip().lower()
    if normalized not in _VALID_PRODUCT_TYPES:
        return "plates"
    return normalized


def _order_level_product_type(order_data: List[Dict], product_type: str) -> str:
    """Legacy mono detection from whole order (all lines same kind)."""
    from core.commercial_pricing import (
        is_bridge_pile_order,
        is_fbs_order,
        is_march_order,
        is_pile_order,
        is_step_order,
    )

    normalized = _normalize_product_type_param(product_type)
    if is_pile_order(order_data):
        return "piles"
    if is_bridge_pile_order(order_data):
        return "bridge_piles"
    if is_fbs_order(order_data):
        return "fbs"
    if is_step_order(order_data):
        return "steps"
    if is_march_order(order_data):
        return "marches"
    return normalized


def _resolve_line_product_type(item: dict[str, Any], order_fallback: str) -> str:
    """Per-line type: explicit product_type → product_kind → order fallback."""
    explicit = str(item.get("product_type") or "").strip().lower()
    if explicit in _VALID_PRODUCT_TYPES:
        return explicit
    kind = str(item.get("product_kind") or "").strip().lower()
    mapped = _PRODUCT_KIND_TO_TYPE.get(kind)
    if mapped:
        return mapped
    return order_fallback


def _meta_product_type(line_types: list[str], order_fallback: str) -> str:
    distinct = {t for t in line_types if t}
    if len(distinct) >= 2:
        return "mixed"
    if len(distinct) == 1:
        return next(iter(distinct))
    return order_fallback


def _line_id_value(item: dict[str, Any]) -> str | None:
    raw = item.get("line_id")
    if raw is None:
        return None
    text = str(raw).strip()
    return text or None


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
        product_type: str = "plates",
        db_path: str = DEFAULT_DB,
        pile_logistics_cost: float = 0.0,
        pile_trip_overrides: dict | None = None,
    ) -> int:
        trip_logistics = max(0.0, float(logistics_cost or 0.0))
        pile_trip = max(0.0, float(pile_logistics_cost or 0.0))
        try:
            from core.commercial_pricing import calculate_total_cost

            totals = calculate_total_cost(
                order_data,
                discount_percent,
                logistics_cost=trip_logistics,
                db_path=db_path,
                require_all_priced=False,
                pile_logistics_cost=pile_trip,
                pile_trip_overrides=coerce_pile_trip_overrides(pile_trip_overrides),
                pile_catalog_db_path=db_path,
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

        order_fallback = _order_level_product_type(order_data, product_type)
        line_types = [
            _resolve_line_product_type(item, order_fallback) for item in order_data
        ]
        meta_type = _meta_product_type(line_types, order_fallback)

        conn = _connect(db_path)
        try:
            conn.execute("PRAGMA foreign_keys = ON")
            cur = conn.cursor()

            cur.execute(
                """
                INSERT INTO KP_offers (
                    creation_date, customer_name, manager_name, discount_percent,
                    subtotal, vat_amount, total_amount,
                    delivery_conditions, payment_conditions, execution_terms,
                    logistics_cost, pile_logistics_cost
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                    pile_trip,
                ),
            )
            kp_id = cur.lastrowid

            for idx, (item, line_type) in enumerate(
                zip(order_data, line_types), start=1
            ):
                KpPersistenceService._insert_line(
                    cur,
                    kp_id=kp_id,
                    position_number=idx,
                    item=item,
                    line_type=line_type,
                    discount_percent=discount_percent,
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
                """
                INSERT INTO kp_meta (
                    kp_id, status, owner_user_id, product_type, pile_trip_overrides_json
                ) VALUES (?, ?, ?, ?, ?)
                """,
                (
                    kp_id,
                    status,
                    owner_user_id,
                    meta_type,
                    dumps_pile_trip_overrides(pile_trip_overrides),
                ),
            )
            conn.commit()
            return kp_id
        finally:
            conn.close()

    @staticmethod
    def update_kp_from_order_data(
        kp_id: int,
        order_data: List[Dict],
        xlsx_file_path: Optional[str] = None,
        customer_name: Optional[str] = None,
        manager_name: Optional[str] = None,
        discount_percent: float | None = None,
        delivery_conditions: Optional[str] = None,
        payment_conditions: Optional[str] = None,
        execution_terms: Optional[str] = None,
        logistics_cost: float | None = None,
        product_type: str = "plates",
        db_path: str = DEFAULT_DB,
        pile_logistics_cost: float | None = None,
        pile_trip_overrides: dict | None = None,
    ) -> int:
        """Sync existing KP lines by ``line_id`` (append/update; same ``kp_id``).

        Allowed only when ``kp_meta.status == «в работе»``. Every incoming line
        must carry a non-empty ``line_id``. Matching ``line_id`` updates in place
        (preserves ``kp_plates.id`` and production fields); new ids INSERT;
        missing ids DELETE (plates «в плане» / «в производстве» or with
        ``plan_id`` are blocked).
        """
        conn = _connect(db_path)
        try:
            conn.execute("PRAGMA foreign_keys = ON")
            cur = conn.cursor()

            cur.execute("SELECT kp_id FROM KP_offers WHERE kp_id = ?", (kp_id,))
            if cur.fetchone() is None:
                raise ValueError(f"КП {kp_id} не найдено.")

            cur.execute(
                "SELECT status FROM kp_meta WHERE kp_id = ?",
                (kp_id,),
            )
            meta_row = cur.fetchone()
            current_status = (
                str(meta_row[0]) if meta_row and meta_row[0] is not None else _STATUS_IN_WORK
            )
            if current_status != _STATUS_IN_WORK:
                raise ValueError(
                    "Обновление КП разрешено только в статусе «в работе»."
                )

            cur.execute(
                """
                SELECT discount_percent, COALESCE(logistics_cost, 0),
                       COALESCE(pile_logistics_cost, 0),
                       customer_name, manager_name,
                       delivery_conditions, payment_conditions, execution_terms
                FROM KP_offers WHERE kp_id = ?
                """,
                (kp_id,),
            )
            offer_row = cur.fetchone()
            assert offer_row is not None
            existing_discount = float(offer_row[0] or 0.0)
            existing_logistics = max(0.0, float(offer_row[1] or 0.0))
            existing_pile_logistics = max(0.0, float(offer_row[2] or 0.0))
            resolved_discount = (
                existing_discount
                if discount_percent is None
                else float(discount_percent)
            )
            trip_logistics = (
                existing_logistics
                if logistics_cost is None
                else max(0.0, float(logistics_cost or 0.0))
            )
            pile_trip = (
                existing_pile_logistics
                if pile_logistics_cost is None
                else max(0.0, float(pile_logistics_cost or 0.0))
            )
            resolved_customer = (
                customer_name if customer_name is not None else offer_row[3]
            )
            resolved_manager = (
                manager_name if manager_name is not None else offer_row[4]
            )
            resolved_delivery = (
                delivery_conditions
                if delivery_conditions is not None
                else offer_row[5]
            )
            resolved_payment = (
                payment_conditions
                if payment_conditions is not None
                else offer_row[6]
            )
            resolved_execution = (
                execution_terms if execution_terms is not None else offer_row[7]
            )
            cur.execute(
                "SELECT pile_trip_overrides_json FROM kp_meta WHERE kp_id = ?",
                (kp_id,),
            )
            meta_overrides_row = cur.fetchone()
            existing_overrides = coerce_pile_trip_overrides(
                meta_overrides_row[0] if meta_overrides_row else None
            )
            resolved_overrides = (
                coerce_pile_trip_overrides(pile_trip_overrides)
                if pile_trip_overrides is not None
                else existing_overrides
            )

            try:
                from core.commercial_pricing import calculate_total_cost

                totals = calculate_total_cost(
                    order_data,
                    resolved_discount,
                    logistics_cost=trip_logistics,
                    db_path=db_path,
                    require_all_priced=False,
                    pile_logistics_cost=pile_trip,
                    pile_trip_overrides=resolved_overrides,
                    pile_catalog_db_path=db_path,
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
                    discounted_price = unit_price * (1 - resolved_discount / 100)
                    subtotal += discounted_price * qty
                vat_amount = round(subtotal * VAT_RATE, 2)
                total_amount = round(subtotal + vat_amount, 2)

            order_fallback = _order_level_product_type(order_data, product_type)
            line_types = [
                _resolve_line_product_type(item, order_fallback) for item in order_data
            ]
            meta_type = _meta_product_type(line_types, order_fallback)

            for item in order_data:
                if not _line_id_value(item):
                    raise ValueError(
                        "Обновление КП требует line_id у каждой позиции order_data."
                    )

            existing_by_line = KpPersistenceService._load_existing_lines_by_line_id(
                cur, kp_id
            )
            incoming_line_ids: set[str] = set()

            for idx, (item, line_type) in enumerate(
                zip(order_data, line_types), start=1
            ):
                line_id = _line_id_value(item)
                assert line_id is not None  # validated above
                incoming_line_ids.add(line_id)
                existing = existing_by_line.get(line_id)
                if existing is not None:
                    existing_type, row_id = existing
                    if existing_type != line_type:
                        # Type change: delete old row, insert as new type.
                        KpPersistenceService._delete_line_row(
                            cur,
                            line_type=existing_type,
                            row_id=row_id,
                            protect_planned_plates=True,
                        )
                        KpPersistenceService._insert_line(
                            cur,
                            kp_id=kp_id,
                            position_number=idx,
                            item=item,
                            line_type=line_type,
                            discount_percent=resolved_discount,
                        )
                    else:
                        KpPersistenceService._update_line(
                            cur,
                            row_id=row_id,
                            position_number=idx,
                            item=item,
                            line_type=line_type,
                            discount_percent=resolved_discount,
                        )
                else:
                    KpPersistenceService._insert_line(
                        cur,
                        kp_id=kp_id,
                        position_number=idx,
                        item=item,
                        line_type=line_type,
                        discount_percent=resolved_discount,
                    )

            for line_id, (existing_type, row_id) in existing_by_line.items():
                if line_id in incoming_line_ids:
                    continue
                KpPersistenceService._delete_line_row(
                    cur,
                    line_type=existing_type,
                    row_id=row_id,
                    protect_planned_plates=True,
                )

            # Orphan rows without line_id (legacy) — remove if not planned.
            KpPersistenceService._delete_orphan_lines_without_line_id(cur, kp_id)

            cur.execute(
                """
                UPDATE KP_offers SET
                    customer_name = ?,
                    manager_name = ?,
                    discount_percent = ?,
                    subtotal = ?,
                    vat_amount = ?,
                    total_amount = ?,
                    delivery_conditions = ?,
                    payment_conditions = ?,
                    execution_terms = ?,
                    logistics_cost = ?,
                    pile_logistics_cost = ?
                WHERE kp_id = ?
                """,
                (
                    resolved_customer,
                    resolved_manager,
                    resolved_discount,
                    subtotal,
                    vat_amount,
                    total_amount,
                    resolved_delivery,
                    resolved_payment,
                    resolved_execution,
                    trip_logistics,
                    pile_trip,
                    kp_id,
                ),
            )
            cur.execute(
                "UPDATE kp_meta SET pile_trip_overrides_json = ? WHERE kp_id = ?",
                (dumps_pile_trip_overrides(resolved_overrides), kp_id),
            )
            cur.execute(
                "UPDATE kp_meta SET product_type = ? WHERE kp_id = ?",
                (meta_type, kp_id),
            )

            if xlsx_file_path:
                from core.kp_file_paths import resolve_kp_xlsx_path_for_read

                resolved_xlsx = resolve_kp_xlsx_path_for_read(xlsx_file_path)
                if resolved_xlsx is not None:
                    with open(resolved_xlsx, "rb") as f:
                        xlsx_blob = f.read()
                    cur.execute(
                        "SELECT id FROM kp_files WHERE kp_id = ?",
                        (kp_id,),
                    )
                    if cur.fetchone():
                        cur.execute(
                            """
                            UPDATE kp_files
                            SET xlsx_file = ?, file_path = ?
                            WHERE kp_id = ?
                            """,
                            (xlsx_blob, str(resolved_xlsx), kp_id),
                        )
                    else:
                        cur.execute(
                            """
                            INSERT INTO kp_files (kp_id, xlsx_file, file_path)
                            VALUES (?, ?, ?)
                            """,
                            (kp_id, xlsx_blob, str(resolved_xlsx)),
                        )

            conn.commit()
            return kp_id
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    @staticmethod
    def _load_existing_lines_by_line_id(
        cur: Any, kp_id: int
    ) -> dict[str, tuple[str, int]]:
        """Map ``line_id`` → ``(product_type, row_id)`` across all kp_* tables."""
        result: dict[str, tuple[str, int]] = {}
        for line_type, table in _LINE_TABLE_BY_TYPE.items():
            cur.execute(
                f"SELECT id, line_id FROM {table} WHERE kp_id = ?",
                (kp_id,),
            )
            for row_id, line_id in cur.fetchall():
                text = str(line_id).strip() if line_id is not None else ""
                if not text:
                    continue
                result[text] = (line_type, int(row_id))
        return result

    @staticmethod
    def _delete_line_row(
        cur: Any,
        *,
        line_type: str,
        row_id: int,
        protect_planned_plates: bool,
    ) -> None:
        table = _LINE_TABLE_BY_TYPE[line_type]
        if line_type == "plates" and protect_planned_plates:
            cur.execute(
                "SELECT status, plan_id FROM kp_plates WHERE id = ?",
                (row_id,),
            )
            row = cur.fetchone()
            if row is not None:
                status = str(row[0] or "")
                plan_id = row[1]
                if status in _PROTECTED_PLATE_STATUSES or (
                    plan_id is not None and str(plan_id).strip()
                ):
                    raise ValueError(
                        "Нельзя удалить позицию плиты со статусом "
                        "«в плане» / «в производстве» или привязанную к плану."
                    )
        cur.execute(f"DELETE FROM {table} WHERE id = ?", (row_id,))

    @staticmethod
    def _delete_orphan_lines_without_line_id(cur: Any, kp_id: int) -> None:
        for line_type, table in _LINE_TABLE_BY_TYPE.items():
            cur.execute(
                f"""
                SELECT id FROM {table}
                WHERE kp_id = ? AND (line_id IS NULL OR TRIM(line_id) = '')
                """,
                (kp_id,),
            )
            for (row_id,) in cur.fetchall():
                KpPersistenceService._delete_line_row(
                    cur,
                    line_type=line_type,
                    row_id=int(row_id),
                    protect_planned_plates=True,
                )

    @staticmethod
    def _update_line(
        cur: Any,
        *,
        row_id: int,
        position_number: int,
        item: dict[str, Any],
        line_type: str,
        discount_percent: float,
    ) -> None:
        qty = int(item.get("qty", 0) or 0)
        unit_price = float(item.get("unit_price", 0.0) or 0.0)
        discounted_price = unit_price * (1 - discount_percent / 100)
        line_id = _line_id_value(item)

        if line_type == "piles":
            mark = str(item.get("mark") or item.get("name") or "").strip()
            concrete_grade = str(item.get("concrete_grade") or "B25").strip()
            cur.execute(
                """
                UPDATE kp_piles SET
                    position_number = ?, mark = ?, concrete_grade = ?,
                    qty = ?, unit_price = ?, discounted_price = ?, line_id = ?
                WHERE id = ?
                """,
                (
                    position_number,
                    mark,
                    concrete_grade,
                    qty,
                    unit_price,
                    discounted_price,
                    line_id,
                    row_id,
                ),
            )
            return

        if line_type == "bridge_piles":
            mark = str(item.get("mark") or item.get("name") or "").strip()
            concrete_grade = str(item.get("concrete_grade") or "B25").strip()
            cur.execute(
                """
                UPDATE kp_bridge_piles SET
                    position_number = ?, mark = ?, concrete_grade = ?,
                    qty = ?, unit_price = ?, discounted_price = ?, line_id = ?
                WHERE id = ?
                """,
                (
                    position_number,
                    mark,
                    concrete_grade,
                    qty,
                    unit_price,
                    discounted_price,
                    line_id,
                    row_id,
                ),
            )
            return

        if line_type == "fbs":
            mark = str(item.get("mark") or item.get("name") or "").strip()
            concrete_grade = str(item.get("concrete_grade") or "B25").strip()
            cur.execute(
                """
                UPDATE kp_fbs SET
                    position_number = ?, mark = ?, concrete_grade = ?,
                    qty = ?, unit_price = ?, discounted_price = ?, line_id = ?
                WHERE id = ?
                """,
                (
                    position_number,
                    mark,
                    concrete_grade,
                    qty,
                    unit_price,
                    discounted_price,
                    line_id,
                    row_id,
                ),
            )
            return

        if line_type == "marches":
            mark = str(item.get("mark") or item.get("name") or "").strip()
            concrete_grade = str(item.get("concrete_grade") or "B25").strip()
            cur.execute(
                """
                UPDATE kp_marches SET
                    position_number = ?, mark = ?, concrete_grade = ?,
                    qty = ?, unit_price = ?, discounted_price = ?, line_id = ?
                WHERE id = ?
                """,
                (
                    position_number,
                    mark,
                    concrete_grade,
                    qty,
                    unit_price,
                    discounted_price,
                    line_id,
                    row_id,
                ),
            )
            return

        if line_type == "steps":
            mark = str(item.get("mark") or item.get("name") or "").strip()
            cur.execute(
                """
                UPDATE kp_steps SET
                    position_number = ?, mark = ?,
                    qty = ?, unit_price = ?, discounted_price = ?, line_id = ?
                WHERE id = ?
                """,
                (
                    position_number,
                    mark,
                    qty,
                    unit_price,
                    discounted_price,
                    line_id,
                    row_id,
                ),
            )
            return

        # plates — update commercial fields only; keep status / plan_id / id
        from core.concrete_grade_resolver import resolve_concrete_grade_from_order
        from core.db_config import PB_DB_PATH

        weight = float(item.get("weight", 0.0) or 0.0)
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
            UPDATE kp_plates SET
                position_number = ?,
                plate_name = ?,
                length_m = ?,
                width_m = ?,
                load_class = ?,
                qty = ?,
                unit_weight = ?,
                total_weight = ?,
                discounted_price = ?,
                unit_price = ?,
                length_dm_raw = ?,
                nomenclature_id = ?,
                concrete_grade = ?,
                line_id = ?
            WHERE id = ?
            """,
            (
                position_number,
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
                line_id,
                row_id,
            ),
        )

    @staticmethod
    def _insert_line(
        cur: Any,
        *,
        kp_id: int,
        position_number: int,
        item: dict[str, Any],
        line_type: str,
        discount_percent: float,
    ) -> None:
        qty = int(item.get("qty", 0) or 0)
        unit_price = float(item.get("unit_price", 0.0) or 0.0)
        discounted_price = unit_price * (1 - discount_percent / 100)
        line_id = _line_id_value(item)

        if line_type == "piles":
            mark = str(item.get("mark") or item.get("name") or "").strip()
            concrete_grade = str(item.get("concrete_grade") or "B25").strip()
            cur.execute(
                """
                INSERT INTO kp_piles (
                    kp_id, position_number, mark, concrete_grade,
                    qty, unit_price, discounted_price, line_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    kp_id,
                    position_number,
                    mark,
                    concrete_grade,
                    qty,
                    unit_price,
                    discounted_price,
                    line_id,
                ),
            )
            return

        if line_type == "bridge_piles":
            mark = str(item.get("mark") or item.get("name") or "").strip()
            concrete_grade = str(item.get("concrete_grade") or "B25").strip()
            cur.execute(
                """
                INSERT INTO kp_bridge_piles (
                    kp_id, position_number, mark, concrete_grade,
                    qty, unit_price, discounted_price, line_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    kp_id,
                    position_number,
                    mark,
                    concrete_grade,
                    qty,
                    unit_price,
                    discounted_price,
                    line_id,
                ),
            )
            return

        if line_type == "fbs":
            mark = str(item.get("mark") or item.get("name") or "").strip()
            concrete_grade = str(item.get("concrete_grade") or "B25").strip()
            cur.execute(
                """
                INSERT INTO kp_fbs (
                    kp_id, position_number, mark, concrete_grade,
                    qty, unit_price, discounted_price, line_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    kp_id,
                    position_number,
                    mark,
                    concrete_grade,
                    qty,
                    unit_price,
                    discounted_price,
                    line_id,
                ),
            )
            return

        if line_type == "marches":
            mark = str(item.get("mark") or item.get("name") or "").strip()
            concrete_grade = str(item.get("concrete_grade") or "B25").strip()
            cur.execute(
                """
                INSERT INTO kp_marches (
                    kp_id, position_number, mark, concrete_grade,
                    qty, unit_price, discounted_price, line_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    kp_id,
                    position_number,
                    mark,
                    concrete_grade,
                    qty,
                    unit_price,
                    discounted_price,
                    line_id,
                ),
            )
            return

        if line_type == "steps":
            mark = str(item.get("mark") or item.get("name") or "").strip()
            cur.execute(
                """
                INSERT INTO kp_steps (
                    kp_id, position_number, mark,
                    qty, unit_price, discounted_price, line_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    kp_id,
                    position_number,
                    mark,
                    qty,
                    unit_price,
                    discounted_price,
                    line_id,
                ),
            )
            return

        # plates (default)
        from core.concrete_grade_resolver import resolve_concrete_grade_from_order
        from core.db_config import PB_DB_PATH

        weight = float(item.get("weight", 0.0) or 0.0)
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
                length_dm_raw, nomenclature_id, concrete_grade, line_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                kp_id,
                position_number,
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
                line_id,
            ),
        )
