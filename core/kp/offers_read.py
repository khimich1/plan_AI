"""Read/query SQL for commercial offers (KP) — extracted from kp_db_offers (A7 slice 2)."""

from __future__ import annotations

import sqlite3
from typing import Dict, List, Optional

from core.kp.plates_resolve import resolve_plates_for_kp_documents
from core.kp_db_common import DEFAULT_DB, _connect

# (kp_data key, SQL table) for typed line tables.
_KP_LINE_TABLES: tuple[tuple[str, str], ...] = (
    ("plates", "kp_plates"),
    ("piles", "kp_piles"),
    ("steps", "kp_steps"),
    ("marches", "kp_marches"),
    ("bridge_piles", "kp_bridge_piles"),
    ("fbs", "kp_fbs"),
)


def _stamp_product_type(rows: List[Dict], product_type: str) -> List[Dict]:
    """Copy rows and set product_type (DB tables have no product_type column)."""
    stamped: List[Dict] = []
    for row in rows:
        item = dict(row)
        item["product_type"] = product_type
        stamped.append(item)
    return stamped


def _fetch_typed_rows(cur: sqlite3.Cursor, table: str, kp_id: int) -> List[Dict]:
    cur.execute(
        f"""
        SELECT * FROM {table}
        WHERE kp_id = ?
        ORDER BY position_number
        """,
        (kp_id,),
    )
    return [dict(r) for r in cur.fetchall()]


def _load_plates_for_kp(
    cur: sqlite3.Cursor,
    *,
    kp_id: int,
    db_path: str,
) -> List[Dict]:
    raw_plates = _fetch_typed_rows(cur, "kp_plates", kp_id)
    return resolve_plates_for_kp_documents(
        raw_plates,
        kp_id=kp_id,
        db_path=db_path,
    )


def _empty_typed_arrays() -> Dict[str, List]:
    return {key: [] for key, _ in _KP_LINE_TABLES}


def get_kp_by_id(kp_id: int, db_path: str = DEFAULT_DB) -> Optional[Dict]:
    """
    Получает информацию о КП по порядковому номеру.

    Возвращает словарь с информацией о КП или None, если не найдено.
    Для product_type=mixed загружает все kp_* таблицы и проставляет product_type
    на каждой строке (сквозной порядок — по position_number).
    """
    conn = _connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        cur.execute("SELECT * FROM KP_offers WHERE kp_id = ?", (kp_id,))
        row = cur.fetchone()
        if not row:
            return None

        kp_data = dict(row)

        cur.execute(
            "SELECT status, owner_user_id, COALESCE(product_type, 'plates') AS product_type "
            "FROM kp_meta WHERE kp_id = ?",
            (kp_id,),
        )
        meta_row = cur.fetchone()
        if meta_row:
            kp_data["status"] = meta_row["status"]
            kp_data["owner_user_id"] = meta_row["owner_user_id"]
            kp_data["product_type"] = meta_row["product_type"] or "plates"
        else:
            kp_data["product_type"] = "plates"

        product_type = str(kp_data.get("product_type") or "plates").lower()
        if product_type == "mixed":
            kp_data.update(_empty_typed_arrays())
            for key, table in _KP_LINE_TABLES:
                if key == "plates":
                    rows = _load_plates_for_kp(cur, kp_id=kp_id, db_path=db_path)
                else:
                    rows = _fetch_typed_rows(cur, table, kp_id)
                kp_data[key] = _stamp_product_type(rows, key)
        elif product_type == "piles":
            kp_data.update(_empty_typed_arrays())
            kp_data["piles"] = _fetch_typed_rows(cur, "kp_piles", kp_id)
        elif product_type == "bridge_piles":
            kp_data.update(_empty_typed_arrays())
            kp_data["bridge_piles"] = _fetch_typed_rows(cur, "kp_bridge_piles", kp_id)
        elif product_type == "fbs":
            kp_data.update(_empty_typed_arrays())
            kp_data["fbs"] = _fetch_typed_rows(cur, "kp_fbs", kp_id)
        elif product_type == "marches":
            kp_data.update(_empty_typed_arrays())
            kp_data["marches"] = _fetch_typed_rows(cur, "kp_marches", kp_id)
        elif product_type == "steps":
            kp_data.update(_empty_typed_arrays())
            kp_data["steps"] = _fetch_typed_rows(cur, "kp_steps", kp_id)
        else:
            kp_data.update(_empty_typed_arrays())
            kp_data["plates"] = _load_plates_for_kp(cur, kp_id=kp_id, db_path=db_path)

        cur.execute("SELECT * FROM kp_files WHERE kp_id = ?", (kp_id,))
        file_row = cur.fetchone()
        if file_row:
            kp_data["file"] = dict(file_row)

        return kp_data

    finally:
        conn.close()


def get_all_kp_by_status(status: str, db_path: str = DEFAULT_DB) -> List[Dict]:
    """Получает все КП с определённым статусом."""
    conn = _connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        cur.execute("SELECT kp_id FROM kp_meta WHERE status = ?", (status,))
        kp_ids = [row["kp_id"] for row in cur.fetchall()]

        result = []
        for kp_id in kp_ids:
            kp_data = get_kp_by_id(kp_id, db_path)
            if kp_data:
                result.append(kp_data)

        return result

    finally:
        conn.close()


def get_xlsx_file(
    kp_id: int,
    output_path: Optional[str] = None,
    db_path: str = DEFAULT_DB,
) -> Optional[bytes]:
    """Извлекает файл XLSX из базы данных."""
    conn = _connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.cursor()
        cur.execute("SELECT xlsx_file FROM kp_files WHERE kp_id = ?", (kp_id,))
        row = cur.fetchone()

        if not row or not row[0]:
            return None

        xlsx_data = row[0]

        if output_path:
            from core.kp_file_paths import resolve_kp_xlsx_path_for_write

            safe_path = resolve_kp_xlsx_path_for_write(output_path)
            if safe_path is None:
                return None
            safe_path.parent.mkdir(parents=True, exist_ok=True)
            with open(safe_path, "wb") as f:
                f.write(xlsx_data)

        return xlsx_data

    finally:
        conn.close()


def get_db_stats(db_path: str = DEFAULT_DB) -> Dict[str, int]:
    """Статистика по базе данных — количество записей в каждой таблице."""
    conn = _connect(db_path)
    try:
        cur = conn.cursor()

        cur.execute("SELECT COUNT(*) FROM KP_offers")
        kp_count = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM kp_plates")
        plates_count = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM completed_plates")
        completed_count = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM plate_rests")
        rests_count = cur.fetchone()[0]

        cur.execute('SELECT COUNT(*) FROM kp_meta WHERE status = "в работе"')
        in_work_count = cur.fetchone()[0]

        cur.execute('SELECT COUNT(*) FROM kp_meta WHERE status = "выполнено"')
        completed_kp_count = cur.fetchone()[0]

        return {
            "kp_total": kp_count,
            "kp_in_work": in_work_count,
            "kp_completed": completed_kp_count,
            "plates_in_work": plates_count,
            "plates_completed": completed_count,
            "plate_rests": rests_count,
        }

    finally:
        conn.close()


def get_next_kp_number(db_path: str = DEFAULT_DB) -> int:
    """Следующий свободный номер КП (MAX(kp_id) + 1)."""
    conn = _connect(db_path)

    try:
        cur = conn.cursor()
        cur.execute("SELECT MAX(kp_id) FROM KP_offers")
        result = cur.fetchone()

        max_id = result[0] if result[0] is not None else 0
        next_id = max_id + 1

        print(f"[DB] Следующий номер КП: {next_id}")
        return next_id

    finally:
        conn.close()


_PRODUCT_TYPE_LINE_TABLE: dict[str, str] = {key: table for key, table in _KP_LINE_TABLES}


def _product_type_sql_filter(product_type: str | None) -> tuple[str, list]:
    """Optional «contains type» filter (Q3/MNA-602).

    Matches mono KP with that meta product_type **or** any KP that has rows in the
    corresponding line table (so mixed-with-plates passes product_type=plates).
    """
    if not product_type or product_type == "all":
        return "", []
    normalized = str(product_type).strip().lower()
    table = _PRODUCT_TYPE_LINE_TABLE.get(normalized)
    if not table:
        return "", []
    return (
        " AND ("
        "COALESCE(m.product_type, 'plates') = ? "
        f"OR EXISTS (SELECT 1 FROM {table} t WHERE t.kp_id = ko.kp_id)"
        ")",
        [normalized],
    )


def _attach_product_types(cur: sqlite3.Cursor, rows: List[Dict]) -> None:
    """Mutate list/search rows: set product_types for UI badges (mono or mixed)."""
    if not rows:
        return
    mixed_ids = [
        int(row["kp_id"])
        for row in rows
        if str(row.get("product_type") or "plates").lower() == "mixed"
    ]
    types_by_kp: dict[int, list[str]] = {kp_id: [] for kp_id in mixed_ids}
    if mixed_ids:
        placeholders = ",".join("?" * len(mixed_ids))
        for key, table in _KP_LINE_TABLES:
            cur.execute(
                f"SELECT DISTINCT kp_id FROM {table} WHERE kp_id IN ({placeholders})",
                mixed_ids,
            )
            for found in cur.fetchall():
                types_by_kp[int(found["kp_id"])].append(key)

    for row in rows:
        meta = str(row.get("product_type") or "plates").lower()
        if meta == "mixed":
            row["product_types"] = list(types_by_kp.get(int(row["kp_id"]), []))
        else:
            row["product_types"] = [meta]


def _offer_access_sql_filters(
    *,
    owner_user_id: int | None = None,
    readable_statuses: tuple[str, ...] | None = None,
    deny_all: bool = False,
) -> tuple[str, list]:
    """Build AND-fragment for offer list/search queries (empty → no extra filter)."""
    if deny_all:
        return " AND 1 = 0", []
    clauses: list[str] = []
    params: list = []
    if owner_user_id is not None:
        clauses.append("m.owner_user_id = ?")
        params.append(owner_user_id)
    if readable_statuses is not None:
        placeholders = ",".join("?" * len(readable_statuses))
        clauses.append(f"COALESCE(m.status, 'в работе') IN ({placeholders})")
        params.extend(readable_statuses)
    if not clauses:
        return "", []
    return " AND " + " AND ".join(clauses), params


def get_all_kp_list(
    db_path: str = DEFAULT_DB,
    *,
    owner_user_id: int | None = None,
    readable_statuses: tuple[str, ...] | None = None,
    deny_all: bool = False,
    product_type: str | None = None,
) -> Dict[str, List[Dict]]:
    """Все КП, сгруппированные по статусам: archived / in_production / completed."""
    conn = _connect(db_path)

    try:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        access_sql, access_params = _offer_access_sql_filters(
            owner_user_id=owner_user_id,
            readable_statuses=readable_statuses,
            deny_all=deny_all,
        )
        product_sql, product_params = _product_type_sql_filter(product_type)

        cur.execute(
            f"""
            SELECT
                ko.kp_id,
                ko.creation_date,
                ko.customer_name,
                ko.manager_name,
                ko.discount_percent,
                ko.subtotal,
                ko.vat_amount,
                ko.total_amount,
                ko.delivery_conditions,
                ko.payment_conditions,
                ko.execution_terms,
                m.status,
                m.owner_user_id,
                COALESCE(m.product_type, 'plates') AS product_type
            FROM KP_offers ko
            LEFT JOIN kp_meta m ON ko.kp_id = m.kp_id
            WHERE 1 = 1{access_sql}{product_sql}
            ORDER BY ko.kp_id ASC
            """,
            (*access_params, *product_params),
        )

        all_kp = [dict(row) for row in cur.fetchall()]
        _attach_product_types(cur, all_kp)

        result: Dict[str, List[Dict]] = {
            "archived": [],
            "in_production": [],
            "completed": [],
        }

        for kp in all_kp:
            status = kp.get("status", "в работе")

            if status == "в архиве":
                result["archived"].append(kp)
            elif status in ("в работе", "На СГП"):
                result["in_production"].append(kp)
            elif status == "выполнено":
                result["completed"].append(kp)

        return result

    finally:
        conn.close()


def _escape_sql_like(value: str) -> str:
    """Экранирует спецсимволы LIKE (% и _) в пользовательском вводе."""
    return value.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def search_kp_by_customer_name(
    name: str,
    limit: int = 50,
    db_path: str = DEFAULT_DB,
    *,
    owner_user_id: int | None = None,
    readable_statuses: tuple[str, ...] | None = None,
    deny_all: bool = False,
    product_type: str | None = None,
) -> tuple[List[Dict], int]:
    """Ищет КП по частичному совпадению имени заказчика."""
    conn = _connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        escaped = _escape_sql_like(name.strip())
        pattern = f"%{escaped}%"
        fetch_limit = limit + 1

        access_sql, access_params = _offer_access_sql_filters(
            owner_user_id=owner_user_id,
            readable_statuses=readable_statuses,
            deny_all=deny_all,
        )
        product_sql, product_params = _product_type_sql_filter(product_type)

        base_select = f"""
            SELECT
                ko.kp_id,
                ko.creation_date,
                ko.customer_name,
                ko.manager_name,
                ko.discount_percent,
                ko.subtotal,
                ko.vat_amount,
                ko.total_amount,
                ko.execution_terms,
                m.status,
                m.owner_user_id,
                COALESCE(m.product_type, 'plates') AS product_type
            FROM KP_offers ko
            LEFT JOIN kp_meta m ON ko.kp_id = m.kp_id
            WHERE casefold(ko.customer_name) LIKE casefold(?) ESCAPE '\\'{access_sql}{product_sql}
        """

        cur.execute(
            f"{base_select} ORDER BY ko.kp_id DESC LIMIT ?",
            (pattern, *access_params, *product_params, fetch_limit),
        )
        rows = [dict(row) for row in cur.fetchall()]
        _attach_product_types(cur, rows)

        if len(rows) > limit:
            cur.execute(
                f"""
                SELECT COUNT(*) AS cnt
                FROM KP_offers ko
                LEFT JOIN kp_meta m ON ko.kp_id = m.kp_id
                WHERE casefold(ko.customer_name) LIKE casefold(?) ESCAPE '\\'{access_sql}{product_sql}
                """,
                (pattern, *access_params, *product_params),
            )
            total = int(cur.fetchone()["cnt"])
            return rows[:limit], total

        return rows, len(rows)
    finally:
        conn.close()


def get_kp_completion_percentage(kp_id: int, db_path: str = DEFAULT_DB) -> Dict:
    """Процент выполнения КП по плитам."""
    conn = _connect(db_path)

    try:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        cur.execute(
            """
            SELECT COALESCE(SUM(qty), 0) as qty
            FROM kp_plates
            WHERE kp_id = ?
            """,
            (kp_id,),
        )
        in_production = cur.fetchone()["qty"]

        cur.execute(
            """
            SELECT COALESCE(SUM(qty), 0) as qty
            FROM completed_plates
            WHERE kp_id = ?
            """,
            (kp_id,),
        )
        completed = cur.fetchone()["qty"]

        total = in_production + completed

        if total > 0:
            percentage = (completed / total) * 100
        else:
            percentage = 0.0

        return {
            "total_plates": total,
            "completed_plates": completed,
            "in_production": in_production,
            "percentage": round(percentage, 1),
        }

    finally:
        conn.close()


def get_kp_plates_in_plan_percentage(kp_id: int, db_path: str = DEFAULT_DB) -> Dict:
    """Процент плит КП, уже включённых в план производства."""
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT COALESCE(SUM(qty), 0) as total
            FROM kp_plates
            WHERE kp_id = ? AND status IN ('в производстве', 'в плане')
            """,
            (kp_id,),
        )
        total = cur.fetchone()[0]
        cur.execute(
            """
            SELECT COALESCE(SUM(qty), 0) as in_plan
            FROM kp_plates
            WHERE kp_id = ? AND status = 'в плане'
            """,
            (kp_id,),
        )
        in_plan = cur.fetchone()[0]
        percentage = (in_plan / total * 100) if total > 0 else 0.0
        return {
            "total_plates": total,
            "in_plan": in_plan,
            "percentage": round(percentage, 1),
        }
    finally:
        conn.close()


def get_kp_total_length(kp_id: int, db_path: str = DEFAULT_DB) -> float:
    """Суммарная длина плит КП в метрах."""
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT COALESCE(SUM(COALESCE(length_m, 0) * qty), 0.0)
            FROM kp_plates
            WHERE kp_id = ? AND status IN ('в производстве', 'в плане')
            """,
            (kp_id,),
        )
        return float(cur.fetchone()[0])
    finally:
        conn.close()
