from __future__ import annotations

import argparse
import sqlite3
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.core.settings import get_settings
from app.planning import plan_manager
from core import kp_db


def _connect(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def _list_plan_ids(db_path: str, requested_plan_id: str | None) -> list[str]:
    if requested_plan_id:
        return [requested_plan_id]

    with _connect(db_path) as conn:
        rows = conn.execute(
            """
            SELECT DISTINCT plan_id
            FROM kp_plates
            WHERE status = 'в плане'
              AND plan_id IS NOT NULL
              AND qty > 0
            ORDER BY plan_id
            """
        ).fetchall()
    return [str(row["plan_id"]) for row in rows]


def _plan_summary(plan_id: str) -> tuple[int, int]:
    plan = plan_manager.load_plan(plan_id)
    if not plan:
        return 0, 0
    days = plan.get("days") or {}
    completed = [
        date_key
        for date_key, day in days.items()
        if isinstance(day, dict) and day.get("completed")
    ]
    return len(completed), len(days)


def _stuck_rows(db_path: str, plan_id: str) -> tuple[int, int, list[sqlite3.Row]]:
    with _connect(db_path) as conn:
        total_qty = conn.execute(
            """
            SELECT COALESCE(SUM(qty), 0)
            FROM kp_plates
            WHERE status = 'в плане'
              AND plan_id = ?
              AND qty > 0
            """,
            (plan_id,),
        ).fetchone()[0]
        rows = conn.execute(
            """
            SELECT kp_id, plate_name, qty
            FROM kp_plates
            WHERE status = 'в плане'
              AND plan_id = ?
              AND qty > 0
            ORDER BY kp_id, plate_name
            LIMIT 10
            """,
            (plan_id,),
        ).fetchall()
    return int(total_qty or 0), len(rows), rows


def recover_stuck_plan_plates(
    *,
    db_path: str,
    plan_id: str | None,
    apply: bool,
) -> int:
    """Shows or returns stuck plan plates using the existing DB transition helper."""
    plan_ids = _list_plan_ids(db_path, plan_id)
    if not plan_ids:
        print("Зависших плит со status='в плане' и plan_id не найдено.")
        return 0

    total_returned = 0
    for current_plan_id in plan_ids:
        stuck_qty, _sample_count, sample = _stuck_rows(db_path, current_plan_id)
        completed_days, total_days = _plan_summary(current_plan_id)
        print(
            f"План {current_plan_id}: зависло {stuck_qty} плит, "
            f"выполнено дней {completed_days}/{total_days}."
        )
        for row in sample:
            print(f"  КП {row['kp_id']}: {row['plate_name']} x{row['qty']}")

        if not apply:
            continue

        returned = kp_db.return_plan_plates_to_production(current_plan_id, db_path)
        total_returned += returned
        print(f"  Возвращено в производство: {returned}")

    if not apply:
        print("Dry-run: изменения не применялись. Для возврата добавьте --apply.")
    else:
        print(f"Итого возвращено в производство: {total_returned}")
    return total_returned


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Диагностика и безопасный возврат зависших плит из статуса "
            "'в плане' в 'в производстве'."
        )
    )
    parser.add_argument("--db-path", default=str(get_settings().plita_db_path))
    parser.add_argument("--plan-id", default=None)
    parser.add_argument(
        "--apply",
        action="store_true",
        help=(
            "Применить возврат только для одного плана (--plan-id обязателен). "
            "Без флага — dry-run по всем планам с зависшими плитами."
        ),
    )
    args = parser.parse_args()

    if args.apply and not args.plan_id:
        parser.error(
            "--apply требует явный --plan-id (один план). "
            "Массовый возврат для всех планов отключён; без --plan-id используйте dry-run."
        )

    recover_stuck_plan_plates(
        db_path=str(args.db_path),
        plan_id=args.plan_id,
        apply=bool(args.apply),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
