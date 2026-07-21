"""Backfill ``kp_plates.day_number`` для существующих планов.

После миграции ``ALTER TABLE kp_plates ADD COLUMN day_number INTEGER`` все
строки имеют ``day_number IS NULL``. ``build_day_view_detail`` для таких
планов будет проваливаться в legacy fuzzy-lookup. Чтобы переключить
существующие активные планы на DB-путь, нужно:

1. Прочитать каждый файл ``data/plans/*.json``.
2. Для каждой даты в ``plan.days[date].tracks`` определить ``day_number``.
3. Для каждой плиты ``kp_id + plate_name`` пометить соответствующую строку
   ``kp_plates`` с этим ``plan_id`` и ``status='в плане'``, проставив
   ``day_number``.

Скрипт идемпотентен: если у строки ``day_number`` уже задан, она не
перезаписывается (можно безопасно запускать повторно).

Использование::

    python scripts/backfill_day_number.py [--db-path DB] [--plans-dir DIR] [--dry-run]
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import sqlite3
import sys
from collections import defaultdict
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from core import plate_name as _plate_name  # noqa: E402

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s %(message)s",
)
logger = logging.getLogger("backfill_day_number")


def _iter_plate_keys(track: dict) -> list[tuple[int, str]]:
    """Возвращает список ``(kp_id, canonical_name)`` для каждого item трека."""
    out: list[tuple[int, str]] = []
    for item in track.get("items") or []:
        if not item:
            continue
        kp_id = item.get("kp_id")
        name = item.get("plate_name") or item.get("label") or ""
        canon = _plate_name.canonical(name)
        if kp_id is None or not canon:
            continue
        out.append((int(kp_id), canon))
    return out


def backfill_plan(
    conn: sqlite3.Connection,
    plan_path: Path,
    dry_run: bool = False,
) -> tuple[int, int]:
    """Заливает day_number в kp_plates для одного plan.json.

    Returns:
        (rows_affected, rows_skipped) — сколько строк kp_plates обновлено
        и сколько пропущено (либо identity не совпадает, либо day_number
        уже задан, либо строки не нашлось).
    """
    try:
        plan = json.loads(plan_path.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.warning("Не удалось прочитать %s: %s", plan_path, exc)
        return 0, 0

    plan_id = plan.get("id")
    if not plan_id:
        logger.warning("Пропускаю %s: нет plan.id", plan_path)
        return 0, 0

    days = plan.get("days") or {}
    if not days:
        return 0, 0

    rows_affected = 0
    rows_skipped = 0
    cur = conn.cursor()

    # demand[(kp_id, canonical_name, day_number)] = qty
    demand: dict[tuple[int, str, int], int] = defaultdict(int)
    for date_key, day_data in days.items():
        day_number = int((day_data or {}).get("day_number") or 0)
        if day_number <= 0:
            continue
        for track in (day_data or {}).get("tracks") or []:
            for kp_id, canon in _iter_plate_keys(track):
                demand[(kp_id, canon, day_number)] += 1

    if not demand:
        return 0, 0

    for (kp_id, canon, day_number), qty in demand.items():
        # Ищем строку kp_plates с этим планом, identity и не-проставленным day_number.
        cur.execute(
            """
            SELECT id, plate_name, qty, day_number
            FROM kp_plates
            WHERE kp_id = ? AND plan_id = ? AND status = 'в плане'
            ORDER BY id
            """,
            (kp_id, plan_id),
        )
        rows = cur.fetchall()
        # Фильтруем по canonical(name)
        matching = [
            (rid, qty_in_db, day_in_db)
            for (rid, name, qty_in_db, day_in_db) in rows
            if _plate_name.canonical(name) == canon
        ]
        if not matching:
            logger.debug(
                "Не найдено строк kp_plates для plan=%s kp=%s name=%s day=%s",
                plan_id, kp_id, canon, day_number,
            )
            rows_skipped += qty
            continue

        # Назначаем day_number первой строке без day_number.
        # Если есть несколько строк, и day_number у одной уже совпадает — пропускаем.
        chosen = next(
            (m for m in matching if m[2] is None),
            None,
        )
        if chosen is None:
            # Все строки уже имеют day_number — ничего не делаем.
            rows_skipped += qty
            continue

        rid = chosen[0]
        if dry_run:
            logger.info(
                "[DRY] UPDATE kp_plates SET day_number=%s WHERE id=%s (plan=%s, %s)",
                day_number, rid, plan_id, canon,
            )
            rows_affected += 1
            continue
        cur.execute(
            "UPDATE kp_plates SET day_number = ? WHERE id = ?",
            (day_number, rid),
        )
        rows_affected += cur.rowcount

    if not dry_run:
        conn.commit()
    return rows_affected, rows_skipped


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--db-path",
        default=str(PROJECT_ROOT / "plita.db"),
        help="Путь к plita.db (по умолчанию ./plita.db)",
    )
    parser.add_argument(
        "--plans-dir",
        default=str(PROJECT_ROOT / "data" / "plans"),
        help="Директория с JSON-планами (по умолчанию data/plans)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Не вносить изменения, только распечатать что будет сделано.",
    )
    args = parser.parse_args()

    db_path = args.db_path
    plans_dir = Path(args.plans_dir)
    if not Path(db_path).exists():
        logger.error("База %s не найдена", db_path)
        return 1
    if not plans_dir.exists():
        logger.error("Директория планов %s не найдена", plans_dir)
        return 1

    plan_files = sorted(plans_dir.glob("*.json"))
    if not plan_files:
        logger.warning("Нет JSON-планов в %s", plans_dir)
        return 0

    total_affected = 0
    total_skipped = 0
    with sqlite3.connect(db_path) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        for plan_path in plan_files:
            affected, skipped = backfill_plan(conn, plan_path, dry_run=args.dry_run)
            total_affected += affected
            total_skipped += skipped
            logger.info(
                "Plan %s: обновлено %s, пропущено %s",
                plan_path.name, affected, skipped,
            )

    logger.info(
        "ИТОГО: обновлено %s строк, пропущено %s. Dry-run=%s.",
        total_affected, total_skipped, args.dry_run,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
