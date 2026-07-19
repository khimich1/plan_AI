#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Разовое выравнивание столбца factory_plate_costs.concrete_grade по pb_reinforcement_series.

Не используется бизнес-логикой приложения — только для согласованности при просмотре БД.

Usage:
    python scripts/backfill_factory_plate_costs_concrete_grade.py [--db PATH_TO_pb.db]
"""
from __future__ import annotations

import argparse
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--db",
        default=str(ROOT / "pb.db"),
        help="Путь к pb.db (таблицы factory_plate_costs и pb_reinforcement_series).",
    )
    args = ap.parse_args()
    db_path = Path(args.db).resolve()

    conn = sqlite3.connect(str(db_path))
    cur = conn.cursor()
    cur.execute(
        """SELECT name FROM sqlite_master WHERE type='table'
           AND name IN ('factory_plate_costs', 'pb_reinforcement_series')"""
    )
    found = {r[0] for r in cur.fetchall()}
    if "factory_plate_costs" not in found or "pb_reinforcement_series" not in found:
        print("Не найдены обе таблицы — выход без изменений.")
        conn.close()
        return 1

    cur.execute("SELECT COUNT(*) FROM factory_plate_costs")
    n_before = int(cur.fetchone()[0])

    sql = """
    UPDATE factory_plate_costs
    SET concrete_grade = (
      SELECT concrete_grade FROM pb_reinforcement_series AS s
      WHERE s.length_dm = factory_plate_costs.length_dm
        AND s.load_code = CAST(ROUND(factory_plate_costs.load_code + 0.499999999) AS INTEGER)
      LIMIT 1
    )
    WHERE EXISTS (
      SELECT 1 FROM pb_reinforcement_series AS s2
      WHERE s2.length_dm = factory_plate_costs.length_dm
        AND s2.load_code = CAST(ROUND(factory_plate_costs.load_code + 0.499999999) AS INTEGER)
    )
    """
    cur.execute(sql)
    conn.commit()
    touched = conn.total_changes
    conn.close()
    print(f"factory_plate_costs: строк до={n_before}, строк обновлено (best-effort)={touched}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
