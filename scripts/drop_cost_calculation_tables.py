#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Одноразовый скрипт: удаление таблиц, связанных с удалённым модулем cost_calculation.

Запуск из корня проекта:
    python scripts/drop_cost_calculation_tables.py
    python scripts/drop_cost_calculation_tables.py --db путь/к/pb.db

Удаляет из pb.db таблицы:
  cost_constants, concrete_norms, reinforcement_norms, izoform_norms,
  plate_kef_values, plate_volumes, reinforcement_costs, concrete_costs,
  izoform_costs, excel_total_costs
"""

from __future__ import annotations

import argparse
import sqlite3
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_DB = PROJECT_ROOT / "pb.db"

TABLES = [
    "cost_constants",
    "concrete_norms",
    "reinforcement_norms",
    "izoform_norms",
    "plate_kef_values",
    "plate_volumes",
    "reinforcement_costs",
    "concrete_costs",
    "izoform_costs",
    "excel_total_costs",
]


def main() -> None:
    parser = argparse.ArgumentParser(description="Удаление таблиц cost_calculation из pb.db")
    parser.add_argument(
        "--db",
        type=Path,
        default=DEFAULT_DB,
        help=f"Путь к pb.db (по умолчанию: {DEFAULT_DB})",
    )
    args = parser.parse_args()
    db_path = args.db

    if not db_path.exists():
        print(f"Файл БД не найден: {db_path}", file=sys.stderr)
        sys.exit(1)

    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.cursor()
        for name in TABLES:
            cur.execute("DROP TABLE IF EXISTS " + name)
            print(f"  DROP TABLE IF EXISTS {name}")
        conn.commit()
        print("Готово.")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
