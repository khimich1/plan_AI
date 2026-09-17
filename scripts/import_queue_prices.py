#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Импорт заполненных цен из очереди GUID/цен в прайс-таблицы pb.db (GPS-008).

Читает лист «Задачи», блок «💰 ввести цену», колонку «Цена за шт (ввод)».
Пустые ячейки пропускаются (человек ещё не ввёл цену). Не запускайте
против продакшен pb.db с выдуманными ценами.

Запуск из корня репозитория:

    .venv/bin/python scripts/import_queue_prices.py --dry-run
    .venv/bin/python scripts/import_queue_prices.py --db /path/to/copy.db
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.price_import_queue import (
    QueueFormatError,
    format_report,
    import_queue_prices,
)
from core.project_paths import PRICE_DB_PATH

DEFAULT_QUEUE = (
    PROJECT_ROOT
    / "банк знаний"
    / "Новая папка"
    / "GUID"
    / "Очередь синхронизации GUID и цен.xlsx"
)


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Залить заполненные «Цена за шт (ввод)» из очереди GPS-001 "
            "в прайс-таблицы (сваи — на все классы бетона марки)."
        )
    )
    parser.add_argument(
        "xlsx_path",
        nargs="?",
        default=str(DEFAULT_QUEUE),
        help=f"Excel-очередь (по умолчанию: {DEFAULT_QUEUE})",
    )
    parser.add_argument(
        "--db",
        default=str(PRICE_DB_PATH),
        help=f"Путь к pb.db (по умолчанию: {PRICE_DB_PATH})",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Проверить строки и показать отчёт без записи в БД",
    )
    return parser.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    xlsx_path = Path(args.xlsx_path)
    db_path = Path(args.db)
    if not xlsx_path.is_file():
        print(f"❌ Файл очереди не найден: {xlsx_path}")
        return 1
    if not db_path.is_file() and not args.dry_run:
        # Schema is created on write; missing file is OK for a fresh temp DB,
        # but a typo against production should be visible.
        print(f"⚠ БД ещё нет, будет создана: {db_path}")
    try:
        report = import_queue_prices(
            xlsx_path,
            db_path,
            dry_run=args.dry_run,
        )
    except QueueFormatError as exc:
        print(f"❌ {exc}")
        return 1
    print(f"Excel: {xlsx_path}")
    print(f"БД:    {db_path}")
    print(format_report(report, dry_run=args.dry_run))
    return 1 if report.errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
