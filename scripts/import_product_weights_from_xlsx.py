#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Импорт весов ФБС/ЛС/ЛМ из выгрузки 1С в plita.db (product_weight_catalog)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.kp_db_common import DEFAULT_DB
from core.product_weight_catalog import (
    PRODUCT_WEIGHT_TYPES,
    parse_product_weights_from_xlsx,
    upsert_product_weights,
)


def _print_report(result, *, xlsx_path: Path, db_path: str, dry_run: bool) -> None:
    mode = "dry-run" if dry_run else "import"
    print(f"Excel: {xlsx_path}")
    print(f"БД:    {db_path}")
    print(f"Режим: {mode}")
    print(f"Импортировано: {len(result.records)}")
    print(f"Quarantine:    {len(result.quarantine)}")
    for item in result.quarantine:
        print(f"  - {item.display_name}: {item.reason}")
    print(f"Skipped:       {len(result.skipped)}")
    for item in result.skipped:
        print(f"  - {item.display_name}: {item.reason}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Загрузить веса изделий из Excel в product_weight_catalog."
    )
    parser.add_argument("xlsx_path", help="Путь к Excel-файлу выгрузки 1С")
    parser.add_argument(
        "--type",
        required=True,
        choices=sorted(PRODUCT_WEIGHT_TYPES),
        help="Тип продукции: fbs, steps или marches",
    )
    parser.add_argument(
        "--db",
        default=DEFAULT_DB,
        help=f"Путь к plita.db (по умолчанию: {DEFAULT_DB})",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Только отчёт, без записи в БД",
    )
    args = parser.parse_args()

    xlsx_path = Path(args.xlsx_path)
    if not xlsx_path.is_file():
        print(f"❌ Файл не найден: {xlsx_path}")
        return 1

    result = parse_product_weights_from_xlsx(str(xlsx_path), args.type)
    _print_report(result, xlsx_path=xlsx_path, db_path=args.db, dry_run=args.dry_run)
    if args.dry_run:
        return 0

    inserted, updated = upsert_product_weights(
        args.db,
        result.records,
        source_file=xlsx_path.name,
    )
    print(f"Готово: добавлено {inserted}, обновлено {updated}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
