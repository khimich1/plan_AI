#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Импорт прайса лестничных маршей (ЛМ) из Excel в pb.db (таблица march_prices)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.march_price_db import import_march_prices_from_xlsx, parse_march_price_rows_from_xlsx
from core.project_paths import PRICE_DB_PATH


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Загрузить прайс лестничных маршей из Excel в pb.db (INSERT OR REPLACE)."
    )
    parser.add_argument("xlsx_path", help="Путь к Excel-файлу прайса")
    parser.add_argument(
        "--sheet",
        default="Прайс",
        help="Имя листа (по умолчанию: Прайс)",
    )
    parser.add_argument(
        "--db",
        default=str(PRICE_DB_PATH),
        help=f"Путь к pb.db (по умолчанию: {PRICE_DB_PATH})",
    )
    parser.add_argument(
        "--date",
        default=None,
        help="Дата прайса в формате YYYY-MM-DD (по умолчанию — из имени файла)",
    )
    args = parser.parse_args()

    xlsx_path = Path(args.xlsx_path)
    if not xlsx_path.is_file():
        print(f"❌ Файл не найден: {xlsx_path}")
        return 1

    preview = parse_march_price_rows_from_xlsx(str(xlsx_path), preferred_sheet=args.sheet)
    if not preview:
        print("❌ Не удалось прочитать прайс — проверьте лист «Прайс»")
        return 1

    inserted = import_march_prices_from_xlsx(
        str(xlsx_path),
        args.db,
        preferred_sheet=args.sheet,
        price_list_date=args.date,
    )
    unique_marks = len({row[0] for row in preview})
    unique_grades = len({row[1] for row in preview})
    print(f"Excel: {xlsx_path}")
    print(f"БД:    {args.db}")
    print(f"Готово: {inserted} строк ({unique_marks} марок × {unique_grades} классов бетона)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
