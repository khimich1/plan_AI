#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Импорт прайса плит ПБ из Excel в pb.db (таблица prices)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.price_db import import_from_xlsx, parse_plate_price_rows_from_xlsx
from core.project_paths import PRICE_DB_PATH


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Загрузить прайс плит ПБ из Excel в pb.db (INSERT OR REPLACE)."
    )
    parser.add_argument("xlsx_path", help="Путь к Excel-файлу прайса")
    parser.add_argument(
        "--sheet",
        default=None,
        help="Имя листа (по умолчанию — первый или единственный)",
    )
    parser.add_argument(
        "--db",
        default=str(PRICE_DB_PATH),
        help=f"Путь к pb.db (по умолчанию: {PRICE_DB_PATH})",
    )
    args = parser.parse_args()

    xlsx_path = Path(args.xlsx_path)
    if not xlsx_path.is_file():
        print(f"❌ Файл не найден: {xlsx_path}")
        return 1

    preview = parse_plate_price_rows_from_xlsx(str(xlsx_path), preferred_sheet=args.sheet)
    if not preview:
        print("❌ Не удалось прочитать прайс — проверьте формат листа")
        return 1

    inserted = import_from_xlsx(str(xlsx_path), args.db, preferred_sheet=args.sheet)
    unique_lengths = len({row[0] for row in preview})
    print(f"Excel: {xlsx_path}")
    print(f"БД:    {args.db}")
    print(f"Готово: {inserted} строк ({unique_lengths} длин × нагрузки 6/8/10/12)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
