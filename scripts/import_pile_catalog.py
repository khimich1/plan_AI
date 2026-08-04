#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Импорт каталога свай из прайса (лист «Вес и объем») в plita.db (SHIP-100)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.kp_db_common import DEFAULT_DB
from core.pile_catalog import parse_pile_catalog_from_xlsx, upsert_pile_catalog


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Загрузить лист «Вес и объем» прайса свай в pile_catalog (upsert по марке)."
    )
    parser.add_argument("--xlsx", required=True, help="Путь к Excel-файлу прайса свай")
    parser.add_argument(
        "--db",
        default=DEFAULT_DB,
        help=f"Путь к plita.db (по умолчанию: {DEFAULT_DB})",
    )
    parser.add_argument(
        "--sheet",
        default="Вес и объем",
        help="Имя листа (по умолчанию: Вес и объем)",
    )
    args = parser.parse_args()

    xlsx_path = Path(args.xlsx)
    if not xlsx_path.is_file():
        print(f"❌ Файл не найден: {xlsx_path}")
        return 1

    try:
        entries = parse_pile_catalog_from_xlsx(str(xlsx_path), sheet=args.sheet)
    except ValueError as exc:
        print(f"❌ {exc}")
        return 1
    if not entries:
        print("❌ Не удалось прочитать каталог — проверьте лист «Вес и объем»")
        return 1

    inserted, updated = upsert_pile_catalog(args.db, entries)
    print(f"Excel: {xlsx_path}")
    print(f"БД:    {args.db}")
    print(f"Готово: {len(entries)} марок (добавлено {inserted}, обновлено {updated})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
