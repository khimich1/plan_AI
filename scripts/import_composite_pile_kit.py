#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Импорт комплектации составных свай (лист «Комплектация») в pb.db."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.composite_pile_kit import import_composite_pile_kit_from_xlsx
from core.project_paths import PRICE_DB_PATH


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Загрузить комплектацию составных свай в composite_pile_assemblies."
    )
    parser.add_argument("xlsx_path", help="Путь к xlsx комплектации серии")
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

    n = import_composite_pile_kit_from_xlsx(str(xlsx_path), args.db)
    print(f"Excel: {xlsx_path}")
    print(f"БД:    {args.db}")
    print(f"Готово: {n} сборок")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
