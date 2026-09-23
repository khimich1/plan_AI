#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Импорт GUID составных свай из книги сопоставления 1С."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.composite_pile_guid_import import (
    SHEET_NAME,
    import_composite_pile_guids_from_xlsx,
)
from core.project_paths import PRICE_DB_PATH


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Импорт GUID составных свай (лист «Сваи составные — только в 1С»)."
    )
    parser.add_argument("xlsx_path", help="Книга сопоставления GUID 1С")
    parser.add_argument("--sheet", default=SHEET_NAME, help=f"Лист (по умолчанию: {SHEET_NAME})")
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

    try:
        report = import_composite_pile_guids_from_xlsx(
            str(xlsx_path), args.db, sheet_name=args.sheet
        )
    except Exception as exc:  # noqa: BLE001 — CLI surface
        print(f"❌ {exc}")
        return 1

    print(f"Excel: {xlsx_path}")
    print(f"БД:    {args.db}")
    print(
        f"Готово: auto={report.written_auto}, "
        f"ambiguous={report.written_ambiguous}, skipped={report.skipped}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
