#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Импорт перевозчиков из Excel-реестра отгрузок в plita.db (SHIP-101).

Берётся только имя контрагента (колонка «Организация») листов «Перевозчики»
и «Транспортные Компании». Дубликаты (после нормализации) — в отчёт, не в БД.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.carrier_catalog import import_carriers_from_xlsx
from core.kp_db_common import DEFAULT_DB


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Импорт перевозчиков из реестра отгрузок в carriers (с авто-дедупом)."
    )
    parser.add_argument("--xlsx", required=True, help="Путь к Excel-файлу реестра отгрузок")
    parser.add_argument(
        "--db",
        default=DEFAULT_DB,
        help=f"Путь к plita.db (по умолчанию: {DEFAULT_DB})",
    )
    args = parser.parse_args()

    xlsx_path = Path(args.xlsx)
    if not xlsx_path.is_file():
        print(f"❌ Файл не найден: {xlsx_path}")
        return 1

    report = import_carriers_from_xlsx(str(xlsx_path), args.db)
    print(f"Excel: {xlsx_path}")
    print(f"БД:    {args.db}")
    for sheet, stats in report.per_sheet.items():
        print(f"Лист «{sheet}»: прочитано {stats['read']}, импортировано {stats['imported']}")
    print(
        f"Итого: добавлено {report.inserted}, "
        f"пропущено существующих {report.skipped_existing}, "
        f"дублей схлопнуто {len(report.duplicates)}"
    )
    if report.duplicates:
        print("Дубли (не импортированы):")
        for dup in report.duplicates:
            print(
                f"  - «{dup['name']}» ({dup['sheet']}, строка {dup['row']}) "
                f"— совпадает с «{dup['kept']}»"
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
