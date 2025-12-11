#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Восстановление базы pb.db с базовыми ценами плит из файла
«банк знаний/Новые цены для прайса с 19.08.24.xlsx».

- Создаёт файл pb.db в корне проекта (если его ещё нет).
- Создаёт таблицу prices (length_dm, load_code, price).
- Записывает/обновляет все цены из Excel.
- Не трогает bot/pb.db и таблицу reinforcement_loads.
"""

from core import config_and_data as cfg
from viz_modules.price_utils import sync_price_xlsx_to_db


def main() -> None:
    """Заливает прайс из XLSX в SQLite-базу pb.db."""
    xlsx_path = cfg.PRICE_XLSX_PATH   # банк знаний/Новые цены для прайса с 19.08.24.xlsx
    db_path = cfg.PRICE_DB_PATH       # корневой pb.db

    print(f"Прайс XLSX: {xlsx_path}")
    print(f"База SQLite: {db_path}")

    inserted = sync_price_xlsx_to_db(
        xlsx_path=xlsx_path,
        db_path=db_path,
    )

    print(f"Готово. В таблицу prices записано строк: {inserted}")


if __name__ == "__main__":
    main()


