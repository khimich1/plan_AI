#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Разовая заливка базовых цен плит в таблицу prices из Excel-файла
из папки «банк знаний».

- Ищет файл по части имени (слова "нов" и "цен"), чтобы не споткнуться
  о разные варианты буквы Й.
- Работает только с КОРНЕВОЙ pb.db (там, где ищет код).
- Не трогает bot/pb.db и reinforcement_loads.
"""

import os
import re
import sqlite3

import pandas as pd

from core import config_and_data as cfg

DB_PATH = cfg.PRICE_DB_PATH
BANK_DIR = os.path.join(cfg.BASE_DIR, "банк знаний")


def _find_price_xlsx() -> str | None:
    """Ищет Excel с прайсом в папке 'банк знаний' по части имени."""
    if not os.path.isdir(BANK_DIR):
        return None

    for name in os.listdir(BANK_DIR):
        # пропускаем временные файлы Excel вида "~$..."
        if name.startswith("~$"):
            continue
        low = name.lower()
        if not low.endswith(".xlsx"):
            continue
        # например: "Новые цены для прайса с 19.08.24.xlsx"
        if "нов" in low and "цен" in low:
            return os.path.join(BANK_DIR, name)
    return None


def main() -> None:
    xlsx_path = _find_price_xlsx()
    if not xlsx_path or not os.path.exists(xlsx_path):
        print("❌ Не нашёл Excel с прайсом в папке 'банк знаний'")
        print("Убедись, что там есть .xlsx, в названии которого есть 'нов' и 'цен'")
        return

    print(f"База: {DB_PATH}")
    print(f"Excel: {xlsx_path}")

    # Читаем лист. header=1 — используем вторую строку как заголовки,
    # чтобы получить колонки вида «Наименование», «6 нагрузка» и т.д.
    df = pd.read_excel(xlsx_path, sheet_name="24.06.2024", header=1)
    print("Колонки листа:", list(df.columns))

    # Первая колонка содержит наименование ("ПБ 17-12", "ПБ 18-12", ...),
    # её имя может быть как "Наименование", так и "Unnamed: 0" — берём как есть.
    if df.shape[1] == 0:
        print("❌ В листе нет колонок, не удалось прочитать прайс")
        return
    name_col = df.columns[0]

    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS prices (
                length_dm INTEGER,
                load_code INTEGER,
                price REAL,
                PRIMARY KEY(length_dm, load_code)
            )
            """
        )

        rows = []
        for _, row in df.iterrows():
            name = str(row.get(name_col, "")).strip()
            if not name:
                continue

            # Из строки вида "ПБ 71-12" или "ПБ 71–12" достаём 71 (длина в дм)
            m = re.search(r"(\d+)\s*[-–]\s*(\d+)", name)
            if not m:
                continue
            length_dm = int(m.group(1))

            # Берём цены для нагрузок 6/8/10/12, если колонки есть и значение не пустое
            for load_code in (6, 8, 10, 12):
                col_name = f"{load_code} нагрузка"
                if col_name not in df.columns:
                    continue
                val = row.get(col_name)
                if pd.notna(val):
                    try:
                        price = float(str(val).replace(" ", "").replace(",", "."))
                    except ValueError:
                        continue
                    rows.append((length_dm, load_code, price))

        cur.executemany(
            "INSERT OR REPLACE INTO prices (length_dm, load_code, price) VALUES (?,?,?)",
            rows,
        )
        conn.commit()
        print(f"Готово. Вставлено строк: {len(rows)}")
    finally:
        conn.close()


if __name__ == "__main__":
    main()


