#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Отладка чтения прайса из XLSX: печатает первые строки листа 24.06.2024.
Нужен, чтобы увидеть реальные значения в первой колонке.
"""

import os

import pandas as pd

from core import config as cfg


def main() -> None:
    bank_dir = os.path.join(cfg.BASE_DIR, "банк знаний")
    # Ищем любой файл с "нов" и "цен" в имени
    xlsx_path = None
    for name in os.listdir(bank_dir):
        if name.startswith("~$"):
            continue
        low = name.lower()
        if low.endswith(".xlsx") and "нов" in low and "цен" in low:
            xlsx_path = os.path.join(bank_dir, name)
            break

    if not xlsx_path or not os.path.exists(xlsx_path):
        print("Файл прайса не найден")
        return

    print("Файл:", xlsx_path)
    df = pd.read_excel(xlsx_path, sheet_name="24.06.2024", header=1)
    print("Колонки:", list(df.columns))

    name_col = df.columns[0]
    print("name_col =", name_col)
    print("Первые 20 значений в первой колонке:")
    for i in range(min(20, len(df))):
        val = df.iloc[i, 0]
        print(i, "->", repr(val))


if __name__ == "__main__":
    main()


