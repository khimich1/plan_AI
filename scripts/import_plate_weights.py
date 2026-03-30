#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Импорт весов плит в pb.db из файла «банк знаний/Вес плит.xls».
Создаёт таблицу plate_weights (name, mass_kg) и заполняет её.
Если файл недоступен или формат не читается — использует встроенные данные из паспорта.
Таблица используется в legacy-режиме расчета веса (WEIGHT_SOURCE="plate_weights").
"""
import sys
from pathlib import Path

# Корень проекта
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

# Встроенные данные (марка панели → масса, кг) из паспорта, если xls недоступен
DEFAULT_WEIGHTS = [
    ("ПБ 90-12-8 п", 3390), ("ПБ 89-12-8 п", 3352), ("ПБ 88-12-8 п", 3315),
    ("ПБ 87-12-8 п", 3278), ("ПБ 86-12-8 п", 3240), ("ПБ 85-12-8 п", 3202),
    ("ПБ 84-12-8 п", 3165), ("ПБ 83-12-8 п", 3128), ("ПБ 82-12-8 п", 3090),
    ("ПБ 81-12-8 п", 3052), ("ПБ 80-12-8 п", 3015), ("ПБ 79-12-8 п", 2978),
    ("ПБ 78-12-8 п", 2940), ("ПБ 77-12-8 п", 2898), ("ПБ 76-12-8 п", 2860),
    ("ПБ 75-12-8 п", 2822), ("ПБ 74-12-8 п", 2785), ("ПБ 73-12-8 п", 2748),
    ("ПБ 72-12-8 п", 2710), ("ПБ 71-12-8 п", 2672), ("ПБ 70-12-8 п", 2635),
    ("ПБ 69-12-8 п", 2598), ("ПБ 68-12-8 п", 2560), ("ПБ 67-12-8 п", 2522),
    ("ПБ 66-12-8 п", 2485), ("ПБ 65-12-8 п", 2448), ("ПБ 64-12-8 п", 2410),
    ("ПБ 63-12-8 п", 2372), ("ПБ 62-12-8 п", 2335), ("ПБ 61-12-8 п", 2298),
    ("ПБ 60-12-8 п", 2260), ("ПБ 59-12-8 п", 2221), ("ПБ 58-12-8 п", 2183),
    ("ПБ 57-12-8 п", 2145), ("ПБ 56-12-8 п", 2107), ("ПБ 55-12-8 п", 2068),
    ("ПБ 54-12-8 п", 2030), ("ПБ 53-12-8 п", 1993), ("ПБ 52-12-8 п", 1956),
    ("ПБ 51-12-8 п", 1918), ("ПБ 50-12-8 п", 1880), ("ПБ 49-12-8 п", 1843),
    ("ПБ 48-12-8 п", 1805), ("ПБ 47-12-8 п", 1768), ("ПБ 46-12-8 п", 1731),
    ("ПБ 45-12-8 п", 1693), ("ПБ 44-12-8 п", 1655), ("ПБ 43-12-8 п", 1618),
    ("ПБ 42-12-8 п", 1580), ("ПБ 41-12-8 п", 1538), ("ПБ 40-12-8 п", 1500),
    ("ПБ 39-12-8 п", 1462), ("ПБ 38-12-8 п", 1425), ("ПБ 37-12-8 п", 1387),
    ("ПБ 36-12-8 п", 1340), ("ПБ 35-12-8 п", 1312), ("ПБ 34-12-8 п", 1274),
    ("ПБ 33-12-8 п", 1237), ("ПБ 32-12-8 п", 1200), ("ПБ 31-12-8 п", 1162),
    ("ПБ 30-12-8 п", 1125), ("ПБ 29-12-8 п", 1087), ("ПБ 28-12-8 п", 1050),
    ("ПБ 27-12-8 п", 1003), ("ПБ 26-12-8 п", 974), ("ПБ 25-12-8 п", 937),
    ("ПБ 24-12-8 п", 900), ("ПБ 23-12-8 п", 864), ("ПБ 22-12-8 п", 828),
    ("ПБ 21-12-8 п", 790), ("ПБ 20-12-8 п", 753), ("ПБ 19-12-8 п", 715),
    ("ПБ 18-12-8 п", 678), ("ПБ 17-12-8 п", 640),
]


def load_from_xls(xls_path: Path) -> list:
    """Читает марки и массы из xls. Возвращает список кортежей (name, mass_kg)."""
    try:
        import pandas as pd
    except ImportError:
        return []
    if not xls_path.exists():
        return []
    try:
        # .xls требует xlrd
        df = pd.read_excel(xls_path, engine="xlrd", header=0)
    except Exception:
        try:
            df = pd.read_excel(xls_path, header=0)
        except Exception:
            return []
    # Ищем колонки по смыслу
    name_col = None
    mass_col = None
    for c in df.columns:
        cstr = str(c).strip().lower()
        if "марка" in cstr or "наименование" in cstr or "панел" in cstr:
            name_col = c
        if "масс" in cstr or "вес" in cstr or "кг" in cstr:
            mass_col = c
    if name_col is None or mass_col is None:
        return []
    rows = []
    for _, row in df.iterrows():
        name = row.get(name_col)
        mass = row.get(mass_col)
        if pd.isna(name) or pd.isna(mass):
            continue
        name = str(name).strip()
        try:
            mass_kg = float(mass)
        except (TypeError, ValueError):
            continue
        if name:
            rows.append((name, mass_kg))
    return rows


def main():
    from core.db_config import PB_DB_PATH
    from core.plate_weights_db import init_plate_weights_table, get_plate_weight_kg

    db_path = PB_DB_PATH
    # Путь к xls: сначала в корне проекта, потом на уровень выше
    for base in (PROJECT_ROOT, PROJECT_ROOT.parent):
        xls_path = base / "банк знаний" / "Вес плит.xls"
        if xls_path.exists():
            break
    else:
        xls_path = PROJECT_ROOT / "банк знаний" / "Вес плит.xls"

    data = load_from_xls(xls_path)
    if not data:
        print("[import_plate_weights] Файл не найден или не прочитан, используем встроенные данные.")
        data = DEFAULT_WEIGHTS
    else:
        print(f"[import_plate_weights] Загружено {len(data)} записей из {xls_path}")

    init_plate_weights_table(db_path)
    import sqlite3
    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.cursor()
        cur.execute("DELETE FROM plate_weights")
        cur.executemany(
            "INSERT INTO plate_weights (name, mass_kg) VALUES (?, ?)",
            data,
        )
        conn.commit()
        cur.execute("SELECT COUNT(*) FROM plate_weights")
        n = cur.fetchone()[0]
        print(f"[import_plate_weights] В таблицу plate_weights записано {n} строк.")
    finally:
        conn.close()

    # Проверка
    w = get_plate_weight_kg("ПБ 78-12-8 п")
    print(f"[import_plate_weights] Проверка: ПБ 78-12-8 п -> {w} кг" if w else "[import_plate_weights] Проверка: запись не найдена")


if __name__ == "__main__":
    main()
