#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Импорт армирования/нагрузок из XLSX в SQLite.

Ожидаем таблицу как в файле `банк знаний/армирование_нагрузки_таблица.xlsx`:
 - колонка "Марка плиты" (например, "ПБ 71-12")
 - пары колонок по каждой нагрузке: "X-нагрузка, по серии" и "X-нагрузка, эрм"
   (X = 6/8/10/12.5/16)

В базе создаётся таблица reinforcement_loads:
 (length_dm INTEGER, load_code INTEGER, source TEXT, value REAL, PRIMARY KEY(...))
где source ∈ {"series", "erm"}.
"""

import os
import re
import sqlite3
from pathlib import Path
from typing import Iterable, Optional, Tuple

try:
    import pandas as pd
except Exception:
    pd = None

from core.db_config import PB_DB_PATH

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_DB = PB_DB_PATH
DEFAULT_XLSX = BASE_DIR / "банк знаний" / "армирование_нагрузки_таблица.xlsx"


def init_schema(db_path: Path | str = DEFAULT_DB) -> None:
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS reinforcement_loads (
                length_dm INTEGER,
                load_code INTEGER,
                source TEXT,    -- 'series' | 'erm'
                value REAL,
                PRIMARY KEY(length_dm, load_code, source)
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def _extract_length_dm(mark: str) -> int | None:
    """
    Берём первую группу цифр (предпочтение: до дефиса; иначе первый int в строке).
    Примеры:
      'ПБ 71-12' -> 71
      'ПБ 17'    -> 17
    """
    s = str(mark)
    m = re.search(r"(\d+)\s*[-–]", s)
    if not m:
        m = re.search(r"(\d+)", s)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def _iter_rows_from_df(df) -> Iterable[Tuple[int, int, str, float]]:
    """
    Генерирует строки (length_dm, load_code, source, value) из датафрейма.
    """
    # Сопоставляем названия колонок -> (load_code, source)
    col_map = {}
    for col in df.columns:
        name = str(col).lower().strip()
        # пример: "8-нагрузка, по серии", "12,5-нагрузка, арм", "6-нагрузка, униф. арм"
        if "нагрузка" not in name:
            continue
        m = re.search(r"(\d+(?:[.,]\d+)?)", name)
        if not m:
            continue
        load_val = m.group(1).replace(",", ".")
        try:
            load_code = int(float(load_val) + 0.5)  # 12.5 -> 13
        except Exception:
            continue
        # Источники в файле: "по серии" и "арм"/"униф. арм"
        if "сер" in name:
            source = "series"
        elif "арм" in name:
            source = "arm"  # объединяем обычную и "униф. арм"
        else:
            source = None
        if source is None:
            continue
        col_map[col] = (load_code, source)

    mark_col = next((c for c in df.columns if "марка" in str(c).lower()), None)
    if mark_col is None or not col_map:
        return []

    rows = []
    for _, row in df.iterrows():
        length_dm = _extract_length_dm(row.get(mark_col, ""))
        if length_dm is None:
            continue
        for col, (load_code, source) in col_map.items():
            val = row.get(col)
            if pd.notna(val):
                try:
                    value = float(str(val).replace(" ", "").replace(",", "."))
                    rows.append((length_dm, load_code, source, value))
                except Exception:
                    continue
    return rows


def import_reinforcement_from_xlsx(
    xlsx_path: Path | str = DEFAULT_XLSX, db_path: Path | str = DEFAULT_DB
) -> int:
    """
    Считывает XLSX и пишет в таблицу reinforcement_loads.
    Возвращает количество вставленных/обновлённых записей.
    """
    if pd is None:
        print("pandas не установлен — импорт пропущен")
        return 0
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        print("Файл не найден:", xlsx_path)
        return 0

    df = pd.read_excel(xlsx_path, sheet_name=0)
    rows = _iter_rows_from_df(df)
    if not rows:
        print("Не удалось извлечь данные из XLSX (проверьте колонки).")
        return 0

    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.executemany(
            """
            INSERT OR REPLACE INTO reinforcement_loads (length_dm, load_code, source, value)
            VALUES (?, ?, ?, ?)
            """,
            rows,
        )
        conn.commit()
        return len(rows)
    finally:
        conn.close()


def get_reinforcement(
    length_m: float,
    load_code: int | float,
    source: str = "erm",
    db_path: Path | str = DEFAULT_DB,
    allow_fallback: bool = True,
) -> float | None:
    """
    Возвращает значение армирования по длине и нагрузке.
    Сначала ищет в таблице pb_reinforcement_series, затем в reinforcement_loads.
    - source: 'erm' (предпочтительно) или 'series' (используется только для reinforcement_loads)
    - allow_fallback: искать ближайшую длину ±1 дм, либо переключиться на другую таблицу
    """
    length_dm = int(round(length_m * 10))
    load_code_int = int(float(load_code) + 0.5)

    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # === ШАГ 1: Ищем в новой таблице pb_reinforcement_series ===
        cur.execute("""
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name='pb_reinforcement_series'
        """)
        if cur.fetchone():
            # Таблица существует - ищем в ней
            cur.execute("""
                SELECT reinforcement_value FROM pb_reinforcement_series
                WHERE length_dm=? AND load_code=?
            """, (length_dm, load_code_int))
            row = cur.fetchone()
            if row:
                return float(row[0])
            
            # Fallback: ближайшая длина ±1 дм в pb_reinforcement_series
            if allow_fallback:
                cur.execute("""
                    SELECT reinforcement_value FROM pb_reinforcement_series
                    WHERE ABS(length_dm - ?) <= 1 AND load_code=?
                    ORDER BY ABS(length_dm-?) LIMIT 1
                """, (length_dm, load_code_int, length_dm))
                row = cur.fetchone()
                if row:
                    return float(row[0])
        
        # === ШАГ 2: Fallback - ищем в старой таблице reinforcement_loads ===
        cur.execute("""
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name='reinforcement_loads'
        """)
        if cur.fetchone():
            # 1) точное совпадение по source
            cur.execute(
                """
                SELECT value FROM reinforcement_loads
                WHERE length_dm=? AND load_code=? AND source=?
                """,
                (length_dm, load_code_int, source),
            )
            row = cur.fetchone()
            if row:
                return float(row[0])

            # 2) fallback: другая source
            if allow_fallback:
                alt_source = "series" if source == "erm" else "erm"
                cur.execute(
                    """
                    SELECT value FROM reinforcement_loads
                    WHERE length_dm=? AND load_code=? AND source=?
                    """,
                    (length_dm, load_code_int, alt_source),
                )
                row = cur.fetchone()
                if row:
                    return float(row[0])

            # 3) fallback: ближайшая длина ±1 дм
            if allow_fallback:
                cur.execute(
                    """
                    SELECT value FROM reinforcement_loads
                    WHERE ABS(length_dm - ?) <= 1 AND load_code=? AND source=?
                    ORDER BY ABS(length_dm-?) LIMIT 1
                    """,
                    (length_dm, load_code_int, source, length_dm),
                )
                row = cur.fetchone()
                if row:
                    return float(row[0])
    finally:
        conn.close()
    return None


def get_concrete_grade_from_series(
    length_m: float,
    load_code: int | float,
    *,
    db_path: Path | str = DEFAULT_DB,
    allow_fallback: bool = True,
) -> Optional[str]:
    """Марка бетона из pb_reinforcement_series (те же ключи, что для get_reinforcement)."""
    length_dm = int(round(float(length_m) * 10))
    load_code_int = int(float(load_code) + 0.5)

    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.cursor()
        cur.execute(
            """SELECT name FROM sqlite_master WHERE type='table' AND name='pb_reinforcement_series'"""
        )
        if not cur.fetchone():
            return None
        cur.execute(
            """SELECT concrete_grade FROM pb_reinforcement_series WHERE length_dm=? AND load_code=?""",
            (length_dm, load_code_int),
        )
        row = cur.fetchone()
        if row and row[0]:
            return str(row[0]).strip()
        if allow_fallback:
            cur.execute(
                """
                SELECT concrete_grade FROM pb_reinforcement_series
                WHERE ABS(length_dm - ?) <= 1 AND load_code=?
                ORDER BY ABS(length_dm - ?) LIMIT 1
                """,
                (length_dm, load_code_int, length_dm),
            )
            row = cur.fetchone()
            if row and row[0]:
                return str(row[0]).strip()
    finally:
        conn.close()
    return None


if __name__ == "__main__":
    inserted = import_reinforcement_from_xlsx()
    print(f"Импорт завершён, записей: {inserted}")

