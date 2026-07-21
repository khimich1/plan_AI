#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Таблица весов плит в pb.db (plate_weights).
Используется в КП для подстановки паспортной массы вместо расчётной.
Активируется в legacy-режиме: WEIGHT_SOURCE="plate_weights".
"""
import re
import sqlite3
from pathlib import Path
from typing import Optional

from .db_config import PB_DB_PATH


def _connect(db_path: Path = None):
    if db_path is None:
        db_path = PB_DB_PATH
    conn = sqlite3.connect(str(db_path))
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_plate_weights_table(db_path: Path = None) -> None:
    """Создаёт таблицу plate_weights в pb.db, если её ещё нет."""
    conn = _connect(db_path)
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS plate_weights (
                name TEXT PRIMARY KEY,
                mass_kg REAL NOT NULL
            )
        """)
        conn.commit()
    finally:
        conn.close()


def _normalize_name_for_lookup(name: str) -> str:
    """Приводит название плиты к виду для поиска в таблице весов."""
    if not name or not isinstance(name, str):
        return ""
    s = name.strip()
    # Убираем префикс "Плиты "
    if s.upper().startswith("ПЛИТЫ "):
        s = s[6:].strip()
    # В таблице весов формат "ПБ 90-12-8 п" (с пробелом перед "п"), в заказах часто "8п"
    s = s.replace(" п", "п").replace("  ", " ")
    return s.strip()


def get_plate_weight_kg(name: str, db_path: Path = None) -> Optional[float]:
    """
    Возвращает массу плиты в кг по названию из таблицы plate_weights.
    Пробует точное совпадение и нормализованное (пробел перед «п», префикс «Плиты »).
    """
    if not name:
        return None
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        # Точное совпадение
        cur.execute("SELECT mass_kg FROM plate_weights WHERE name = ?", (name.strip(),))
        row = cur.fetchone()
        if row is not None:
            return float(row[0])
        # Нормализованное: убираем/добавляем пробел перед "п"
        normalized = _normalize_name_for_lookup(name)
        if normalized:
            cur.execute("SELECT mass_kg FROM plate_weights WHERE name = ?", (normalized,))
            row = cur.fetchone()
            if row is not None:
                return float(row[0])
        # Пробуем вариант с пробелом перед "п" (если в name было "8п")
        if re.search(r"\dп$", normalized):
            alt = re.sub(r"(\d)п$", r"\1 п", normalized)
            cur.execute("SELECT mass_kg FROM plate_weights WHERE name = ?", (alt,))
            row = cur.fetchone()
            if row is not None:
                return float(row[0])
        return None
    finally:
        conn.close()


def get_plate_weight_kg_by_dimensions(
    length_m: float,
    width_m: float,
    db_path: Path = None,
) -> Optional[float]:
    """
    Возвращает массу плиты в кг по длине и ширине.

    База `plate_weights` хранит массу только для базовых плит:
    - ширина 12 дм (1.2м)
    - нагрузка 8

    Нагрузка для расчёта массы не учитывается.
    Для ширины меньше 1.2м масса считается пропорционально:
    base_mass * (width_m / 1.2)
    """
    try:
        length_dm = int(round(float(length_m) * 10))
        width_value = float(width_m)
    except (TypeError, ValueError):
        return None

    if length_dm <= 0 or width_value <= 0:
        return None

    base_mass: Optional[float] = None
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        # Каноническое имя для таблицы весов
        canonical_name = f"ПБ {length_dm}-12-8 п"
        cur.execute("SELECT mass_kg FROM plate_weights WHERE name = ?", (canonical_name,))
        row = cur.fetchone()
        if row is not None:
            base_mass = float(row[0])
        else:
            # На случай хранения без пробела перед "п"
            alt_name = f"ПБ {length_dm}-12-8п"
            cur.execute("SELECT mass_kg FROM plate_weights WHERE name = ?", (alt_name,))
            row = cur.fetchone()
            if row is not None:
                base_mass = float(row[0])
    finally:
        conn.close()

    if base_mass is None:
        return None

    if width_value >= 1.2:
        return base_mass

    return round(base_mass * (width_value / 1.2), 1)
