#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Номенклатура плит (prays_plity / pb.db) — первый срез декомпозиции kp_db (A2)."""

from __future__ import annotations

import os
import sqlite3
from typing import Dict, List, Optional, Tuple

_PB_DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "pb.db")


def _extract_length_dm_val(name: str) -> Optional[float]:
    """Возвращает числовое значение длины в дм из марки плиты или None."""
    import re as _re

    m = _re.search(r"П[БК]\s*([\d,\.]+)\s*-", str(name or ""))
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", "."))
    except ValueError:
        return None


def _fetch_plate_row(pb_cur: sqlite3.Cursor, where_sql: str, param: str):
    pb_cur.execute(
        'SELECT "Уникальный идентификатор (Номенклатура)", "Товар" '
        f"FROM prays_plity WHERE {where_sql} "
        'ORDER BY "Уникальный идентификатор (Номенклатура)"',  # name dupes: smallest 1C GUID
        (param,),
    )
    return pb_cur.fetchone()


def _count_plate_guids(pb_cur: sqlite3.Cursor, where_sql: str, param: str) -> int:
    pb_cur.execute(
        'SELECT COUNT(DISTINCT "Уникальный идентификатор (Номенклатура)") '
        f"FROM prays_plity WHERE {where_sql}",
        (param,),
    )
    row = pb_cur.fetchone()
    return int(row[0]) if row and row[0] is not None else 0


def count_nomenclature_by_plate_name(
    plate_name: str,
    pb_cur: sqlite3.Cursor,
) -> int:
    """Число различных GUID в prays_plity по тому же пути, что lookup (exact/variant).

    count > 1 — дубль 1С: без plate_guid_choice резолвер должен вернуть missing.
    """
    n = _count_plate_guids(pb_cur, '"Товар" = ? COLLATE NOCASE', plate_name)
    if n:
        return n

    from core.config_and_data import plate_name_to_prays_variants

    for prays_variant in plate_name_to_prays_variants(plate_name):
        n = _count_plate_guids(pb_cur, '"Товар" = ? COLLATE NOCASE', prays_variant)
        if n:
            return n
    return 0


def lookup_nomenclature_by_plate_name(
    plate_name: str,
    pb_cur: sqlite3.Cursor,
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Ищет запись в prays_plity по имени плиты.

    Возвращает (canonical_name, nomenclature_id, match_type).
    match_type: "exact" | "like" | None (не найдено).
    """
    row = _fetch_plate_row(pb_cur, '"Товар" = ? COLLATE NOCASE', plate_name)
    if row:
        return row[1], row[0], "exact"

    from core.config_and_data import plate_name_to_prays_variants

    for prays_variant in plate_name_to_prays_variants(plate_name):
        row = _fetch_plate_row(pb_cur, '"Товар" = ? COLLATE NOCASE', prays_variant)
        if row:
            return row[1], row[0], "exact_prays_variant"

    req_len_val = _extract_length_dm_val(plate_name)
    normalized = plate_name.replace("Плиты ", "").replace("Плита ", "")
    row = _fetch_plate_row(pb_cur, '"Товар" LIKE ?', f"%{normalized}%")
    if row:
        can_len_val = _extract_length_dm_val(row[1])
        if req_len_val is not None and can_len_val is not None:
            if abs(req_len_val - can_len_val) > 0.05:
                row = None
    if row:
        return row[1], row[0], "like"

    return None, None, None


def fill_plate_nomenclature_cache() -> None:
    """Заполняет PLATE_NOMENCLATURE_CACHE из prays_plity для всех позиций в PLATE_LOAD_DETAILS."""
    import logging as _logging

    _log = _logging.getLogger(__name__)

    from core.config_and_data import make_plate_name
    from core.plate_runtime_state import get_plate_mutable_runtime

    _rt = get_plate_mutable_runtime()

    if not os.path.exists(_PB_DB_PATH):
        _log.debug("fill_plate_nomenclature_cache: pb.db не найден, кэш не заполняется")
        return

    pb_conn = sqlite3.connect(_PB_DB_PATH)
    try:
        pb_cur = pb_conn.cursor()
        for key, _qty in _rt.plate_load_details.items():
            if key in _rt.plate_nomenclature_cache:
                continue
            length_m, width_m, load_code = key[0], key[1], key[2]
            length_dm_raw = key[3] if len(key) > 3 else _rt.plate_length_dm_raw.get(key, "")
            plate_name = make_plate_name(
                length_m, width_m, load_code=load_code, length_dm_raw=length_dm_raw
            )
            canonical_name, nomenclature_id, _ = lookup_nomenclature_by_plate_name(plate_name, pb_cur)
            _rt.plate_nomenclature_cache[key] = {
                "canonical_name": canonical_name,
                "nomenclature_id": nomenclature_id,
            }
            _log.debug(
                f"Кэш: {plate_name!r} → canonical={canonical_name!r}, id={nomenclature_id!r}"
            )
    except Exception as e:
        _log.warning(f"fill_plate_nomenclature_cache: ошибка: {e}")
    finally:
        pb_conn.close()


def enrich_order_data_with_nomenclature(order_data: List[Dict]) -> List[Dict]:
    """Обогащает order_data названиями и nomenclature_id из prays_plity (pb.db)."""
    if not os.path.exists(_PB_DB_PATH):
        print(f"[DB] ⚠️ Файл pb.db не найден по пути: {_PB_DB_PATH}")
        return order_data

    pb_conn = sqlite3.connect(_PB_DB_PATH)
    try:
        pb_cur = pb_conn.cursor()

        for item in order_data:
            if item.get("nomenclature_id") is not None:
                continue

            original_name = item.get("name", "")
            canonical_name, nomenclature_id, _match_type = lookup_nomenclature_by_plate_name(
                original_name, pb_cur
            )
            if canonical_name is not None:
                item["name"] = canonical_name
                item["nomenclature_id"] = nomenclature_id
            else:
                item["nomenclature_id"] = None

    except Exception as e:
        print(f"[DB] ❌ Ошибка при обогащении order_data номенклатурами: {e}")
    finally:
        pb_conn.close()

    return order_data


__all__ = [
    "count_nomenclature_by_plate_name",
    "enrich_order_data_with_nomenclature",
    "fill_plate_nomenclature_cache",
    "lookup_nomenclature_by_plate_name",
]
