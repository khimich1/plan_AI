#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Номенклатура плит (prays_plity / pb.db) — первый срез декомпозиции kp_db (A2)."""

from __future__ import annotations

import os
import sqlite3
from typing import Dict, List, Optional, Tuple

from core.debug_paths import get_debug_log_path

_PB_DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "pb.db")
_DEBUG_NOMENCLATURE_LOG = get_debug_log_path("debug-00f316.log")
_DEBUG_LOG_A9176E = get_debug_log_path("debug-a9176e.log")
_DEBUG_LOG_B59370 = get_debug_log_path("debug-b59370.log")
_DEBUG_LOG_8E9428 = get_debug_log_path("debug-8e9428.log")


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


def lookup_nomenclature_by_plate_name(
    plate_name: str,
    pb_cur: sqlite3.Cursor,
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Ищет запись в prays_plity по имени плиты.

    Возвращает (canonical_name, nomenclature_id, match_type).
    match_type: "exact" | "like" | None (не найдено).
    """
    pb_cur.execute(
        'SELECT "Уникальный идентификатор (Номенклатура)", "Товар" '
        'FROM prays_plity WHERE "Товар" = ? COLLATE NOCASE',
        (plate_name,),
    )
    row = pb_cur.fetchone()
    if row:
        return row[1], row[0], "exact"

    from core.config_and_data import plate_name_to_prays_variants

    for prays_variant in plate_name_to_prays_variants(plate_name):
        pb_cur.execute(
            'SELECT "Уникальный идентификатор (Номенклатура)", "Товар" '
            'FROM prays_plity WHERE "Товар" = ? COLLATE NOCASE',
            (prays_variant,),
        )
        row = pb_cur.fetchone()
        if row:
            return row[1], row[0], "exact_prays_variant"

    req_len_val = _extract_length_dm_val(plate_name)
    normalized = plate_name.replace("Плиты ", "").replace("Плита ", "")
    pb_cur.execute(
        'SELECT "Уникальный идентификатор (Номенклатура)", "Товар" '
        'FROM prays_plity WHERE "Товар" LIKE ?',
        (f"%{normalized}%",),
    )
    row = pb_cur.fetchone()
    if row:
        can_len_val = _extract_length_dm_val(row[1])
        if req_len_val is not None and can_len_val is not None:
            if abs(req_len_val - can_len_val) > 0.05:
                row = None
    if row:
        return row[1], row[0], "like"

    try:
        with open(_DEBUG_NOMENCLATURE_LOG, "a", encoding="utf-8") as _f:
            _f.write(
                __import__("json").dumps(
                    {
                        "sessionId": "00f316",
                        "hypothesisId": "nomenclature_stage",
                        "location": "kp_db_nomenclature:lookup_nomenclature_by_plate_name",
                        "message": "nomenclature not found after exact, variants, LIKE",
                        "data": {"plate_name": (plate_name or "")[:120], "stage": "not_found"},
                        "timestamp": __import__("time").time() * 1000,
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
    except Exception:
        pass
    return None, None, None


def fill_plate_nomenclature_cache() -> None:
    """Заполняет PLATE_NOMENCLATURE_CACHE из prays_plity для всех позиций в PLATE_LOAD_DETAILS."""
    import logging as _logging

    _log = _logging.getLogger(__name__)

    from core import config_and_data as cfg

    if not os.path.exists(_PB_DB_PATH):
        _log.debug("fill_plate_nomenclature_cache: pb.db не найден, кэш не заполняется")
        return

    pb_conn = sqlite3.connect(_PB_DB_PATH)
    try:
        pb_cur = pb_conn.cursor()
        for key, _qty in cfg.PLATE_LOAD_DETAILS.items():
            if key in cfg.PLATE_NOMENCLATURE_CACHE:
                continue
            length_m, width_m, load_code = key[0], key[1], key[2]
            length_dm_raw = key[3] if len(key) > 3 else cfg.PLATE_LENGTH_DM_RAW.get(key, "")
            plate_name = cfg.make_plate_name(
                length_m, width_m, load_code=load_code, length_dm_raw=length_dm_raw
            )
            canonical_name, nomenclature_id, _ = lookup_nomenclature_by_plate_name(plate_name, pb_cur)
            cfg.PLATE_NOMENCLATURE_CACHE[key] = {
                "canonical_name": canonical_name,
                "nomenclature_id": nomenclature_id,
            }
            if 5.69 <= length_m <= 5.73:
                try:
                    with open(_DEBUG_LOG_A9176E, "a", encoding="utf-8") as _f:
                        _f.write(
                            __import__("json").dumps(
                                {
                                    "sessionId": "a9176e",
                                    "hypothesisId": "H5",
                                    "location": "kp_db_nomenclature:fill_plate_nomenclature_cache",
                                    "message": "57/57,1 cache fill",
                                    "data": {
                                        "key": [length_m, width_m, load_code],
                                        "length_dm_raw": length_dm_raw,
                                        "plate_name": (plate_name or "")[:60],
                                        "canonical_name": (canonical_name or "")[:60]
                                        if canonical_name
                                        else None,
                                    },
                                    "timestamp": __import__("time").time() * 1000,
                                },
                                ensure_ascii=False,
                            )
                            + "\n"
                        )
                except Exception:
                    pass
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
                try:
                    with open(_DEBUG_NOMENCLATURE_LOG, "a", encoding="utf-8") as _f:
                        _f.write(
                            __import__("json").dumps(
                                {
                                    "sessionId": "00f316",
                                    "hypothesisId": "nomenclature_check",
                                    "location": "kp_db_nomenclature:enrich_order_data_with_nomenclature",
                                    "message": "plate skipped, has nomenclature_id from cache",
                                    "data": {
                                        "plate_name": (item.get("name", "") or "")[:120],
                                        "has_nomenclature_id": True,
                                        "match_type": "from_cache",
                                    },
                                    "timestamp": __import__("time").time() * 1000,
                                },
                                ensure_ascii=False,
                            )
                            + "\n"
                        )
                except Exception:
                    pass
                try:
                    with open(_DEBUG_LOG_B59370, "a", encoding="utf-8") as _f:
                        _f.write(
                            __import__("json").dumps(
                                {
                                    "sessionId": "b59370",
                                    "hypothesisId": "H_prays",
                                    "location": "kp_db_nomenclature:enrich_skipped",
                                    "message": "item had nomenclature_id from cache",
                                    "data": {"name": (item.get("name", "") or "")[:60]},
                                    "timestamp": __import__("time").time() * 1000,
                                },
                                ensure_ascii=False,
                            )
                            + "\n"
                        )
                except Exception:
                    pass
                continue

            original_name = item.get("name", "")
            canonical_name, nomenclature_id, match_type = lookup_nomenclature_by_plate_name(
                original_name, pb_cur
            )
            if canonical_name is not None:
                item["name"] = canonical_name
                item["nomenclature_id"] = nomenclature_id
            else:
                item["nomenclature_id"] = None

            if (
                "57,1" in (original_name or "")
                or ("57-12" in (original_name or "") and "57," not in (original_name or ""))
            ):
                try:
                    with open(_DEBUG_LOG_A9176E, "a", encoding="utf-8") as _f:
                        _f.write(
                            __import__("json").dumps(
                                {
                                    "sessionId": "a9176e",
                                    "hypothesisId": "H4",
                                    "location": "kp_db_nomenclature:enrich_order_data_with_nomenclature",
                                    "message": "57/57,1 enrich lookup",
                                    "data": {
                                        "original_name": (original_name or "")[:80],
                                        "canonical_name": (canonical_name or "")[:80]
                                        if canonical_name
                                        else None,
                                        "match_type": match_type,
                                        "replaced": canonical_name is not None,
                                    },
                                    "timestamp": __import__("time").time() * 1000,
                                },
                                ensure_ascii=False,
                            )
                            + "\n"
                        )
                except Exception:
                    pass

            try:
                with open(_DEBUG_NOMENCLATURE_LOG, "a", encoding="utf-8") as _f:
                    _f.write(
                        __import__("json").dumps(
                            {
                                "sessionId": "00f316",
                                "hypothesisId": "nomenclature_check",
                                "location": "kp_db_nomenclature:enrich_order_data_with_nomenclature",
                                "message": "plate nomenclature result",
                                "data": {
                                    "plate_name": (original_name or "")[:120],
                                    "has_nomenclature_id": nomenclature_id is not None,
                                    "match_type": match_type,
                                },
                                "timestamp": __import__("time").time() * 1000,
                            },
                            ensure_ascii=False,
                        )
                        + "\n"
                    )
            except Exception:
                pass

            _no_nom_substrings = ("61,8-5", "45-7", "37,9-9")
            if any(sub in (original_name or "") for sub in _no_nom_substrings):
                try:
                    with open(_DEBUG_LOG_8E9428, "a", encoding="utf-8") as _f:
                        _f.write(
                            __import__("json").dumps(
                                {
                                    "sessionId": "8e9428",
                                    "hypothesisId": "H_no_nomenclature",
                                    "location": "kp_db_nomenclature:enrich_order_data_with_nomenclature",
                                    "message": "lookup prays_plity for 61,8-5 / 45-7 / 37,9-9",
                                    "data": {
                                        "original_name": (original_name or "")[:80],
                                        "canonical_name": (canonical_name[:80] if canonical_name else None),
                                        "nomenclature_id": nomenclature_id,
                                        "match_type": match_type,
                                    },
                                    "timestamp": __import__("time").time() * 1000,
                                },
                                ensure_ascii=False,
                            )
                            + "\n"
                        )
                except Exception:
                    pass

            try:
                with open(_DEBUG_LOG_B59370, "a", encoding="utf-8") as _f:
                    _f.write(
                        __import__("json").dumps(
                            {
                                "sessionId": "b59370",
                                "hypothesisId": "H_prays",
                                "location": "kp_db_nomenclature:enrich_lookup",
                                "message": "prays_plity lookup result",
                                "data": {
                                    "original_name": original_name[:60],
                                    "canonical_name": (canonical_name[:60] if canonical_name else None),
                                    "match_type": match_type,
                                },
                                "timestamp": __import__("time").time() * 1000,
                            },
                            ensure_ascii=False,
                        )
                        + "\n"
                    )
            except Exception:
                pass

    except Exception as e:
        print(f"[DB] ❌ Ошибка при обогащении order_data номенклатурами: {e}")
    finally:
        pb_conn.close()

    return order_data


__all__ = [
    "enrich_order_data_with_nomenclature",
    "fill_plate_nomenclature_cache",
    "lookup_nomenclature_by_plate_name",
]
