#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Номенклатура плит (prays_plity / pb.db) — первый срез декомпозиции kp_db (A2)."""

from __future__ import annotations

import os
import re
import sqlite3
from dataclasses import dataclass
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


PLATE_FOUND = "found"
PLATE_DUP = "dup"
PLATE_VARIANTS = "variants"
PLATE_NONE = "none"

_LOAD_SUFFIX_RE = re.compile(
    r"^(?P<body>.*-)(?P<load>\d+(?:[.,]\d+)?)п\s*$",
    re.IGNORECASE,
)
_NAME_COLUMN = '"Товар"'
_GUID_COLUMN = '"Уникальный идентификатор (Номенклатура)"'


@dataclass(frozen=True)
class PlateCardMatch:
    """Результат разбора плиты по prays_plity.

    status: found (один GUID), dup (одно имя, несколько GUID),
    variants (LIKE дал несколько разных карточек, сам не выбирает), none.
    guid при dup — наименьший (стабильный для lookup); выбор делает человек.
    """

    status: str
    canonical_name: Optional[str] = None
    guid: Optional[str] = None
    guid_count: int = 0
    candidates: Tuple[str, ...] = ()
    match_type: Optional[str] = None


def plate_load_12_alias(plate_name: str) -> Optional[str]:
    """Хвост нагрузки 12 или 12,5 в любом написании → имя с хвостом «12,5п».

    Уже «12,5п» и любая другая нагрузка возвращают None.
    Спека: ai_docs/specs/plate-load-12-guid-alias.md, D-load12.
    """
    text = str(plate_name or "").strip()
    match = _LOAD_SUFFIX_RE.match(text)
    if not match:
        return None
    try:
        value = float(match.group("load").replace(",", "."))
    except ValueError:
        return None
    if abs(value - 12.0) > 1e-6 and abs(value - 12.5) > 1e-6:
        return None
    alias = f"{match.group('body')}12,5п"
    if alias == text:
        return None
    return alias


def load_class_from_plate_name(name: str) -> Optional[int]:
    """Нагрузка из имени карточки плиты: «…12,5п» → 1250, «…12п» → 1200, «…8п» → 800.

    Нет хвоста «…п» (в т. ч. особая карточка «…12п Пв-7.1») → None.
    Не parse_load_code_from_name: тот округляет 12,5 до 13.
    """
    match = _LOAD_SUFFIX_RE.match(str(name or "").strip())
    if not match:
        return None
    try:
        value = float(match.group("load").replace(",", "."))
    except ValueError:
        return None
    return int(round(value * 100))


def _select_cards(pb_cur: sqlite3.Cursor, where_sql: str, param: str) -> List[Tuple[str, str]]:
    pb_cur.execute(
        f"SELECT {_GUID_COLUMN}, {_NAME_COLUMN} FROM prays_plity WHERE {where_sql} "
        f"ORDER BY {_GUID_COLUMN}",  # name dupes: smallest 1C GUID
        (param,),
    )
    return pb_cur.fetchall()


def _card_match(rows: List[Tuple[str, str]], match_type: str) -> PlateCardMatch:
    guid_count = len({row[0] for row in rows})
    return PlateCardMatch(
        status=PLATE_FOUND if guid_count == 1 else PLATE_DUP,
        canonical_name=rows[0][1],
        guid=rows[0][0],
        guid_count=guid_count,
        match_type=match_type,
    )


def _match_exact_or_variant(
    pb_cur: sqlite3.Cursor,
    name: str,
    exact_type: str,
    variant_type: str,
) -> Optional[PlateCardMatch]:
    from core.config_and_data import plate_name_to_prays_variants

    attempts = [(name, exact_type)]
    attempts += [(variant, variant_type) for variant in plate_name_to_prays_variants(name)]
    for candidate, match_type in attempts:
        rows = _select_cards(pb_cur, f"{_NAME_COLUMN} = ? COLLATE NOCASE", candidate)
        if rows:
            return _card_match(rows, match_type)
    return None


def _length_matches(req_len_val: Optional[float], card_name: str) -> bool:
    can_len_val = _extract_length_dm_val(card_name)
    if req_len_val is None or can_len_val is None:
        return True
    return abs(req_len_val - can_len_val) <= 0.05


def _match_like(pb_cur: sqlite3.Cursor, plate_name: str) -> PlateCardMatch:
    normalized = plate_name.replace("Плиты ", "").replace("Плита ", "")
    if not normalized.strip():
        return PlateCardMatch(status=PLATE_NONE)
    req_len_val = _extract_length_dm_val(plate_name)
    rows = [
        row
        for row in _select_cards(pb_cur, f"{_NAME_COLUMN} LIKE ?", f"%{normalized}%")
        if _length_matches(req_len_val, row[1])
    ]
    names = sorted({row[1] for row in rows})
    if not names:
        return PlateCardMatch(status=PLATE_NONE)
    if len(names) > 1:
        return PlateCardMatch(status=PLATE_VARIANTS, candidates=tuple(names))
    return _card_match([row for row in rows if row[1] == names[0]], "like")


def find_plate_card(pb_cur: sqlite3.Cursor, plate_name: str) -> PlateCardMatch:
    """Единый разбор карточки плиты (D-load12).

    Порядок: точное имя → варианты длины/ширины → «12 → 12,5п» (точное имя и те же
    варианты) → LIKE. LIKE с несколькими карточками сам не выбирает (variants).
    Правило одностороннее: набрано 12,5п — на 12п не откатываемся.
    """
    name = str(plate_name or "")
    found = _match_exact_or_variant(pb_cur, name, "exact", "exact_prays_variant")
    if found is not None:
        return found
    alias = plate_load_12_alias(name)
    if alias is not None:
        found = _match_exact_or_variant(
            pb_cur, alias, "load12_alias", "load12_alias_prays_variant"
        )
        if found is not None:
            return found
    return _match_like(pb_cur, name)


def count_nomenclature_by_plate_name(
    plate_name: str,
    pb_cur: sqlite3.Cursor,
) -> int:
    """Число различных GUID карточки, найденной тем же правилом, что lookup.

    count > 1 — дубль 1С: без plate_guid_choice резолвер должен вернуть missing.
    """
    match = find_plate_card(pb_cur, plate_name)
    if match.status in (PLATE_FOUND, PLATE_DUP):
        return match.guid_count
    return 0


def lookup_nomenclature_by_plate_name(
    plate_name: str,
    pb_cur: sqlite3.Cursor,
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Ищет запись в prays_plity по имени плиты.

    Возвращает (canonical_name, nomenclature_id, match_type).
    match_type: "exact" | "exact_prays_variant" | "load12_alias" |
    "load12_alias_prays_variant" | "like" | None (не найдено или несколько карточек).
    """
    match = find_plate_card(pb_cur, plate_name)
    if match.status in (PLATE_FOUND, PLATE_DUP):
        return match.canonical_name, match.guid, match.match_type
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
                card_load_class = load_class_from_plate_name(canonical_name)
                if card_load_class is not None and canonical_name != original_name:
                    item["load_class"] = card_load_class
            else:
                item["nomenclature_id"] = None

    except Exception as e:
        print(f"[DB] ❌ Ошибка при обогащении order_data номенклатурами: {e}")
    finally:
        pb_conn.close()

    return order_data


__all__ = [
    "PLATE_DUP",
    "PLATE_FOUND",
    "PLATE_NONE",
    "PLATE_VARIANTS",
    "PlateCardMatch",
    "count_nomenclature_by_plate_name",
    "enrich_order_data_with_nomenclature",
    "fill_plate_nomenclature_cache",
    "find_plate_card",
    "load_class_from_plate_name",
    "lookup_nomenclature_by_plate_name",
    "plate_load_12_alias",
]
