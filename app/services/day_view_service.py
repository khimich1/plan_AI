"""Сборка детального вида дня для веб-клиента.

Повторяет логику из `bot/handlers/production_day_view.py::process_day_selection`:
агрегирует дорожки из всех планов на дату и для каждой плиты ищет информацию
в lookup-таблицах с tolerance 0.03 м (fuzzy-поиск).
"""
from __future__ import annotations

import copy
import logging
from typing import Any

from bot.handlers import plan_manager

logger = logging.getLogger(__name__)

FUZZY_TOLERANCE_M = 0.03


def _reinforcement_to_load_code(reinforcement: float) -> int:
    if reinforcement <= 0:
        return 8
    if reinforcement < 8:
        return 6
    if reinforcement < 12:
        return 8
    if reinforcement < 15:
        return 10
    return 12


def _make_plate_name(length_m: float, width_mm: int, load_code: int) -> str:
    length_dm = length_m * 10
    if abs(length_dm - round(length_dm)) < 0.01:
        length_str = str(int(round(length_dm)))
    else:
        length_str = f"{length_dm:.1f}".rstrip("0").rstrip(".").replace(".", ",")

    if width_mm == 1200:
        width_str = "12"
    else:
        width_dm = width_mm / 100.0
        if abs(width_dm - int(width_dm)) < 0.01:
            width_str = str(int(width_dm))
        else:
            width_str = str(width_dm).replace(".", ",")

    return f"ПБ {length_str}-{width_str}-{load_code}п"


def _build_smart_lookup(
    plate_lookup_exact: dict,
    plate_lookup_by_length: dict,
):
    """Создаёт функцию fuzzy-поиска, работающую с копией lookup-таблиц.

    Возвращает запись (словарь) с ключами `customer`, `plate_name`,
    `kp_date`, `kp_id`, `reinforcement` или дефолт, если ничего не найдено.
    Каждый вызов списывает одну плиту (уменьшает `qty_remaining`), поэтому
    lookup должен быть копией.
    """
    formovka_exact = copy.deepcopy(plate_lookup_exact)
    formovka_by_length = copy.deepcopy(plate_lookup_by_length)

    def lookup(length_m: float, width_mm: int) -> dict[str, Any]:
        rounded_length = round(length_m, 2)

        key = (rounded_length, width_mm)
        entries = formovka_exact.get(key, [])
        for entry in entries:
            if entry.get("qty_remaining", 0) > 0:
                entry["qty_remaining"] -= 1
                return entry.copy()

        if width_mm < 1200:
            key_original = (rounded_length, 1200)
            for entry in formovka_exact.get(key_original, []):
                if entry.get("qty_remaining", 0) > 0:
                    entry["qty_remaining"] -= 1
                    return entry.copy()

        for lookup_key, entries in formovka_exact.items():
            key_length, key_width = lookup_key
            if key_width != width_mm and key_width != 1200:
                continue
            if abs(key_length - rounded_length) <= FUZZY_TOLERANCE_M:
                for entry in entries:
                    if entry.get("qty_remaining", 0) > 0:
                        entry["qty_remaining"] -= 1
                        return entry.copy()

        entries = formovka_by_length.get(rounded_length, [])
        for entry in entries:
            if entry.get("qty_remaining", 0) > 0:
                entry["qty_remaining"] -= 1
                return entry.copy()

        for lookup_length, entries in formovka_by_length.items():
            if abs(lookup_length - rounded_length) <= FUZZY_TOLERANCE_M:
                for entry in entries:
                    if entry.get("qty_remaining", 0) > 0:
                        entry["qty_remaining"] -= 1
                        return entry.copy()

        return {
            "kp_id": None,
            "kp_date": "неизвестно",
            "customer": "неизвестно",
            "plate_name": "",
            "reinforcement": 0,
        }

    return lookup


def _iter_plate_items(track: dict):
    """Перебирает основные плиты + вторичные резы (остатки) внутри дорожки.

    Возвращает кортежи `(length_m, width_mm, is_secondary, label_hint)`.
    """
    for item in track.get("items", []) or []:
        if item is None:
            continue
        length = item.get("length")
        if not length:
            continue

        mode = item.get("mode", "solid")
        if mode == "transverse" and item.get("width"):
            width_mm = round(item["width"] * 1000)
        elif mode == "split" and item.get("main_w"):
            width_mm = round(item["main_w"] * 1000)
        else:
            width_mm = 1200

        yield float(length), int(width_mm), False, item, None

        for sec in item.get("secondary_cuts") or []:
            sec_width_m = sec.get("width", 0)
            if sec_width_m <= 0:
                continue
            sec_width_mm = round(sec_width_m * 1000)
            sec_length = sec.get("target_length") or length
            label_hint = None
            if sec.get("label"):
                label_hint = sec["label"].replace("О ", "").strip()
            yield float(sec_length), int(sec_width_mm), True, item, label_hint


def _aggregate_plates_for_track(track: dict, lookup) -> list[dict[str, Any]]:
    plates: list[dict[str, Any]] = []
    is_rescue = track.get("label") == "РЕСКЬЮ"

    for length_m, width_mm, _is_secondary, parent_item, label_hint in _iter_plate_items(track):
        info = lookup(length_m, width_mm)

        plate_name = info.get("plate_name") or ""
        if is_rescue and parent_item and (parent_item.get("plate_name") or parent_item.get("label")):
            plate_name = parent_item.get("plate_name") or parent_item.get("label", "")
        if not plate_name and label_hint:
            plate_name = label_hint

        reinforcement = float(info.get("reinforcement") or 0)
        load_code = int(info.get("load_code") or _reinforcement_to_load_code(reinforcement))

        if not plate_name:
            plate_name = _make_plate_name(length_m, width_mm, load_code)

        kp_id = info.get("kp_id") or (parent_item.get("kp_id") if parent_item else None)

        existing = next(
            (
                p
                for p in plates
                if round(p["length_m"], 2) == round(length_m, 2)
                and p["width_mm"] == width_mm
                and abs(p["reinforcement"] - reinforcement) < 0.1
                and p["kp_date"] == info.get("kp_date", "неизвестно")
                and p["customer"] == info.get("customer", "неизвестно")
                and p.get("kp_id") == kp_id
                and p["plate_name"] == plate_name
            ),
            None,
        )
        if existing:
            existing["qty"] += 1
            continue

        plates.append(
            {
                "length_m": round(length_m, 3),
                "width_mm": int(width_mm),
                "qty": 1,
                "reinforcement": reinforcement,
                "kp_date": info.get("kp_date", "неизвестно"),
                "customer": info.get("customer", "неизвестно"),
                "kp_id": kp_id,
                "plate_name": plate_name,
                "load_code": load_code,
            }
        )

    plates.sort(key=lambda p: p["length_m"], reverse=True)
    return plates


def _plan_completion_map(source_plan_ids: list[str], date_key: str) -> dict[str, bool]:
    result: dict[str, bool] = {}
    for plan_id in source_plan_ids:
        plan = plan_manager.load_plan(plan_id)
        if not plan:
            continue
        day = plan.get("days", {}).get(date_key, {})
        result[plan_id] = bool(day.get("completed"))
    return result


def build_day_view_detail(date_key: str) -> dict | None:
    """Собирает детальный вид дня: дорожки с плитами, сгруппированные по планам.

    Возвращает:
        - ``None``, если даты нет ни в одном плане (фронт получит 404);
        - структуру с пустым ``plans`` и ``total_tracks=0``, если дата есть,
          но массив ``tracks`` в плане пустой. Это нужно, чтобы фронт мог
          отличить «дня нет» от «день есть, но без дорожек» и показать
          info-алерт, а не generic-ошибку.
    """
    multi = plan_manager.get_tracks_for_date_from_all_plans(date_key)
    if not multi:
        return None

    tracks: list[dict] = multi.get("tracks") or []
    if not tracks:
        return {
            "date": date_key,
            "plans": [],
            "plans_count": 0,
            "total_tracks": 0,
        }

    lookup = _build_smart_lookup(
        multi.get("plate_lookup_exact", {}),
        multi.get("plate_lookup_by_length", {}),
    )

    source_plans: list[str] = multi.get("source_plans") or []
    completion = _plan_completion_map(source_plans, date_key)

    plan_blocks: dict[str, dict[str, Any]] = {}
    plan_order: list[str] = []

    for track_index, track in enumerate(tracks, start=1):
        plan_id = track.get("source_plan_id") or "unknown"
        plan_name = track.get("source_plan_name") or plan_id

        block = plan_blocks.get(plan_id)
        if block is None:
            block = {
                "plan_id": plan_id,
                "plan_name": plan_name,
                "completed": completion.get(plan_id, False),
                "tracks": [],
            }
            plan_blocks[plan_id] = block
            plan_order.append(plan_id)

        plates_info = _aggregate_plates_for_track(track, lookup)

        block["tracks"].append(
            {
                "track_number": track_index,
                "length": track.get("length"),
                "max_reinforcement": float(track.get("max_reinforcement") or 0),
                "label": track.get("label"),
                "source_plan_id": plan_id,
                "source_plan_name": plan_name,
                "plates_info": plates_info,
            }
        )

    return {
        "date": date_key,
        "plans": [plan_blocks[pid] for pid in plan_order],
        "plans_count": len(plan_order),
        "total_tracks": len(tracks),
    }
