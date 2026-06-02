# -*- coding: utf-8 -*-
"""Чистые утилиты layout_sequence (ключи карты армирования, разделители, разбиение групп)."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Literal

from viz_modules.layout_sequence.debug_trace import append_json_line


def canonical_load_code(load_code: object) -> int | None:
    """Единый ключ нагрузки для карты: 12.5 и 13 дают 13 (как в БД pb_reinforcement_series)."""
    if load_code is None:
        return None
    try:
        return int(float(load_code) + 0.5)
    except Exception:
        return None


def get_reinforcement_from_map(
    reinforcement_map: dict,
    length: float,
    width_mm: float | int,
    load_code: int | float | None = None,
    *,
    hypothesis_debug_log: Path | None = None,
) -> float | None:
    """
    Получает армирование из карты по (length, width_mm, load_code).

    Карта имеет ключи (length, width_mm, canonical_load_code).
    """
    if load_code is not None:
        canonical = canonical_load_code(load_code)
        if canonical is not None:
            result = reinforcement_map.get((length, width_mm, canonical))
            if result is None and canonical == 12:
                result = reinforcement_map.get((length, width_mm, 13))
            if hypothesis_debug_log is not None and abs(length - 7.1) < 0.05 and canonical in (12, 13):
                try:
                    keys_71 = [
                        (k[0], k[1], k[2])
                        for k in list(reinforcement_map.keys())[:50]
                        if abs(k[0] - 7.1) < 0.05
                    ]
                    append_json_line(
                        hypothesis_debug_log,
                        {
                            "hypothesisId": "H71lookup",
                            "location": "layout_sequence:get_reinforcement_from_map",
                            "message": "71-12п поиск в карте",
                            "data": {
                                "length": length,
                                "width_mm": width_mm,
                                "load_code": load_code,
                                "canonical": canonical,
                                "key": (round(length, 3), width_mm, canonical),
                                "result": result,
                                "keys_71_in_map": keys_71[:10],
                            },
                            "timestamp": __import__("time").time() * 1000,
                        },
                    )
                except Exception:
                    pass
            return result
    for (l, w, lc), reinforcement in reinforcement_map.items():
        if abs(l - length) < 0.01 and w == width_mm:
            return reinforcement
    return None


def ensure_sequence_layout_uid(sequence: list | None, prefix: str = "seq") -> None:
    """Гарантирует наличие identity у каждого root item sequence."""
    for idx, item in enumerate(sequence or []):
        if not isinstance(item, dict):
            continue
        if item.get("layout_uid"):
            continue
        unit_id = item.get("unit_id")
        item["layout_uid"] = str(unit_id) if unit_id else f"{prefix}:{idx}"


def _solid_plate_reinforcement(
    plate: dict[str, Any],
    reinforcement_map: dict,
) -> tuple[float, float, int]:
    length = plate["lengths"][0] if plate.get("lengths") else 6.0
    width_mm = plate["width"]
    reinforcement = get_reinforcement_from_map(reinforcement_map, length, width_mm) or float(
        plate.get("reinforcement", 999.0)
    )
    return length, width_mm, reinforcement


def choose_closest_solid(
    solid_list: list[dict[str, Any]],
    target_reinf: float,
    reinforcement_map: dict,
    *,
    log: logging.Logger | None = None,
) -> int | None:
    """Индекс целой с минимальным |reinforcement - target_reinf| (tie-break: меньший индекс)."""
    if not solid_list:
        return None
    candidates: list[dict[str, Any]] = []
    for idx, plate in enumerate(solid_list):
        length, width_mm, reinforcement = _solid_plate_reinforcement(plate, reinforcement_map)
        candidates.append(
            {
                "index": idx,
                "length": length,
                "width_mm": width_mm,
                "reinforcement": reinforcement,
            }
        )
    best = min(
        candidates,
        key=lambda x: (abs(float(x["reinforcement"]) - float(target_reinf)), x["index"]),
    )
    if log is not None:
        log.info(
            "[VISUAL] ✅ Выбрана ближайшая целая к %.1f кг/м: %.2fм x %sмм, армирование %.1f кг/м",
            target_reinf,
            best["length"],
            best["width_mm"],
            best["reinforcement"],
        )
    return best["index"]


def choose_best_separator(
    solid_list: list[dict[str, Any]],
    next_group: list[dict[str, Any]],
    reinforcement_map: dict,
    *,
    reinforcement_order: Literal["asc", "desc"] = "asc",
    log: logging.Logger | None = None,
) -> int | None:
    """
    Выбирает плиту-разделитель: при asc — min tier + скачок к next_group;
    при desc — глобально ближайшая по армированию к next_group.
    """
    if not solid_list:
        return None
    next_reinf: float | None = None
    if next_group:
        try:
            next_reinf = float(next_group[0].get("reinforcement", 999.0))
        except (TypeError, ValueError):
            next_reinf = None

    if reinforcement_order == "desc":
        target = next_reinf if next_reinf is not None else min(
            _solid_plate_reinforcement(p, reinforcement_map)[2] for p in solid_list
        )
        return choose_closest_solid(solid_list, target, reinforcement_map, log=log)

    candidates: list[dict[str, Any]] = []
    for idx, plate in enumerate(solid_list):
        length, width_mm, reinforcement = _solid_plate_reinforcement(plate, reinforcement_map)
        candidates.append({"index": idx, "length": length, "width_mm": width_mm, "reinforcement": reinforcement})
    tier_reinf = min(c["reinforcement"] for c in candidates)
    tier = [c for c in candidates if c["reinforcement"] == tier_reinf]
    if next_reinf is not None:
        best = min(tier, key=lambda x: (abs(float(x["reinforcement"]) - next_reinf), x["index"]))
    else:
        best = min(tier, key=lambda x: x["index"])
    msg = (
        f"[VISUAL] ✅ Выбран разделитель с мин. армированием: {best['length']:.2f}м x {best['width_mm']}мм, "
        f"армирование {best['reinforcement']:.1f} кг/м"
    )
    if log is not None:
        log.info(msg)
    return best["index"]


def split_group_into_subgroups(
    cut_group: list[dict[str, Any]],
    max_length: float = 90.0,
    *,
    log: logging.Logger | None = None,
) -> list[list[dict[str, Any]]]:
    """
    Разбивает группу резов на подгруппы по max_length метров.
    """
    subgroups: list[list[dict[str, Any]]] = []
    current_subgroup: list[dict[str, Any]] = []
    current_length = 0.0

    if log is not None:
        log.info(f"[VISUAL] Разбиваю группу на подгруппы (макс {max_length}м)...")

    for cut in cut_group:
        qty = cut["qty"]
        lengths = cut.get("lengths", [])
        for i in range(qty):
            length = lengths[i] if i < len(lengths) else (lengths[0] if lengths else 6.0)
            if current_length + length > max_length and current_subgroup:
                subgroups.append(current_subgroup)
                if log is not None:
                    log.info(
                        f"[VISUAL]   Подгруппа #{len(subgroups)} закрыта: {current_length:.1f}м "
                        f"({len(current_subgroup)} записей)"
                    )
                current_subgroup = []
                current_length = 0.0
            single_cut = cut.copy()
            single_cut["qty"] = 1
            single_cut["lengths"] = [length]
            current_subgroup.append(single_cut)
            current_length += length

    if current_subgroup:
        subgroups.append(current_subgroup)
        if log is not None:
            log.info(
                f"[VISUAL]   Подгруппа #{len(subgroups)} закрыта: {current_length:.1f}м "
                f"({len(current_subgroup)} записей)"
            )

    if log is not None:
        log.info(f"[VISUAL] ✓ Группа разбита на {len(subgroups)} подгрупп")
    return subgroups
