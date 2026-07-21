"""Унифицированный построитель RESCUE-дорожек (общий для bot и web).

RESCUE — это дорожки с плитами, которые ОЖИДАЛИСЬ в плане (есть в orders_2d),
но не попали в вывод оптимизатора (всякие edge-cases CP-SAT). Раньше эта
логика жила только в :mod:`bot.handlers.production_execution`, поэтому
web-side планирование молча терял эти плиты, и затем они «всплывали» как
unexpected при complete_day.

Phase 4 (P8): отказ от подсчёта плит по ``all_tracks_list``. Раньше
:func:`compute_missing_counts` шёл по дереву трек-items с восстановлением
identity по размерам (fuzzy-матч strict→tolerance). Это было хрупко (BUG-4,
BUG-5, BUG-6 в боте) — фантомные RESCUE появлялись для secondary cuts с
``target_order_key``, у которых ``load_code`` отличался от родителя.

Теперь источник правды один: ``plate_assignments``. После
:func:`core.plate_attribution.backfill_assignment_identity` каждая запись
имеет ``kp_id``+``plate_name``, и подсчёт сводится к
``Counter[(kp_id, plate_name)]`` — никаких допусков, никакого
``target_order_key``, никакой логики mode/transverse/split.

Контракт:

>>> rescue_tracks, missing_counts, rescue_assignments = build_rescue_tracks(
...     orders_2d, plate_assignments
... )
>>> all_tracks.extend(rescue_tracks)
>>> optimization_result["plate_assignments"].extend(rescue_assignments)

После этого ``count_assigned_plates(optimization_result, all_tracks)`` в
``commit_plan_plates`` видит rescue-плиты через единый источник
(``plate_assignments``) и помечает их в БД. ``all_tracks_list`` остаётся
структурой для визуализации.
"""
from __future__ import annotations

import logging
from collections import Counter
from typing import Any

from core import plate_name as _plate_name
from core.domain.plate_order import normalize_load_code

logger = logging.getLogger(__name__)


MAX_TRACK_LENGTH_M = 101.0
"""Максимальная длина одной дорожки (как в production_execution.py)."""


def _canon_key(length_m: float, width: Any, load_code: Any) -> tuple[float, int, int]:
    """Канонический ключ ``(length_round, width_mm, load_code)``."""
    L = round(float(length_m or 0), 2)
    if width is None:
        W = 1200
    elif isinstance(width, (int, float)) and width > 0 and width < 10:
        W = int(round(float(width) * 1000))
    else:
        W = int(round(float(width or 1200)))
    LC = normalize_load_code(load_code or 8)
    return (L, W, LC)


def _build_order_info_map(
    orders_2d: list[dict[str, Any]],
) -> dict[tuple[int, str], dict[str, Any]]:
    """Карта identity ``(kp_id, plate_name)`` → info-запись с ``qty_remaining``.

    Используется при создании RESCUE-плит, чтобы каждой плите присвоить
    конкретного заказчика (kp_date/customer) и декрементировать спрос.
    """
    info_map: dict[tuple[int, str], dict[str, Any]] = {}
    for order in orders_2d or []:
        kp_id = order.get("kp_id")
        plate_name = str(order.get("plate_name") or "")
        if kp_id is None or not plate_name:
            continue
        identity = (int(kp_id), plate_name)
        if identity in info_map:
            info_map[identity]["qty_remaining"] += int(order.get("qty", 1) or 0)
            if not info_map[identity].get("concrete_grade") and order.get("concrete_grade"):
                info_map[identity]["concrete_grade"] = order.get("concrete_grade")
        else:
            info_map[identity] = {
                "kp_id": int(kp_id),
                "customer": order.get("customer"),
                "kp_date": order.get("kp_date"),
                "plate_name": plate_name,
                "qty_remaining": int(order.get("qty", 1) or 0),
                "length_dm_raw": order.get("length_dm_raw") or "",
                "length": order.get("length"),
                "width": order.get("width"),
                "load_code": normalize_load_code(order.get("load_code", 8)),
                "concrete_grade": order.get("concrete_grade"),
            }
    return info_map


def compute_missing_counts(
    orders_2d: list[dict[str, Any]],
    plate_assignments: list[dict[str, Any]],
) -> tuple[
    dict[tuple[int, str], int],
    dict[tuple[int, str], dict[str, Any]],
]:
    """Считает дефицит плит по точной identity ``(kp_id, plate_name)``.

    Контракт устранения дуальности:
        ``need[(kp_id, plate_name)]`` берётся из ``orders_2d.qty``
        ``have[(kp_id, plate_name)]`` берётся из ``plate_assignments`` со
        всеми источниками (``primary``/``secondary``/``rescue``).

    Возвращает ``(missing, info_map)`` с теми же identity-ключами.
    """
    info_map = _build_order_info_map(orders_2d)

    have: Counter[tuple[int, str]] = Counter()
    for assignment in plate_assignments or []:
        kp_id = assignment.get("kp_id")
        plate_name = assignment.get("plate_name")
        if kp_id is None or not plate_name:
            continue
        have[(int(kp_id), str(plate_name))] += 1

    missing: dict[tuple[int, str], int] = {}
    for identity, info in info_map.items():
        need = int(info.get("qty_remaining", 0) or 0)
        deficit = need - have[identity]
        if deficit > 0:
            missing[identity] = deficit

    return missing, info_map


def _create_rescue_tracks_from_missing(
    missing_counts: dict[tuple[int, str], int],
    info_map: dict[tuple[int, str], dict[str, Any]],
    *,
    max_track_length_m: float = MAX_TRACK_LENGTH_M,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Строит RESCUE-дорожки и параллельно — плоский список ``rescue_assignments``.

    Возвращает кортеж ``(rescue_tracks, rescue_assignments)``. Каждой плите в
    ``rescue_tracks[*].items[*]`` соответствует ровно одна запись
    ``rescue_assignments[i]`` с теми же ``kp_id`` / ``plate_name``.
    """
    rescue_tracks: list[dict[str, Any]] = []
    rescue_assignments: list[dict[str, Any]] = []
    current_track: list[dict[str, Any]] = []
    current_len = 0.0

    def _flush() -> None:
        nonlocal current_track, current_len
        if current_track:
            rescue_tracks.append({
                "items": current_track,
                "length": current_len,
                "load_code": 0,
                "label": "РЕСКЬЮ",
                "max_reinforcement": 0.0,
            })
        current_track = []
        current_len = 0.0

    for identity, qty_missing in missing_counts.items():
        info = info_map.get(identity)
        if not info:
            logger.warning(
                "[RESCUE] Нет info_map для identity %s, пропускаем", identity,
            )
            continue
        length = round(float(info.get("length") or 0), 2)
        width_raw = info.get("width") or 1200
        canon = _canon_key(length, width_raw, info.get("load_code", 8))
        length_canon, width_mm, load_code = canon
        width_m = width_mm / 1000.0
        plate_name = info.get("plate_name") or ""
        kp_id = info.get("kp_id")

        for _ in range(qty_missing):
            if current_track and current_len + length_canon > max_track_length_m:
                _flush()

            display_name = plate_name or _plate_name.make(
                length_canon, width_mm, load_code,
                length_dm_raw=info.get("length_dm_raw") or None,
            )

            current_track.append({
                "length": length_canon,
                "mode": "solid",
                "width": width_m,
                "load_code": load_code,
                "label": display_name,
                "reinforcement": 0,
                "kp_id": kp_id,
                "customer": info.get("customer"),
                "kp_date": info.get("kp_date"),
                "plate_name": plate_name or None,
                "concrete_grade": info.get("concrete_grade"),
                "rescue_order_missing": False,
            })
            rescue_assignments.append({
                "length": length_canon,
                "width": width_mm,
                "load_code": load_code,
                "source": "rescue",
                "kp_id": kp_id,
                "customer": info.get("customer"),
                "kp_date": info.get("kp_date"),
                "plate_name": plate_name or None,
                "identity_match_type": "rescue",
                "rescue_order_missing": False,
                "concrete_grade": info.get("concrete_grade"),
            })
            current_len += length_canon
    _flush()
    return rescue_tracks, rescue_assignments


def build_rescue_tracks(
    orders_2d: list[dict[str, Any]],
    plate_assignments: list[dict[str, Any]],
    *,
    max_track_length_m: float = MAX_TRACK_LENGTH_M,
) -> tuple[
    list[dict[str, Any]],
    dict[tuple[int, str], int],
    list[dict[str, Any]],
]:
    """Возвращает ``(rescue_tracks, missing_counts, rescue_assignments)``.

    Phase 4 (P8): принимает ``plate_assignments`` (не ``all_tracks_list``).
    Caller обязан перед вызовом убедиться, что в ``plate_assignments``
    выполнен backfill identity (см.
    :func:`core.plate_attribution.backfill_assignment_identity`).

    После вызова caller:

    1. ``all_tracks_list.extend(rescue_tracks)`` — добавить дорожки в
       визуализацию.
    2. ``optimization_result["plate_assignments"].extend(rescue_assignments)``
       — расширить flat-список для учёта в ``count_assigned_plates``.
    """
    missing_counts, info_map = compute_missing_counts(orders_2d, plate_assignments)
    if not missing_counts:
        return [], {}, []
    rescue_tracks, rescue_assignments = _create_rescue_tracks_from_missing(
        missing_counts, info_map, max_track_length_m=max_track_length_m
    )
    return rescue_tracks, missing_counts, rescue_assignments


def compute_track_gap_counts(
    orders_2d: list[dict[str, Any]],
    tracks_list: list[dict[str, Any]],
) -> tuple[
    dict[tuple[int, str], int],
    dict[tuple[int, str], dict[str, Any]],
]:
    """Считает, какие заказанные плиты отсутствуют в реальных track-items.

    ``plate_assignments`` может содержать все 244 плиты, но визуальная
    последовательность/разбиение на дорожки иногда возвращает меньше физических
    items. Такие плиты нельзя добавлять в ``rescue_assignments`` повторно —
    они уже учтены optimizer'ом. Нужно добавить только track-items, чтобы
    ``commit_plan_plates`` смог распределить их по дням и записать
    ``kp_plate_id``.
    """
    info_map = _build_order_info_map(orders_2d)
    have: Counter[tuple[int, str]] = Counter()

    for track in tracks_list or []:
        if not isinstance(track, dict):
            continue
        for item in track.get("items") or []:
            if not isinstance(item, dict):
                continue
            kp_id = item.get("kp_id")
            plate_name = item.get("plate_name")
            if kp_id is not None and plate_name:
                have[(int(kp_id), str(plate_name))] += 1
            for sec in item.get("secondary_cuts") or []:
                if not isinstance(sec, dict):
                    continue
                sec_kp_id = sec.get("kp_id")
                sec_plate_name = sec.get("plate_name")
                if sec_kp_id is not None and sec_plate_name:
                    have[(int(sec_kp_id), str(sec_plate_name))] += 1

    missing: dict[tuple[int, str], int] = {}
    for identity, info in info_map.items():
        need = int(info.get("qty_remaining", 0) or 0)
        deficit = need - have[identity]
        if deficit > 0:
            missing[identity] = deficit

    return missing, info_map


def build_track_gap_rescue_tracks(
    orders_2d: list[dict[str, Any]],
    tracks_list: list[dict[str, Any]],
    *,
    max_track_length_m: float = MAX_TRACK_LENGTH_M,
) -> tuple[list[dict[str, Any]], dict[tuple[int, str], int]]:
    """Добавляет RESCUE-дорожки для плит, потерянных между assignments и tracks.

    Возвращает только ``rescue_tracks`` и ``missing_counts``. Flat
    ``rescue_assignments`` намеренно не возвращается: эти плиты уже есть в
    ``optimization_result["plate_assignments"]`` и повторное добавление
    завысило бы ``qty_to_mark``.
    """
    missing_counts, info_map = compute_track_gap_counts(orders_2d, tracks_list)
    if not missing_counts:
        return [], {}
    rescue_tracks, _ = _create_rescue_tracks_from_missing(
        missing_counts,
        info_map,
        max_track_length_m=max_track_length_m,
    )
    return rescue_tracks, missing_counts


__all__ = [
    "build_rescue_tracks",
    "build_track_gap_rescue_tracks",
    "compute_missing_counts",
    "compute_track_gap_counts",
]
