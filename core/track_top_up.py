"""Дозаполнение хвоста дорожек переносом плит с последующих дорожек.

После ``split_sequence_into_tracks`` каждая дорожка стремится заполниться до
``max_length_m`` (101 м), забирая целые solid-плиты с дорожек с большим индексом.
Плиты не копируются — только ``pop`` с донора и ``append`` в хвост получателя.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Any

from core import config_and_data as cfg
from core.config.constants import TRACK_LENGTH_M
from core.track_reconciliation import _track_label, _width_mm

logger = logging.getLogger(__name__)

DEFAULT_EPS_M = 0.01
_INVALID_REINFORCEMENT = {0.0, 999.0}


@dataclass(slots=True)
class TopUpResult:
    """Статистика прохода дозаполнения."""

    moves: int = 0
    meters_added: float = 0.0
    tracks_touched: int = 0
    move_log: list[dict[str, Any]] = field(default_factory=list)


def _item_length_m(item: dict[str, Any]) -> float:
    try:
        return float(item.get("length") or item.get("target_length") or 0)
    except (TypeError, ValueError):
        return 0.0


def _item_width_mm(item: dict[str, Any]) -> int:
    width_raw = item.get("width")
    if width_raw is None:
        width_raw = item.get("main_w")
    return _width_mm(width_raw)


def _item_reinforcement(item: dict[str, Any]) -> float | None:
    try:
        value = float(item.get("reinforcement") or 0)
    except (TypeError, ValueError):
        return None
    if value in _INVALID_REINFORCEMENT:
        return None
    return value


def _track_max_reinforcement(track: dict[str, Any]) -> float:
    cached = track.get("max_reinforcement")
    if cached is not None:
        try:
            value = float(cached)
            if value not in _INVALID_REINFORCEMENT:
                return value
        except (TypeError, ValueError):
            pass
    max_reinf = 0.0
    for item in track.get("items") or []:
        if not isinstance(item, dict):
            continue
        reinf = _item_reinforcement(item)
        if reinf is not None:
            max_reinf = max(max_reinf, reinf)
    return max_reinf


def _item_uid(item: dict[str, Any]) -> str:
    uid = item.get("plate_uid")
    if uid:
        return str(uid)
    kp_id = item.get("kp_id")
    plate_name = item.get("plate_name") or item.get("label")
    length = round(_item_length_m(item), 3)
    width_mm = _item_width_mm(item)
    load_code = cfg.normalize_load_code(item.get("load_code", 8))
    if kp_id is not None and plate_name:
        return f"{kp_id}|{plate_name}|{length}|{width_mm}|{load_code}"
    return f"obj:{id(item)}"


def is_eligible_donor(
    item: dict[str, Any],
    item_index: int,
    *,
    gap: float,
    max_reinforcement: float,
) -> bool:
    """True, если плиту можно перенести в хвост текущей дорожки."""
    if item_index <= 0:
        return False
    if item.get("is_separator"):
        return False
    if item.get("mode") != "solid":
        return False
    if item.get("secondary_cuts"):
        return False

    length = _item_length_m(item)
    if length <= 0 or length > gap + DEFAULT_EPS_M:
        return False

    reinf = _item_reinforcement(item)
    if reinf is None:
        return False
    if max_reinforcement > 0 and reinf > max_reinforcement + DEFAULT_EPS_M:
        return False
    return True


def _donor_sort_key(item: dict[str, Any]) -> tuple[float, float, float]:
    length = _item_length_m(item)
    width_mm = float(_item_width_mm(item))
    reinf = _item_reinforcement(item) or 0.0
    return (-length, -width_mm, -reinf)


def recalc_track(track: dict[str, Any]) -> None:
    """Пересчитывает length, max_reinforcement и label дорожки."""
    items = track.get("items") or []
    total_length = 0.0
    max_reinf = 0.0
    for item in items:
        if not isinstance(item, dict):
            continue
        total_length += _item_length_m(item)
        reinf = _item_reinforcement(item)
        if reinf is not None:
            max_reinf = max(max_reinf, reinf)
    track["length"] = round(total_length, 3)
    track["max_reinforcement"] = max_reinf
    if items:
        track["label"] = _track_label(items)
        load_codes = {
            cfg.normalize_load_code(item.get("load_code", 8))
            for item in items
            if isinstance(item, dict) and item.get("load_code") is not None
        }
        if len(load_codes) == 1:
            track["load_code"] = next(iter(load_codes))


def top_up_tracks_from_following(
    tracks: list[dict[str, Any]],
    *,
    max_length_m: float = TRACK_LENGTH_M,
    eps: float = DEFAULT_EPS_M,
) -> TopUpResult:
    """Дозаполняет дорожки переносом плит с последующих дорожек."""
    result = TopUpResult()
    if not tracks:
        return result

    used_ids: set[str] = set()
    touched: set[int] = set()

    for i, track in enumerate(tracks):
        if not isinstance(track, dict):
            continue
        recalc_track(track)
        gap = max_length_m - float(track.get("length") or 0)
        if gap <= eps:
            continue

        while gap > eps:
            candidates: list[tuple[int, int, dict[str, Any]]] = []
            max_reinf = _track_max_reinforcement(track)

            for j in range(i + 1, len(tracks)):
                donor = tracks[j]
                if not isinstance(donor, dict):
                    continue
                for k, item in enumerate(donor.get("items") or []):
                    if not isinstance(item, dict):
                        continue
                    uid = _item_uid(item)
                    if uid in used_ids:
                        continue
                    if not is_eligible_donor(
                        item,
                        k,
                        gap=gap,
                        max_reinforcement=max_reinf,
                    ):
                        continue
                    candidates.append((j, k, item))

            if not candidates:
                break

            candidates.sort(key=lambda row: _donor_sort_key(row[2]))
            donor_idx, item_idx, item = candidates[0]
            donor_track = tracks[donor_idx]
            donor_items = donor_track.get("items") or []
            moved = donor_items.pop(item_idx)

            moved["top_up"] = True
            moved["top_up_from_track_idx"] = donor_idx

            track.setdefault("items", []).append(moved)
            uid = _item_uid(moved)
            used_ids.add(uid)

            recalc_track(donor_track)
            recalc_track(track)

            added = _item_length_m(moved)
            gap = max_length_m - float(track.get("length") or 0)
            result.moves += 1
            result.meters_added += added
            touched.add(i)
            result.move_log.append(
                {
                    "to_track": i,
                    "from_track": donor_idx,
                    "length": added,
                    "plate_uid": uid,
                }
            )

    result.tracks_touched = len(touched)
    if result.moves:
        logger.info(
            "[TOP_UP] Переносов=%s, добрано %.2f м, дорожек=%s",
            result.moves,
            result.meters_added,
            result.tracks_touched,
        )
    return result


__all__ = [
    "TopUpResult",
    "is_eligible_donor",
    "recalc_track",
    "top_up_tracks_from_following",
]
