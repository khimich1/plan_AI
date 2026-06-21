"""Reconcile optimizer assignments with physical track items.

The optimizer's ``plate_assignments`` is the source of truth for which
physical plates must be present in the production plan.  Layout and track
splitting may reorder plates, but they must not silently drop them.
"""
from __future__ import annotations

import logging
from collections import Counter
from dataclasses import dataclass, field
from typing import Any

from core import plate_name as _plate_name
from core.config_and_data import format_reinforcement_from_load_code
from core.domain.plate_order import normalize_load_code

logger = logging.getLogger(__name__)

MAX_TRACK_LENGTH_M = 101.0


@dataclass(slots=True)
class ReconcileStats:
    assignments_count: int = 0
    track_items_count: int = 0
    deficit_count: int = 0
    extra_track_count: int = 0
    missing_uids: list[str] = field(default_factory=list)
    missing_counts: dict[tuple[Any, ...], int] = field(default_factory=dict)


@dataclass(slots=True)
class ReconcileResult:
    extra_tracks: list[dict[str, Any]]
    stats: ReconcileStats


class TrackReconciliationError(RuntimeError):
    """Raised when strict reconciliation detects dropped physical plates."""


def _width_mm(value: Any, default: int = 1200) -> int:
    if value is None:
        return default
    try:
        width = float(value)
    except (TypeError, ValueError):
        return default
    if 0 < width < 10:
        return int(round(width * 1000))
    return int(round(width))


def _load_code(value: Any) -> int:
    return normalize_load_code(value or 8)


def _assignment_key(item: dict[str, Any]) -> tuple[Any, ...] | None:
    kp_id = item.get("kp_id")
    plate_name = item.get("plate_name")
    if kp_id is None or not plate_name:
        return None
    length = round(float(item.get("length") or item.get("target_length") or 0), 3)
    width = _width_mm(item.get("width") if item.get("source") != "primary" else item.get("width"))
    return (
        int(kp_id),
        str(plate_name),
        length,
        width,
        _load_code(item.get("load_code")),
        str(item.get("source") or ""),
    )


def _track_item_key(item: dict[str, Any], *, source_hint: str = "") -> tuple[Any, ...] | None:
    kp_id = item.get("kp_id")
    plate_name = item.get("plate_name")
    if kp_id is None or not plate_name:
        return None
    length = round(float(item.get("length") or item.get("target_length") or 0), 3)
    width_raw = item.get("width")
    if width_raw is None:
        width_raw = item.get("main_w")
    return (
        int(kp_id),
        str(plate_name),
        length,
        _width_mm(width_raw),
        _load_code(item.get("load_code")),
        source_hint,
    )


def _iter_physical_track_items(tracks_list: list[dict[str, Any]] | None):
    for track in tracks_list or []:
        if not isinstance(track, dict):
            continue
        for item in track.get("items") or []:
            if not isinstance(item, dict):
                continue
            yield item, "primary"
            for sec in item.get("secondary_cuts") or []:
                if isinstance(sec, dict):
                    yield sec, "secondary"


def _track_label(items: list[dict[str, Any]]) -> str:
    loads = {
        _load_code(item.get("load_code"))
        for item in items
        if item.get("load_code") is not None
    }
    if not loads:
        return "Нагрузка"
    rendered = [format_reinforcement_from_load_code(lc) for lc in sorted(loads)]
    return "Нагрузка " + ", ".join(rendered)


def _assignment_to_track_item(assignment: dict[str, Any]) -> dict[str, Any]:
    length = round(float(assignment.get("length") or 0), 3)
    width_mm = _width_mm(assignment.get("width"))
    load_code = _load_code(assignment.get("load_code"))
    plate_name = assignment.get("plate_name") or _plate_name.make(
        length,
        width_mm,
        load_code,
        length_dm_raw=assignment.get("length_dm_raw") or None,
    )
    return {
        "length": length,
        "mode": "solid",
        "width": width_mm / 1000.0,
        "load_code": load_code,
        "label": plate_name,
        "reinforcement": 0,
        "kp_id": assignment.get("kp_id"),
        "customer": assignment.get("customer"),
        "kp_date": assignment.get("kp_date"),
        "plate_name": plate_name,
        "plate_uid": assignment.get("plate_uid"),
        "reconciled_from_assignment": True,
        "concrete_grade": assignment.get("concrete_grade"),
    }


def _pack_assignments_into_tracks(
    assignments: list[dict[str, Any]],
    *,
    max_track_length_m: float = MAX_TRACK_LENGTH_M,
) -> list[dict[str, Any]]:
    tracks: list[dict[str, Any]] = []
    current_items: list[dict[str, Any]] = []
    current_length = 0.0

    def flush() -> None:
        nonlocal current_items, current_length
        if not current_items:
            return
        tracks.append({
            "items": current_items,
            "length": current_length,
            "load_code": current_items[0].get("load_code", 0),
            "label": _track_label(current_items),
            "max_reinforcement": max(
                (float(item.get("reinforcement") or 0) for item in current_items),
                default=0.0,
            ),
            "reconciled_extra_track": True,
        })
        current_items = []
        current_length = 0.0

    for assignment in assignments:
        item = _assignment_to_track_item(assignment)
        item_length = float(item.get("length") or 0)
        if current_items and current_length + item_length > max_track_length_m:
            flush()
        current_items.append(item)
        current_length += item_length

    flush()
    return tracks


def _missing_by_uid(
    plate_assignments: list[dict[str, Any]],
    tracks_list: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[str]]:
    assignment_by_uid = {
        str(assignment["plate_uid"]): assignment
        for assignment in plate_assignments
        if assignment.get("plate_uid")
    }
    if not assignment_by_uid:
        return [], []

    track_uids = {
        str(item["plate_uid"])
        for item, _ in _iter_physical_track_items(tracks_list)
        if item.get("plate_uid")
    }
    missing_uids = [uid for uid in assignment_by_uid if uid not in track_uids]
    return [assignment_by_uid[uid] for uid in missing_uids], missing_uids


def _missing_by_attributes(
    plate_assignments: list[dict[str, Any]],
    tracks_list: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict[tuple[Any, ...], int]]:
    assignment_counts: Counter[tuple[Any, ...]] = Counter()
    assignments_by_key: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
    for assignment in plate_assignments:
        source = str(assignment.get("source") or "")
        if source not in ("primary", "secondary"):
            continue
        key = _assignment_key(assignment)
        if key is None:
            continue
        assignment_counts[key] += 1
        assignments_by_key.setdefault(key, []).append(assignment)

    track_counts: Counter[tuple[Any, ...]] = Counter()
    for item, source_hint in _iter_physical_track_items(tracks_list):
        key = _track_item_key(item, source_hint=source_hint)
        if key is not None:
            track_counts[key] += 1

    missing_assignments: list[dict[str, Any]] = []
    missing_counts: dict[tuple[Any, ...], int] = {}
    for key, expected in assignment_counts.items():
        deficit = expected - track_counts.get(key, 0)
        if deficit <= 0:
            continue
        missing_counts[key] = deficit
        missing_assignments.extend(assignments_by_key[key][:deficit])
    return missing_assignments, missing_counts


def reconcile_tracks_with_assignments(
    *,
    orders_2d: list[dict[str, Any]] | None,
    plate_assignments: list[dict[str, Any]],
    tracks_list: list[dict[str, Any]],
    strict: bool = False,
) -> ReconcileResult:
    """Return ordinary extra tracks for assignments missing from track items."""
    del orders_2d  # kept for API symmetry and future order-aware packing.

    missing_assignments, missing_uids = _missing_by_uid(plate_assignments, tracks_list)
    missing_counts: dict[tuple[Any, ...], int] = {}
    if not missing_assignments and not missing_uids:
        missing_assignments, missing_counts = _missing_by_attributes(
            plate_assignments,
            tracks_list,
        )

    stats = ReconcileStats(
        assignments_count=len(plate_assignments or []),
        track_items_count=sum(1 for _ in _iter_physical_track_items(tracks_list)),
        deficit_count=len(missing_assignments),
        missing_uids=missing_uids,
        missing_counts=missing_counts,
    )

    if strict and stats.deficit_count:
        raise TrackReconciliationError(
            f"Track reconciliation deficit: {stats.deficit_count} physical plates"
        )

    extra_tracks = _pack_assignments_into_tracks(missing_assignments)
    stats.extra_track_count = len(extra_tracks)
    if stats.deficit_count:
        logger.warning(
            "[RECONCILE] Добавлено %s ordinary extra tracks для %s потерянных плит",
            stats.extra_track_count,
            stats.deficit_count,
        )
    return ReconcileResult(extra_tracks=extra_tracks, stats=stats)


__all__ = [
    "ReconcileResult",
    "ReconcileStats",
    "TrackReconciliationError",
    "reconcile_tracks_with_assignments",
]
