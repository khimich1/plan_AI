"""Дозаполнение хвоста дорожек переносом плит с последующих дорожек.

После ``split_sequence_into_tracks`` каждая дорожка стремится заполниться до
``max_length_m`` (101 м), забирая целые плиты с дорожек с большим индексом.
Единственная плита последней дорожки тоже может уехать. Если весь остаток
последней дорожки влезает в зазор, он уезжает по порядку, включая первую плиту
и продольный рез. Плиты не копируются — только перенос в хвост получателя.
Пустой донор удаляется.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Any

from core.config.constants import TRACK_LENGTH_M
from core.domain.plate_order import normalize_load_code
from core.track_reconciliation import _track_label, _width_mm

logger = logging.getLogger(__name__)

DEFAULT_EPS_M = 0.01
_INVALID_REINFORCEMENT = {0.0, 999.0}
_GRADE_RANK = {"М400": 400, "М500": 500}


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


def _canonical_grade(raw: Any) -> str | None:
    if raw is None:
        return None
    text = str(raw).strip().upper().replace("M", "М")
    if text in _GRADE_RANK:
        return text
    return None


def plate_grade(item: dict[str, Any]) -> str:
    """Марка плиты: явные М400/М500, иначе нагрузка 6/8 и длина до 6,6 м."""
    explicit = _canonical_grade(item.get("concrete_grade"))
    if explicit is not None:
        return explicit
    length_dm = round(_item_length_m(item) * 10)
    load = normalize_load_code(item.get("load_code", 8))
    if load in (6, 8) and length_dm <= 66:
        return "М400"
    return "М500"


def _plates_of(item: dict[str, Any]) -> list[dict[str, Any]]:
    plates = [item]
    cuts = item.get("secondary_cuts") or []
    if isinstance(cuts, list):
        plates.extend(cut for cut in cuts if isinstance(cut, dict))
    return plates


def _receiver_grade(track: dict[str, Any]) -> str:
    ranks: list[int] = []
    for item in track.get("items") or []:
        if not isinstance(item, dict):
            continue
        ranks.extend(_GRADE_RANK[plate_grade(plate)] for plate in _plates_of(item))
    if not ranks or max(ranks) < _GRADE_RANK["М500"]:
        return "М400"
    return "М500"


def _item_uid(item: dict[str, Any]) -> str:
    uid = item.get("plate_uid")
    if uid:
        return str(uid)
    kp_id = item.get("kp_id")
    plate_name = item.get("plate_name") or item.get("label")
    length = round(_item_length_m(item), 3)
    width_mm = _item_width_mm(item)
    load_code = normalize_load_code(item.get("load_code", 8))
    if kp_id is not None and plate_name:
        return f"{kp_id}|{plate_name}|{length}|{width_mm}|{load_code}"
    return f"obj:{id(item)}"


def is_eligible_donor(
    item: dict[str, Any],
    item_index: int,
    *,
    gap: float,
    max_reinforcement: float,
    receiver_grade: str,
    sole_on_last_track: bool = False,
    last_track_remainder: bool = False,
) -> bool:
    """True, если плиту можно перенести в хвост текущей дорожки.

    ``sole_on_last_track`` разрешает индекс 0 только у единственной плиты
    последней дорожки. ``last_track_remainder`` разрешает индекс 0 и
    ``mode split`` без пары, когда в зазор влезает весь остаток последней
    дорожки. Марка и проволока проверяются у корня и у каждого элемента
    ``secondary_cuts``. Длина в зазор — ``length`` корня.
    """
    if item_index <= 0 and not sole_on_last_track and not last_track_remainder:
        return False
    if item.get("is_separator") or item.get("mode") == "transverse":
        return False
    if item.get("mode") != "solid" and not item.get("secondary_cuts"):
        if not (last_track_remainder and item.get("mode") == "split"):
            return False

    length = _item_length_m(item)
    if length <= 0 or length > gap + DEFAULT_EPS_M:
        return False

    receiver = _canonical_grade(receiver_grade)
    if receiver is None:
        return False
    receiver_rank = _GRADE_RANK[receiver]
    for plate in _plates_of(item):
        reinf = _item_reinforcement(plate)
        if reinf is None:
            return False
        if max_reinforcement > 0 and reinf > max_reinforcement + DEFAULT_EPS_M:
            return False
        if _GRADE_RANK[plate_grade(plate)] > receiver_rank:
            return False
    return True


def _donor_sort_key(item: dict[str, Any]) -> tuple[float, float, float]:
    length = _item_length_m(item)
    width_mm = float(_item_width_mm(item))
    reinf = _item_reinforcement(item) or 0.0
    return (-length, -width_mm, -reinf)


def _remainder_sort_key(items: list[dict[str, Any]]) -> tuple[float, float, float]:
    length = sum(_item_length_m(item) for item in items)
    width_mm = max(float(_item_width_mm(item)) for item in items)
    reinf = max((_item_reinforcement(item) or 0.0) for item in items)
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
            normalize_load_code(item.get("load_code", 8))
            for item in items
            if isinstance(item, dict) and item.get("load_code") is not None
        }
        if len(load_codes) == 1:
            track["load_code"] = next(iter(load_codes))


def _discard_empty_donors(tracks: list[dict[str, Any]], receiver_index: int) -> None:
    """Удаляет пустые дорожки только после получателя, чтобы его индекс не съехал."""
    donor_index = receiver_index + 1
    while donor_index < len(tracks):
        donor = tracks[donor_index]
        if isinstance(donor, dict) and not (donor.get("items") or []):
            del tracks[donor_index]
            continue
        donor_index += 1


def _fitting_last_remainder(
    tracks: list[dict[str, Any]],
    receiver_index: int,
    *,
    gap: float,
    max_reinforcement: float,
    receiver_grade: str,
    used_ids: set[str],
) -> tuple[int, list[dict[str, Any]]] | None:
    """Весь остаток последней дорожки, если сумма влезает и ни один item не старше получателя."""
    last_index = len(tracks) - 1
    if last_index <= receiver_index:
        return None
    donor = tracks[last_index]
    if not isinstance(donor, dict):
        return None
    items = donor.get("items") or []
    if not items or any(not isinstance(item, dict) for item in items):
        return None
    if any(_item_uid(item) in used_ids for item in items):
        return None
    total = sum(_item_length_m(item) for item in items)
    if total <= 0 or total > gap + DEFAULT_EPS_M:
        return None
    for item_index, item in enumerate(items):
        if not is_eligible_donor(
            item,
            item_index,
            gap=gap,
            max_reinforcement=max_reinforcement,
            receiver_grade=receiver_grade,
            last_track_remainder=True,
        ):
            return None
    return last_index, items


def _best_donor(
    tracks: list[dict[str, Any]],
    receiver_index: int,
    *,
    gap: float,
    max_reinforcement: float,
    receiver_grade: str,
    used_ids: set[str],
) -> tuple[tuple[float, float, float], int, int | None] | None:
    """Лучший донор. ``item_index is None`` — перенести весь остаток последней дорожки."""
    scored: list[tuple[tuple[float, float, float], int, int | None]] = []
    last_index = len(tracks) - 1
    for donor_index in range(receiver_index + 1, len(tracks)):
        donor = tracks[donor_index]
        if not isinstance(donor, dict):
            continue
        donor_items = donor.get("items") or []
        sole_on_last_track = donor_index == last_index and len(donor_items) == 1
        for item_index, item in enumerate(donor_items):
            if not isinstance(item, dict):
                continue
            uid = _item_uid(item)
            if uid in used_ids:
                continue
            if not is_eligible_donor(
                item,
                item_index,
                gap=gap,
                max_reinforcement=max_reinforcement,
                receiver_grade=receiver_grade,
                sole_on_last_track=sole_on_last_track,
            ):
                continue
            scored.append((_donor_sort_key(item), donor_index, item_index))
    remainder = _fitting_last_remainder(
        tracks,
        receiver_index,
        gap=gap,
        max_reinforcement=max_reinforcement,
        receiver_grade=receiver_grade,
        used_ids=used_ids,
    )
    if remainder is not None:
        donor_index, items = remainder
        scored.append((_remainder_sort_key(items), donor_index, None))
    if not scored:
        return None
    scored.sort(key=lambda row: row[0])
    return scored[0]


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
    receiver_index = 0
    while receiver_index < len(tracks):
        track = tracks[receiver_index]
        if not isinstance(track, dict):
            receiver_index += 1
            continue
        recalc_track(track)
        gap = max_length_m - float(track.get("length") or 0)
        if gap <= eps:
            receiver_index += 1
            continue

        while gap > eps:
            _discard_empty_donors(tracks, receiver_index)
            chosen = _best_donor(
                tracks,
                receiver_index,
                gap=gap,
                max_reinforcement=_track_max_reinforcement(track),
                receiver_grade=_receiver_grade(track),
                used_ids=used_ids,
            )
            if chosen is None:
                break

            _key, donor_idx, item_idx = chosen
            donor_track = tracks[donor_idx]
            donor_items = donor_track.setdefault("items", [])
            if item_idx is None:
                moved_items = list(donor_items)
                donor_items.clear()
            else:
                moved_items = [donor_items.pop(item_idx)]

            for moved in moved_items:
                moved["top_up"] = True
                moved["top_up_from_track_idx"] = donor_idx
                track.setdefault("items", []).append(moved)
                uid = _item_uid(moved)
                used_ids.add(uid)
                added = _item_length_m(moved)
                result.moves += 1
                result.meters_added += added
                touched.add(receiver_index)
                result.move_log.append(
                    {
                        "to_track": receiver_index,
                        "from_track": donor_idx,
                        "length": added,
                        "plate_uid": uid,
                    }
                )

            recalc_track(donor_track)
            recalc_track(track)
            if not donor_items:
                del tracks[donor_idx]
            gap = max_length_m - float(track.get("length") or 0)

        receiver_index += 1

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
    "plate_grade",
    "recalc_track",
    "top_up_tracks_from_following",
]
