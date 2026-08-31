"""Pure substrate-recommendation helpers (no I/O, no app imports).

Extracts cross-KP primary→secondary matches from an optimizer result and
aggregates them into ``SubstrateRecommendation`` rows.

Aggregation key: ``(plate_id, under_plate_id)`` where ``plate_*`` is the late
substrate (secondary cut) and ``under_*`` is the urgent primary plate.

Saving choice: among cuts sharing an aggregation key, keep the cut with the
**maximum** ``saving_mm`` (parent ``rest``); ``saving_m`` is taken from that
same cut (``rest_mm * length_m / 1000``). ``qty_recommended`` is the sum of
secondary cut quantities.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any, Mapping, Sequence


@dataclass(frozen=True, slots=True)
class SubstrateRecommendation:
    plate_id: int
    kp_id: int
    plate_name: str
    qty_recommended: int
    under_plate_id: int
    under_kp_id: int
    under_plate_name: str
    needed_by: date
    storage_days: int
    saving_mm: int
    saving_m: float


@dataclass(frozen=True, slots=True)
class _AggBucket:
    plate_id: int
    kp_id: int
    plate_name: str
    under_plate_id: int
    under_kp_id: int
    under_plate_name: str
    needed_by: date
    storage_days: int
    qty_recommended: int
    saving_mm: int
    saving_m: float


def _primary_length_m(primary: Mapping[str, Any], secondary: Mapping[str, Any]) -> float:
    lengths = primary.get("lengths") or ()
    if lengths:
        try:
            return float(lengths[0])
        except (TypeError, ValueError):
            pass
    sec_lengths = secondary.get("lengths") or secondary.get("source_lengths") or ()
    if sec_lengths:
        try:
            return float(sec_lengths[0])
        except (TypeError, ValueError):
            pass
    return 0.0


def _saving_m(rest_mm: int, length_m: float) -> float:
    return float(rest_mm) * float(length_m) / 1000.0


def _resolve_plate_id(
    plate_id_by_kp_name: Mapping[tuple[int, str], int],
    kp_id: Any,
    plate_name: Any,
) -> int | None:
    if kp_id is None:
        return None
    try:
        kid = int(kp_id)
    except (TypeError, ValueError):
        return None
    name = str(plate_name or "")
    return plate_id_by_kp_name.get((kid, name))


def extract_substrate_recommendations(
    result: Mapping[str, Any],
    *,
    plate_id_by_kp_name: Mapping[tuple[int, str], int],
    urgent_plate_ids: Sequence[int] | set[int],
    deadline_by_plate_id: Mapping[int, date],
    first_fill_target_date: date,
) -> list[SubstrateRecommendation]:
    """Extract aggregated cross-KP substrate recommendations from opt output.

    Rules:
    - Index ``primary_cuts`` by ``primary_instance_id``.
    - For each secondary with ``parent_instance_id``: skip missing parent,
      same-KP pairs, unresolved plate ids, primary not in ``urgent_plate_ids``,
      or substrate without a deadline.
    - Aggregate by ``(plate_id, under_plate_id)``: sum qty; keep max saving_mm
      (and that cut's saving_m).
    - Sort by ``saving_m`` descending.
    """
    urgent = {int(x) for x in urgent_plate_ids}

    primary_by_id: dict[str, Mapping[str, Any]] = {}
    for cut in result.get("primary_cuts") or ():
        unit_id = cut.get("primary_instance_id")
        if unit_id:
            primary_by_id[str(unit_id)] = cut

    buckets: dict[tuple[int, int], _AggBucket] = {}

    for sec in result.get("secondary_cuts") or ():
        parent_id = sec.get("parent_instance_id")
        if not parent_id:
            continue
        primary = primary_by_id.get(str(parent_id))
        if primary is None:
            continue

        sec_kp = sec.get("kp_id")
        pri_kp = primary.get("kp_id")
        if sec_kp is None or pri_kp is None:
            continue
        try:
            sec_kp_i = int(sec_kp)
            pri_kp_i = int(pri_kp)
        except (TypeError, ValueError):
            continue
        if sec_kp_i == pri_kp_i:
            continue

        under_plate_id = _resolve_plate_id(
            plate_id_by_kp_name, pri_kp_i, primary.get("plate_name")
        )
        if under_plate_id is None or under_plate_id not in urgent:
            continue

        plate_id = _resolve_plate_id(
            plate_id_by_kp_name, sec_kp_i, sec.get("plate_name")
        )
        if plate_id is None:
            continue

        needed_by = deadline_by_plate_id.get(plate_id)
        if needed_by is None:
            continue

        rest_mm = int(primary.get("rest") or 0)
        length_m = _primary_length_m(primary, sec)
        saving_m = _saving_m(rest_mm, length_m)
        qty = int(sec.get("qty") or 1)
        if qty <= 0:
            qty = 1

        under_name = str(primary.get("plate_name") or "")
        plate_name = str(sec.get("plate_name") or "")
        storage_days = (needed_by - first_fill_target_date).days
        key = (plate_id, under_plate_id)
        existing = buckets.get(key)
        if existing is None:
            buckets[key] = _AggBucket(
                plate_id=plate_id,
                kp_id=sec_kp_i,
                plate_name=plate_name,
                under_plate_id=under_plate_id,
                under_kp_id=pri_kp_i,
                under_plate_name=under_name,
                needed_by=needed_by,
                storage_days=storage_days,
                qty_recommended=qty,
                saving_mm=rest_mm,
                saving_m=saving_m,
            )
            continue

        new_qty = existing.qty_recommended + qty
        if rest_mm > existing.saving_mm:
            buckets[key] = _AggBucket(
                plate_id=existing.plate_id,
                kp_id=existing.kp_id,
                plate_name=existing.plate_name,
                under_plate_id=existing.under_plate_id,
                under_kp_id=existing.under_kp_id,
                under_plate_name=existing.under_plate_name,
                needed_by=existing.needed_by,
                storage_days=existing.storage_days,
                qty_recommended=new_qty,
                saving_mm=rest_mm,
                saving_m=saving_m,
            )
        else:
            buckets[key] = _AggBucket(
                plate_id=existing.plate_id,
                kp_id=existing.kp_id,
                plate_name=existing.plate_name,
                under_plate_id=existing.under_plate_id,
                under_kp_id=existing.under_kp_id,
                under_plate_name=existing.under_plate_name,
                needed_by=existing.needed_by,
                storage_days=existing.storage_days,
                qty_recommended=new_qty,
                saving_mm=existing.saving_mm,
                saving_m=existing.saving_m,
            )

    recommendations = [
        SubstrateRecommendation(
            plate_id=b.plate_id,
            kp_id=b.kp_id,
            plate_name=b.plate_name,
            qty_recommended=b.qty_recommended,
            under_plate_id=b.under_plate_id,
            under_kp_id=b.under_kp_id,
            under_plate_name=b.under_plate_name,
            needed_by=b.needed_by,
            storage_days=b.storage_days,
            saving_mm=b.saving_mm,
            saving_m=b.saving_m,
        )
        for b in buckets.values()
    ]
    recommendations.sort(key=lambda r: (-r.saving_m, r.kp_id, r.plate_id))
    return recommendations
