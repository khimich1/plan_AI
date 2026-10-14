"""Pure driver-day pool: eligible codes, rotation rank, pick (no DB)."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass

ELIGIBLE_CODES = frozenset({"Я", "В"})


@dataclass(frozen=True, slots=True)
class DriverCandidate:
    """One driver who may sign a waybill on a date.

    ``personnel_number`` and ``has_own_vehicle`` feed the rotation rank.
    The generator re-ranks each day because a substitution increments the
    counter for later days of the same vehicle.
    """

    driver_id: int
    code: str
    personnel_number: str = ""
    has_own_vehicle: bool = False


def order_candidates(
    primary_id: int,
    candidates: Sequence[DriverCandidate],
    counts: Mapping[int, int],
) -> tuple[DriverCandidate, ...]:
    """Primary, then drivers without a vehicle, then other primaries.

    Inside a group the rank is ``(substitute_count, personnel_number)``.
    Codes other than Я/В are dropped.
    """
    eligible = [item for item in candidates if item.code in ELIGIBLE_CODES]
    primary = [item for item in eligible if item.driver_id == primary_id]
    rest = [item for item in eligible if item.driver_id != primary_id]
    no_vehicle = [item for item in rest if not item.has_own_vehicle]
    others = [item for item in rest if item.has_own_vehicle]
    no_vehicle.sort(key=lambda item: _rank(item, counts))
    others.sort(key=lambda item: _rank(item, counts))

    ordered: list[DriverCandidate] = []
    seen: set[int] = set()
    for item in (*primary, *no_vehicle, *others):
        if item.driver_id in seen:
            continue
        seen.add(item.driver_id)
        ordered.append(item)
    return tuple(ordered)


def pick_driver(
    day_candidates: Sequence[DriverCandidate],
) -> tuple[int | None, tuple[str, ...]]:
    """Head of an already ordered day. Empty → ``no_available_driver``.

    Code ``В`` on the chosen driver adds soft ``weekend_driver``.
    """
    if not day_candidates:
        return None, ("no_available_driver",)
    head = day_candidates[0]
    if head.code == "В":
        return head.driver_id, ("weekend_driver",)
    return head.driver_id, ()


def _rank(item: DriverCandidate, counts: Mapping[int, int]) -> tuple[int, int, int]:
    number = item.personnel_number.strip()
    personnel = int(number) if number.isdigit() else 10**9
    return (int(counts.get(item.driver_id, 0)), personnel, item.driver_id)
