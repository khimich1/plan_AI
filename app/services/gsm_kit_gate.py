"""Eligibility of a vehicle for the GSM month-close kit / generate."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any, Literal

from app.repositories.gsm_repository import GsmRepository
from app.services.gsm_overview_service import _chain_broken

CODE_TAIL = "gsm_kit_tail"
CODE_CHAIN = "gsm_kit_chain"
CODE_RED = "gsm_kit_red"

KitPurpose = Literal["kit", "generate"]

_MONTHS_RU = (
    "",
    "Январь",
    "Февраль",
    "Март",
    "Апрель",
    "Май",
    "Июнь",
    "Июль",
    "Август",
    "Сентябрь",
    "Октябрь",
    "Ноябрь",
    "Декабрь",
)


@dataclass(frozen=True)
class KitEligibility:
    vehicle_id: int
    allowed: bool
    code: str | None
    message: str | None


def _period_ym(period_from: date) -> str:
    return f"{period_from.year:04d}-{period_from.month:02d}"


def _month_label(yyyy_mm: str) -> str:
    try:
        month = int(yyyy_mm.split("-")[1])
    except (IndexError, ValueError):
        return yyyy_mm
    if 1 <= month <= 12:
        return _MONTHS_RU[month]
    return yyyy_mm


def evaluate_kit_vehicle(
    *,
    vehicle_id: int,
    period_from: date,
    open_before: int,
    open_before_month: str | None,
    chain_broken: bool,
    red_days: int,
    purpose: KitPurpose = "kit",
) -> KitEligibility:
    """Decide kit/generate eligibility from the same fields as fleet overview."""
    current_ym = _period_ym(period_from)
    tail_blocks = int(open_before or 0) > 0 and current_ym != open_before_month
    tail_ym = open_before_month or current_ym

    if purpose == "generate":
        if tail_blocks:
            return KitEligibility(
                vehicle_id=vehicle_id,
                allowed=False,
                code=CODE_TAIL,
                message=f"сначала выгрузите {_month_label(tail_ym)}",
            )
        return KitEligibility(vehicle_id=vehicle_id, allowed=True, code=None, message=None)

    if int(red_days or 0) > 0:
        return KitEligibility(
            vehicle_id=vehicle_id,
            allowed=False,
            code=CODE_RED,
            message=f"Исправьте дни ручной доработки ({int(red_days)}).",
        )
    if tail_blocks:
        return KitEligibility(
            vehicle_id=vehicle_id,
            allowed=False,
            code=CODE_TAIL,
            message=f"Сначала выгрузите {_month_label(tail_ym)}",
        )
    if chain_broken:
        return KitEligibility(
            vehicle_id=vehicle_id,
            allowed=False,
            code=CODE_CHAIN,
            message=(
                f"Пересчитайте {_month_label(current_ym)}: бак не сходится с предыдущим"
            ),
        )
    return KitEligibility(vehicle_id=vehicle_id, allowed=True, code=None, message=None)


def evaluate_from_overview_row(
    row: dict[str, Any],
    *,
    period_from: date,
    purpose: KitPurpose = "kit",
) -> KitEligibility:
    open_before = int(row.get("open_before") or 0)
    return evaluate_kit_vehicle(
        vehicle_id=int(row["vehicle_id"]),
        period_from=period_from,
        open_before=open_before,
        open_before_month=row.get("open_before_month") if open_before else None,
        chain_broken=_chain_broken(row),
        red_days=int(row.get("red_days") or 0),
        purpose=purpose,
    )


def overview_rows_by_vehicle(
    repo: GsmRepository,
    *,
    period_from: date,
    period_to: date,
) -> dict[int, dict[str, Any]]:
    return {
        int(row["vehicle_id"]): row
        for row in repo.fleet_overview(period_from=period_from, period_to=period_to)
    }


def evaluate_vehicle(
    repo: GsmRepository,
    vehicle_id: int,
    *,
    period_from: date,
    period_to: date,
    purpose: KitPurpose = "kit",
    overview_by_id: dict[int, dict[str, Any]] | None = None,
) -> KitEligibility:
    rows = overview_by_id if overview_by_id is not None else overview_rows_by_vehicle(
        repo, period_from=period_from, period_to=period_to
    )
    row = rows.get(int(vehicle_id))
    if row is None:
        return evaluate_kit_vehicle(
            vehicle_id=int(vehicle_id),
            period_from=period_from,
            open_before=0,
            open_before_month=None,
            chain_broken=False,
            red_days=0,
            purpose=purpose,
        )
    return evaluate_from_overview_row(row, period_from=period_from, purpose=purpose)


def filter_kit_vehicle_ids(
    repo: GsmRepository,
    vehicle_ids: list[int],
    *,
    period_from: date,
    period_to: date,
    purpose: KitPurpose = "kit",
) -> tuple[list[int], list[KitEligibility]]:
    """Split requested ids into allowed vs blocked (blocked keep eligibility)."""
    overview_by_id = overview_rows_by_vehicle(
        repo, period_from=period_from, period_to=period_to
    )
    allowed: list[int] = []
    blocked: list[KitEligibility] = []
    for vid in vehicle_ids:
        elig = evaluate_vehicle(
            repo,
            vid,
            period_from=period_from,
            period_to=period_to,
            purpose=purpose,
            overview_by_id=overview_by_id,
        )
        if elig.allowed:
            allowed.append(vid)
        else:
            blocked.append(elig)
    return allowed, blocked
