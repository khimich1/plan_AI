"""Fleet overview aggregates + per-vehicle period status (read-only)."""

from __future__ import annotations

from datetime import date
from typing import Any

from app.repositories.gsm_repository import GsmRepository
from app.schemas.gsm import (
    FleetOverviewRow,
    FleetOverviewVehicle,
    VehiclePeriodStatus,
)

_CHAIN_LITERS_EPS = 0.01


class GsmOverviewError(Exception):
    """Domain error for fleet overview (invalid period)."""

    def __init__(self, message: str, *, code: str) -> None:
        super().__init__(message)
        self.code = code
        self.details: dict[str, object] = {}


class GsmOverviewService:
    """Fleet overview aggregates + per-vehicle period status (read-only)."""

    def __init__(self, *, repo: GsmRepository) -> None:
        self._repo = repo

    def overview(self, *, period_from: date, period_to: date) -> list[FleetOverviewRow]:
        if period_to < period_from:
            raise GsmOverviewError(
                "period_to must be >= period_from",
                code="gsm_invalid_period",
            )
        rows = self._repo.fleet_overview(period_from=period_from, period_to=period_to)
        return [_to_row(row, status=_status_of(row)) for row in rows]


def _chain_broken(row: dict[str, Any]) -> bool:
    """True when last PL before period does not stitch to the first in period."""
    prev_fuel = row.get("chain_prev_fuel_end")
    first_fuel = row.get("chain_first_fuel_start")
    if prev_fuel is None or first_fuel is None:
        return False
    fuel_gap = abs(float(prev_fuel) - float(first_fuel)) > _CHAIN_LITERS_EPS
    odo_gap = row.get("chain_prev_odometer_end") != row.get("chain_first_odometer_start")
    return fuel_gap or odo_gap


def _status_of(agg: dict[str, Any]) -> VehiclePeriodStatus:
    if agg["tx_count"] == 0 and agg["wb_count"] == 0:
        return "no_data"
    if agg["red_days"] > 0:
        return "has_red_days"
    if agg["tx_count"] > 0 and (
        agg["wb_count"] == 0 or agg["tx_last_date"] > (agg["wb_last_date"] or "")
    ):
        return "needs_generation"
    if agg["draft_count"] > 0:
        return "drafts_pending"
    if agg["exported_count"] < agg["wb_count"]:
        return "pending_export"
    return "ready"


def _to_row(row: dict[str, Any], *, status: VehiclePeriodStatus) -> FleetOverviewRow:
    tx_liters = round(float(row.get("tx_liters") or 0), 2)
    wb_fuel_issued = round(float(row.get("wb_fuel_issued") or 0), 2)
    fuel_end_last = row.get("fuel_end_last")
    open_before = int(row["open_before"] or 0)
    return FleetOverviewRow(
        vehicle=FleetOverviewVehicle(
            id=int(row["vehicle_id"]),
            name=str(row["name"]),
            plate_number=str(row["plate_number"]),
        ),
        tx_count=int(row["tx_count"] or 0),
        tx_liters=tx_liters,
        tx_amount=round(float(row.get("tx_amount") or 0), 2),
        tx_last_date=row.get("tx_last_date"),
        wb_count=int(row["wb_count"] or 0),
        wb_km=int(row["wb_km"] or 0),
        wb_fuel_issued=wb_fuel_issued,
        wb_last_date=row.get("wb_last_date"),
        red_days=int(row["red_days"] or 0),
        draft_count=int(row["draft_count"] or 0),
        confirmed_count=int(row["confirmed_count"] or 0),
        exported_count=int(row["exported_count"] or 0),
        fuel_end_last=None if fuel_end_last is None else round(float(fuel_end_last), 2),
        liters_diff=round(tx_liters - wb_fuel_issued, 2),
        open_before=open_before,
        open_before_month=row.get("open_before_month") if open_before else None,
        chain_broken=_chain_broken(row),
        status=status,
    )
