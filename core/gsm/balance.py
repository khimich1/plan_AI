"""Pure tank-balance math: burn, day apply, day chain, corridor checks."""

from __future__ import annotations

from collections.abc import Sequence
from datetime import date

from core.gsm.models import TankState


class BalanceViolation(RuntimeError):
    """Tank corridor ``0 ≤ fuel_end ≤ tank_volume`` was breached."""


def burn_for_km(km: int, norm_per_100km: float) -> float:
    """Daily fuel burn by norm: ``round(km * norm / 100, 2)``."""
    return round(km * norm_per_100km / 100, 2)


def apply_day(
    day: date,
    *,
    fuel_start: float,
    fuel_issued: float,
    km: int,
    odometer_start: int,
    norm_per_100km: float,
    tank_volume_liters: float,
) -> TankState:
    """Apply one day's issue + burn; raise if fuel_end leaves ``[0, tank]``."""
    burn = burn_for_km(km, norm_per_100km)
    fuel_end = round(fuel_start + fuel_issued - burn, 2)
    if fuel_end < 0 or fuel_end > tank_volume_liters:
        raise BalanceViolation(
            f"fuel_end={fuel_end} outside [0, {tank_volume_liters}] on {day.isoformat()}"
        )
    return TankState(
        date=day,
        fuel_start=fuel_start,
        fuel_issued=fuel_issued,
        fuel_end=fuel_end,
        km=km,
        odometer_start=odometer_start,
        odometer_end=odometer_start + km,
    )


def apply_day_chain(
    days: Sequence[tuple[date, float, int]],
    *,
    fuel_start: float,
    odometer_start: int,
    norm_per_100km: float,
    tank_volume_liters: float,
) -> tuple[TankState, ...]:
    """Apply a sequence of ``(date, fuel_issued, km)`` carrying fuel/odometer."""
    states: list[TankState] = []
    fuel = fuel_start
    odometer = odometer_start
    for day, fuel_issued, km in days:
        state = apply_day(
            day,
            fuel_start=fuel,
            fuel_issued=fuel_issued,
            km=km,
            odometer_start=odometer,
            norm_per_100km=norm_per_100km,
            tank_volume_liters=tank_volume_liters,
        )
        states.append(state)
        fuel = state.fuel_end
        odometer = state.odometer_end
    return tuple(states)
