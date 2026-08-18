"""Frozen DTOs for the GSM waybill domain (pure, no I/O)."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime


@dataclass(frozen=True, slots=True)
class Transaction:
    """Single fuel-card transaction (fuel, wash, or other)."""

    card_id: int
    ts: datetime
    service_type: str
    qty_liters: float | None
    amount: float
    station_id: int | None
    raw_address: str


@dataclass(frozen=True, slots=True)
class Anchor:
    """Calendar day that must have a waybill (one or more transactions)."""

    date: date
    transactions: tuple[Transaction, ...]
    station_ids: tuple[int, ...]


@dataclass(frozen=True, slots=True)
class RouteRef:
    """Library route reference used on a waybill day."""

    route_id: int
    addr_a: str
    addr_b: str
    km: int


@dataclass(frozen=True, slots=True)
class LegPlan:
    """One waybill leg; ``km`` is library (per-leg) distance."""

    route_id: int
    addr_a: str
    addr_b: str
    km: int


@dataclass(frozen=True, slots=True)
class TankState:
    """Fuel/odometer balance for a single calendar day."""

    date: date
    fuel_start: float
    fuel_issued: float
    fuel_end: float
    km: int
    odometer_start: int
    odometer_end: int


@dataclass(frozen=True, slots=True)
class WaybillDay:
    """One draft waybill day (route + tank balance + metadata)."""

    date: date
    driver_id: int
    route: RouteRef
    tank: TankState
    source: str
    warnings: tuple[str, ...]
    legs: tuple[LegPlan, ...] = ()
