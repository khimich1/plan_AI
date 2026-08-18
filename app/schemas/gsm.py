"""Pydantic contracts for the GSM (fuel / waybill) API."""

from __future__ import annotations

import re
from datetime import date

from pydantic import BaseModel, ConfigDict, Field, field_validator

_WINTER_START_RE = re.compile(r"^(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$")


# ---------------------------------------------------------------------------
# Transactions import
# ---------------------------------------------------------------------------


class FileImportReport(BaseModel):
    """Per-file result of a transaction import."""

    model_config = ConfigDict(extra="forbid")

    filename: str
    rows_total: int
    rows_inserted: int
    rows_duplicate: int
    sum_liters: float
    sum_amount: float
    footer_liters: float | None = None
    footer_amount: float | None = None
    warnings: list[str] = Field(default_factory=list)
    unmatched_cards: list[str] = Field(default_factory=list)


class TransactionImportReport(BaseModel):
    """Aggregate report for a multi-file transaction import."""

    model_config = ConfigDict(extra="forbid")

    files: list[FileImportReport]
    rows_inserted: int
    rows_duplicate: int


# ---------------------------------------------------------------------------
# Registry: vehicles
# ---------------------------------------------------------------------------


class VehicleCreateRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str = Field(min_length=1)
    plate_number: str = Field(min_length=1)
    tank_volume_liters: float = Field(gt=0)
    norm_summer: float = Field(gt=0)
    norm_winter: float = Field(gt=0)
    primary_driver_id: int | None = None
    is_active: bool = True


class VehiclePatchRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str | None = Field(default=None, min_length=1)
    plate_number: str | None = Field(default=None, min_length=1)
    tank_volume_liters: float | None = Field(default=None, gt=0)
    norm_summer: float | None = Field(default=None, gt=0)
    norm_winter: float | None = Field(default=None, gt=0)
    primary_driver_id: int | None = None
    is_active: bool | None = None


class VehicleOut(BaseModel):
    model_config = ConfigDict(extra="forbid")

    id: int
    name: str
    plate_number: str
    tank_volume_liters: float
    norm_summer: float
    norm_winter: float
    primary_driver_id: int | None = None
    is_active: bool = True


# ---------------------------------------------------------------------------
# Registry: drivers
# ---------------------------------------------------------------------------


class DriverCreateRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    full_name: str = Field(min_length=1)
    license_number: str = Field(min_length=1)
    license_issued_at: str | None = None
    personnel_number: str | None = None
    snils: str | None = None
    is_active: bool = True


class DriverPatchRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    full_name: str | None = Field(default=None, min_length=1)
    license_number: str | None = Field(default=None, min_length=1)
    license_issued_at: str | None = None
    personnel_number: str | None = None
    snils: str | None = None
    is_active: bool | None = None


class DriverOut(BaseModel):
    model_config = ConfigDict(extra="forbid")

    id: int
    full_name: str
    license_number: str
    license_issued_at: str | None = None
    personnel_number: str | None = None
    snils: str | None = None
    is_active: bool = True


# ---------------------------------------------------------------------------
# Registry: fuel cards
# ---------------------------------------------------------------------------


class CardCreateRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    card_number: str = Field(min_length=1)
    vehicle_id: int | None = None
    assigned_at: str | None = None


class CardPatchRequest(BaseModel):
    """Bind to vehicle and/or archive (sets archived_at; never DELETE)."""

    model_config = ConfigDict(extra="forbid")

    vehicle_id: int | None = None
    archive: bool | None = None


class CardOut(BaseModel):
    model_config = ConfigDict(extra="forbid")

    id: int
    card_number: str
    vehicle_id: int | None = None
    assigned_at: str
    archived_at: str | None = None


# ---------------------------------------------------------------------------
# Registry: stations
# ---------------------------------------------------------------------------


class StationCreateRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    address: str = Field(min_length=1)
    brand: str | None = None
    lat: float | None = None
    lon: float | None = None
    geocode_source: str | None = None


class StationPatchRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    address: str | None = Field(default=None, min_length=1)
    brand: str | None = None
    lat: float | None = None
    lon: float | None = None
    geocode_source: str | None = None


class StationOut(BaseModel):
    model_config = ConfigDict(extra="forbid")

    id: int
    address: str
    brand: str | None = None
    lat: float | None = None
    lon: float | None = None
    geocode_source: str | None = None


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------


class GsmSettings(BaseModel):
    model_config = ConfigDict(extra="forbid")

    winter_start: str = "11-01"
    hook_threshold_km: float = Field(default=13.0, gt=0)
    max_daily_km: int = Field(default=700, gt=0)

    @field_validator("winter_start")
    @classmethod
    def _validate_winter_start(cls, value: str) -> str:
        if not _WINTER_START_RE.match(value):
            raise ValueError("winter_start must be MM-DD")
        return value


# ---------------------------------------------------------------------------
# Registry: routes (vehicle library)
# ---------------------------------------------------------------------------


class RouteOut(BaseModel):
    model_config = ConfigDict(extra="forbid")

    id: int
    vehicle_id: int
    addr_a: str
    addr_b: str
    km: int
    frequency: int = 1
    typical_station_ids: list[int] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# Waybills / generation
# ---------------------------------------------------------------------------


class WaybillGenerateRequest(BaseModel):
    """POST /gsm/waybills/generate — explicit period (product may default to tx min/max in UI)."""

    model_config = ConfigDict(extra="forbid")

    vehicle_id: int
    period_from: date
    period_to: date
    force: bool = False
    fuel_start: float | None = Field(default=None, ge=0)
    odometer_start: int | None = Field(default=None, ge=0)


class WaybillRouteLeg(BaseModel):
    model_config = ConfigDict(extra="forbid", populate_by_name=True)

    from_addr: str = Field(alias="from")
    to_addr: str = Field(alias="to")
    km: int
    route_id: int | None = None
    station_id: int | None = None
    dep_time: str | None = None
    arr_time: str | None = None


class WaybillOut(BaseModel):
    model_config = ConfigDict(extra="forbid", populate_by_name=True)

    id: int
    vehicle_id: int
    date: str
    driver_id: int
    status: str
    source: str = "auto"
    odometer_start: int | None = None
    odometer_end: int | None = None
    fuel_start: float | None = None
    fuel_issued: float | None = None
    fuel_end: float | None = None
    km: int = 0
    route: list[WaybillRouteLeg] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class ProblematicDayOut(BaseModel):
    """Anchor that could not stay in the tank corridor; day is still saved."""

    model_config = ConfigDict(extra="forbid")

    date: str
    reason: str
    detail: str
    fuel_before: float
    fuel_to_issue: float
    tank_volume: float


class WaybillGenerateResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    waybills: list[WaybillOut]
    warnings: list[str] = Field(default_factory=list)
    days_created: int = 0
    problematic_days: list[ProblematicDayOut] = Field(default_factory=list)
    manual_days: int = 0


class WaybillPatchRequest(BaseModel):
    """PATCH /gsm/waybills/{id} — route/driver/km; day becomes manual, downstream drafts rechain."""

    model_config = ConfigDict(extra="forbid")

    driver_id: int | None = None
    km: int | None = Field(default=None, ge=0)
    route: list[WaybillRouteLeg] | None = None


class WaybillCreateRequest(BaseModel):
    """POST /gsm/waybills — manual constructor; fuel/odo auto-derived from previous day."""

    model_config = ConfigDict(extra="forbid")

    vehicle_id: int
    date: date
    driver_id: int
    route: list[WaybillRouteLeg] = Field(min_length=1)
    fuel_issued: float = Field(default=0.0, ge=0)
    fuel_start: float | None = Field(default=None, ge=0)
    odometer_start: int | None = Field(default=None, ge=0)


class WaybillExportRequest(BaseModel):
    """POST /gsm/waybills/export — zip of «ПЛ DD.MM.YY.xls» for selected vehicles/period."""

    model_config = ConfigDict(extra="forbid", populate_by_name=True)

    vehicle_ids: list[int] = Field(min_length=1)
    period_from: date = Field(alias="from")
    period_to: date = Field(alias="to")
