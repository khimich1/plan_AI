"""CRUD for GSM registry (vehicles, drivers, cards, stations) + settings."""

from __future__ import annotations

import json
import sqlite3
from datetime import date, datetime
from typing import Any, Literal, cast

from app.repositories.gsm_repository import GsmRepository
from app.schemas.gsm import (
    CardOut,
    DriverOut,
    GsmSettings,
    RouteOut,
    StationOut,
    VehicleOut,
)
from core.gsm.season import (
    SEASON_MODES,
    SeasonSwitch,
    parse_season_switches,
    serialize_season_switches,
)

DEFAULT_WINTER_START = "11-01"
DEFAULT_HOOK_THRESHOLD_KM = 13.0
DEFAULT_MAX_DAILY_KM = 700


class GsmRegistryError(ValueError):
    """Domain error for GSM registry operations (maps to 404/422)."""

    def __init__(self, message: str, *, code: str = "gsm_registry_error") -> None:
        super().__init__(message)
        self.code = code


class GsmRegistryService:
    """Registry CRUD; cards are archived, never deleted."""

    def __init__(self, *, repo: GsmRepository) -> None:
        self._repo = repo

    # ------------------------------------------------------------------
    # Vehicles
    # ------------------------------------------------------------------

    def list_vehicles(self, *, active_only: bool = True) -> list[VehicleOut]:
        return [self._vehicle_out(row) for row in self._repo.list_vehicles(active_only=active_only)]

    def create_vehicle(
        self,
        *,
        name: str,
        plate_number: str,
        tank_volume_liters: float,
        norm_summer: float,
        norm_winter: float,
        primary_driver_id: int | None = None,
        is_active: bool = True,
    ) -> VehicleOut:
        self._validate_vehicle_numbers(
            tank_volume_liters=tank_volume_liters,
            norm_summer=norm_summer,
            norm_winter=norm_winter,
        )
        if primary_driver_id is not None:
            self._require_driver(primary_driver_id)
        vehicle_id = self._repo.create_vehicle(
            name=name,
            plate_number=plate_number,
            tank_volume_liters=tank_volume_liters,
            norm_summer=norm_summer,
            norm_winter=norm_winter,
            primary_driver_id=primary_driver_id,
            is_active=1 if is_active else 0,
        )
        return self._require_vehicle_out(vehicle_id)

    def patch_vehicle(self, vehicle_id: int, **fields: Any) -> VehicleOut:
        self._require_vehicle(vehicle_id)
        if not fields:
            return self._require_vehicle_out(vehicle_id)

        if "tank_volume_liters" in fields or "norm_summer" in fields or "norm_winter" in fields:
            current = self._repo.get_vehicle(vehicle_id) or {}
            self._validate_vehicle_numbers(
                tank_volume_liters=fields.get(
                    "tank_volume_liters", current.get("tank_volume_liters")
                ),
                norm_summer=fields.get("norm_summer", current.get("norm_summer")),
                norm_winter=fields.get("norm_winter", current.get("norm_winter")),
            )

        db_fields: dict[str, Any] = {}
        for key, value in fields.items():
            if key == "is_active" and value is not None:
                db_fields["is_active"] = 1 if value else 0
            elif key == "primary_driver_id":
                if value is not None:
                    self._require_driver(int(value))
                db_fields["primary_driver_id"] = value
            elif value is not None:
                db_fields[key] = value

        if db_fields:
            self._repo.update_vehicle(vehicle_id, **db_fields)
        return self._require_vehicle_out(vehicle_id)

    # ------------------------------------------------------------------
    # Drivers
    # ------------------------------------------------------------------

    def list_drivers(self, *, active_only: bool = True) -> list[DriverOut]:
        return [self._driver_out(row) for row in self._repo.list_drivers(active_only=active_only)]

    def create_driver(
        self,
        *,
        full_name: str,
        license_number: str,
        license_issued_at: str | None = None,
        personnel_number: str | None = None,
        snils: str | None = None,
        is_active: bool = True,
    ) -> DriverOut:
        driver_id = self._repo.create_driver(
            full_name=full_name,
            license_number=license_number,
            license_issued_at=license_issued_at,
            personnel_number=personnel_number,
            snils=snils,
            is_active=1 if is_active else 0,
        )
        return self._require_driver_out(driver_id)

    def patch_driver(self, driver_id: int, **fields: Any) -> DriverOut:
        self._require_driver(driver_id)
        if not fields:
            return self._require_driver_out(driver_id)

        db_fields: dict[str, Any] = {}
        for key, value in fields.items():
            if key == "is_active" and value is not None:
                db_fields["is_active"] = 1 if value else 0
            elif value is not None or key in {
                "license_issued_at",
                "personnel_number",
                "snils",
            }:
                db_fields[key] = value

        if db_fields:
            self._repo.update_driver(driver_id, **db_fields)
        return self._require_driver_out(driver_id)

    # ------------------------------------------------------------------
    # Cards
    # ------------------------------------------------------------------

    def list_cards(self, *, include_archived: bool = False) -> list[CardOut]:
        return [
            self._card_out(row)
            for row in self._repo.list_cards(include_archived=include_archived)
        ]

    def create_card(
        self,
        *,
        card_number: str,
        vehicle_id: int | None = None,
        assigned_at: str | None = None,
    ) -> CardOut:
        number = card_number.strip()
        if not number:
            raise GsmRegistryError("card_number is required", code="gsm_validation")
        if self._repo.get_card_by_number(number) is not None:
            raise GsmRegistryError(
                f"card_number «{number}» already exists",
                code="gsm_card_duplicate",
            )
        if vehicle_id is not None:
            self._require_vehicle(vehicle_id)
        assigned = assigned_at or date.today().isoformat()
        try:
            card_id = self._repo.create_card(
                card_number=number,
                vehicle_id=vehicle_id,
                assigned_at=assigned,
            )
        except sqlite3.IntegrityError as exc:
            raise GsmRegistryError(
                f"card_number «{number}» already exists",
                code="gsm_card_duplicate",
            ) from exc
        return self._require_card_out(card_id)

    def patch_card(self, card_id: int, **fields: Any) -> CardOut:
        """Partial update. Keys: ``vehicle_id`` (bind/unbind), ``archive`` (bool)."""
        card = self._require_card(card_id)
        db_fields: dict[str, Any] = {}

        if "vehicle_id" in fields:
            vehicle_id = fields["vehicle_id"]
            if vehicle_id is not None:
                self._require_vehicle(int(vehicle_id))
                db_fields["assigned_at"] = date.today().isoformat()
            db_fields["vehicle_id"] = vehicle_id

        if "archive" in fields and fields["archive"] is not None:
            if fields["archive"] is True:
                if card.get("archived_at") is None:
                    db_fields["archived_at"] = datetime.now().isoformat(
                        timespec="seconds"
                    )
            else:
                db_fields["archived_at"] = None

        if db_fields:
            try:
                self._repo.update_card(card_id, **db_fields)
            except sqlite3.IntegrityError as exc:
                raise GsmRegistryError(
                    "card update conflict",
                    code="gsm_card_duplicate",
                ) from exc
        return self._require_card_out(card_id)

    # ------------------------------------------------------------------
    # Stations
    # ------------------------------------------------------------------

    def list_stations(self) -> list[StationOut]:
        return [self._station_out(row) for row in self._repo.list_stations()]

    def create_station(
        self,
        *,
        address: str,
        brand: str | None = None,
        lat: float | None = None,
        lon: float | None = None,
        geocode_source: str | None = None,
    ) -> StationOut:
        try:
            station_id = self._repo.create_station(
                address=address,
                brand=brand,
                lat=lat,
                lon=lon,
                geocode_source=geocode_source,
            )
        except sqlite3.IntegrityError as exc:
            raise GsmRegistryError(
                f"station address «{address}» already exists",
                code="gsm_station_duplicate",
            ) from exc
        return self._require_station_out(station_id)

    def patch_station(self, station_id: int, **fields: Any) -> StationOut:
        self._require_station(station_id)
        if not fields:
            return self._require_station_out(station_id)
        try:
            self._repo.update_station(station_id, **fields)
        except sqlite3.IntegrityError as exc:
            raise GsmRegistryError(
                "station address already exists",
                code="gsm_station_duplicate",
            ) from exc
        return self._require_station_out(station_id)

    # ------------------------------------------------------------------
    # Routes (library)
    # ------------------------------------------------------------------

    def list_routes(self, *, vehicle_id: int | None = None) -> list[RouteOut]:
        if vehicle_id is not None:
            self._require_vehicle(vehicle_id)
        return [self._route_out(row) for row in self._repo.list_routes(vehicle_id=vehicle_id)]

    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------

    def get_settings(self) -> GsmSettings:
        winter = self._repo.get_setting("winter_start") or DEFAULT_WINTER_START
        hook_raw = self._repo.get_setting("hook_threshold_km")
        hook = float(hook_raw) if hook_raw is not None else DEFAULT_HOOK_THRESHOLD_KM
        max_raw = self._repo.get_setting("max_daily_km")
        max_daily = int(max_raw) if max_raw is not None else DEFAULT_MAX_DAILY_KM
        switches = self._read_season_switches()
        season_mode = (
            cast(Literal["summer", "winter"], switches[-1][1]) if switches else "summer"
        )
        season_switched_at = switches[-1][0] if switches else None
        return GsmSettings(
            winter_start=winter,
            hook_threshold_km=hook,
            max_daily_km=max_daily,
            season_mode=season_mode,
            season_switched_at=season_switched_at,
        )

    def put_settings(
        self,
        *,
        winter_start: str,
        hook_threshold_km: float,
        max_daily_km: int = DEFAULT_MAX_DAILY_KM,
    ) -> GsmSettings:
        if hook_threshold_km <= 0:
            raise GsmRegistryError(
                "hook_threshold_km must be > 0",
                code="gsm_validation",
            )
        if max_daily_km <= 0:
            raise GsmRegistryError(
                "max_daily_km must be > 0",
                code="gsm_validation",
            )
        settings = GsmSettings(
            winter_start=winter_start,
            hook_threshold_km=hook_threshold_km,
            max_daily_km=max_daily_km,
        )
        self._repo.set_setting("winter_start", settings.winter_start)
        self._repo.set_setting("hook_threshold_km", str(settings.hook_threshold_km))
        self._repo.set_setting("max_daily_km", str(settings.max_daily_km))
        return self.get_settings()

    def switch_season(self, *, mode: str, day: date) -> GsmSettings:
        """Append a manual season switch; same-mode call is a no-op.

        Switches are append-only and must not move backwards in time.
        """
        if mode not in SEASON_MODES:
            raise GsmRegistryError(
                f"invalid season mode: {mode!r}",
                code="gsm_validation",
            )
        switches = self._read_season_switches()
        current_mode = switches[-1][1] if switches else "summer"
        if mode == current_mode:
            return self.get_settings()
        if switches and day < switches[-1][0]:
            raise GsmRegistryError(
                "season date must not be before last switch",
                code="gsm_validation",
            )
        switches.append((day, mode))
        self._repo.set_setting("season_switches", serialize_season_switches(switches))
        return self.get_settings()

    def _read_season_switches(self) -> list[SeasonSwitch]:
        """Tolerant read of the switch journal; corrupt data → no switches."""
        try:
            return list(parse_season_switches(self._repo.get_setting("season_switches")))
        except ValueError:
            return []

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _validate_vehicle_numbers(
        *,
        tank_volume_liters: Any,
        norm_summer: Any,
        norm_winter: Any,
    ) -> None:
        try:
            tank = float(tank_volume_liters)
            summer = float(norm_summer)
            winter = float(norm_winter)
        except (TypeError, ValueError) as exc:
            raise GsmRegistryError(
                "tank and norms must be numbers > 0",
                code="gsm_validation",
            ) from exc
        if tank <= 0:
            raise GsmRegistryError("tank_volume_liters must be > 0", code="gsm_validation")
        if summer <= 0:
            raise GsmRegistryError("norm_summer must be > 0", code="gsm_validation")
        if winter <= 0:
            raise GsmRegistryError("norm_winter must be > 0", code="gsm_validation")

    def _require_vehicle(self, vehicle_id: int) -> dict[str, Any]:
        row = self._repo.get_vehicle(vehicle_id)
        if row is None:
            raise GsmRegistryError(
                f"vehicle #{vehicle_id} not found",
                code="gsm_vehicle_not_found",
            )
        return row

    def _require_driver(self, driver_id: int) -> dict[str, Any]:
        row = self._repo.get_driver(driver_id)
        if row is None:
            raise GsmRegistryError(
                f"driver #{driver_id} not found",
                code="gsm_driver_not_found",
            )
        return row

    def _require_card(self, card_id: int) -> dict[str, Any]:
        row = self._repo.get_card(card_id)
        if row is None:
            raise GsmRegistryError(
                f"card #{card_id} not found",
                code="gsm_card_not_found",
            )
        return row

    def _require_station(self, station_id: int) -> dict[str, Any]:
        row = self._repo.get_station(station_id)
        if row is None:
            raise GsmRegistryError(
                f"station #{station_id} not found",
                code="gsm_station_not_found",
            )
        return row

    def _require_vehicle_out(self, vehicle_id: int) -> VehicleOut:
        return self._vehicle_out(self._require_vehicle(vehicle_id))

    def _require_driver_out(self, driver_id: int) -> DriverOut:
        return self._driver_out(self._require_driver(driver_id))

    def _require_card_out(self, card_id: int) -> CardOut:
        return self._card_out(self._require_card(card_id))

    def _require_station_out(self, station_id: int) -> StationOut:
        return self._station_out(self._require_station(station_id))

    @staticmethod
    def _vehicle_out(row: dict[str, Any]) -> VehicleOut:
        return VehicleOut(
            id=int(row["id"]),
            name=str(row["name"]),
            plate_number=str(row["plate_number"]),
            tank_volume_liters=float(row["tank_volume_liters"]),
            norm_summer=float(row["norm_summer"]),
            norm_winter=float(row["norm_winter"]),
            primary_driver_id=(
                int(row["primary_driver_id"])
                if row.get("primary_driver_id") is not None
                else None
            ),
            is_active=bool(row.get("is_active", 1)),
        )

    @staticmethod
    def _driver_out(row: dict[str, Any]) -> DriverOut:
        return DriverOut(
            id=int(row["id"]),
            full_name=str(row["full_name"]),
            license_number=str(row["license_number"]),
            license_issued_at=row.get("license_issued_at"),
            personnel_number=row.get("personnel_number"),
            snils=row.get("snils"),
            is_active=bool(row.get("is_active", 1)),
        )

    @staticmethod
    def _card_out(row: dict[str, Any]) -> CardOut:
        return CardOut(
            id=int(row["id"]),
            card_number=str(row["card_number"]),
            vehicle_id=(
                int(row["vehicle_id"]) if row.get("vehicle_id") is not None else None
            ),
            assigned_at=str(row["assigned_at"]),
            archived_at=row.get("archived_at"),
        )

    @staticmethod
    def _station_out(row: dict[str, Any]) -> StationOut:
        return StationOut(
            id=int(row["id"]),
            address=str(row["address"]),
            brand=row.get("brand"),
            lat=float(row["lat"]) if row.get("lat") is not None else None,
            lon=float(row["lon"]) if row.get("lon") is not None else None,
            geocode_source=row.get("geocode_source"),
        )

    @staticmethod
    def _route_out(row: dict[str, Any]) -> RouteOut:
        return RouteOut(
            id=int(row["id"]),
            vehicle_id=int(row["vehicle_id"]),
            addr_a=str(row["addr_a"]),
            addr_b=str(row["addr_b"]),
            km=int(row["km"]),
            frequency=int(row.get("frequency") or 1),
            typical_station_ids=_parse_station_ids(row.get("typical_station_ids")),
        )


def _parse_station_ids(raw: Any) -> list[int]:
    if raw is None or raw == "":
        return []
    if isinstance(raw, list):
        return [int(x) for x in raw]
    try:
        parsed = json.loads(raw) if isinstance(raw, str) else raw
    except (TypeError, ValueError, json.JSONDecodeError):
        return []
    if not isinstance(parsed, list):
        return []
    out: list[int] = []
    for item in parsed:
        try:
            out.append(int(item))
        except (TypeError, ValueError):
            continue
    return out
