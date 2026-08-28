"""Orchestrate GSM waybill generation: DB → core.gsm.generator → draft waybills."""

from __future__ import annotations

import json
from dataclasses import fields
from datetime import date, datetime
from typing import Any

from app.repositories.gsm_repository import GsmRepository
from app.schemas.gsm import (
    ProblematicDayOut,
    WaybillBulkGenerateResult,
    WaybillBulkVehicleError,
    WaybillBulkVehicleResult,
    WaybillGenerateResult,
    WaybillOut,
    WaybillRouteLeg,
    WaybillWarningDetail,
)
from app.services.gsm_registry_service import (
    DEFAULT_HOOK_THRESHOLD_KM,
    DEFAULT_MAX_DAILY_KM,
)
from core.gsm.balance import BalanceViolation, apply_day
from core.gsm.generator import LibraryRoute, generate
from core.gsm.geo import GeoPoint
from core.gsm.models import Transaction, WaybillDay
from core.gsm.season import SeasonSwitch, norm_for, parse_season_switches
from core.work_calendar import load_extra_workdays, load_holidays

_LIBRARY_ROUTE_FIELDS = frozenset(f.name for f in fields(LibraryRoute))

_PROTECTED_STATUSES = frozenset({"confirmed", "exported"})


class GsmGenerationError(ValueError):
    """Domain error for generation (maps to 404/409/422)."""

    def __init__(
        self,
        message: str,
        *,
        code: str = "gsm_generation_error",
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.code = code
        self.details = details or {}


class GsmGenerationService:
    """Load vehicle context, run pure generator, persist draft waybills."""

    def __init__(
        self,
        *,
        repo: GsmRepository,
        holidays: frozenset[date] | set[date] | None = None,
        extra_workdays: frozenset[date] | set[date] | None = None,
    ) -> None:
        self._repo = repo
        self._holidays = holidays
        self._extra_workdays = extra_workdays

    def generate(
        self,
        *,
        vehicle_id: int,
        period_from: date,
        period_to: date,
        force: bool = False,
        fuel_start: float | None = None,
        odometer_start: int | None = None,
    ) -> WaybillGenerateResult:
        if period_to < period_from:
            raise GsmGenerationError(
                "period_to must be >= period_from",
                code="gsm_invalid_period",
            )

        vehicle = self._repo.get_vehicle(vehicle_id)
        if vehicle is None:
            raise GsmGenerationError(
                f"vehicle #{vehicle_id} not found",
                code="gsm_vehicle_not_found",
            )

        driver_id = vehicle.get("primary_driver_id")
        if driver_id is None:
            raise GsmGenerationError(
                "vehicle has no primary_driver_id",
                code="gsm_driver_required",
            )
        driver_id = int(driver_id)

        existing = self._repo.list_waybills(
            vehicle_id=vehicle_id,
            period_from=period_from,
            period_to=period_to,
        )
        protected = [row for row in existing if row.get("status") in _PROTECTED_STATUSES]
        if protected and not force:
            raise GsmGenerationError(
                "period has confirmed/exported waybills; pass force=true to overwrite",
                code="gsm_confirmed_conflict",
                details={
                    "dates": [str(row["date"]) for row in protected],
                    "count": len(protected),
                },
            )

        start_fuel, start_odo = self._resolve_start(
            vehicle_id=vehicle_id,
            period_from=period_from,
            fuel_start=fuel_start,
            odometer_start=odometer_start,
        )

        tx_rows = self._repo.list_transactions(
            vehicle_id=vehicle_id,
            period_from=period_from,
            period_to=period_to,
        )
        transactions = tuple(self._to_transaction(row) for row in tx_rows)
        routes = self._load_routes()
        if not routes and transactions:
            raise GsmGenerationError(
                "vehicle has no routes in library",
                code="gsm_routes_required",
            )

        season_switches = self._resolve_season_switches()
        hook_threshold = self._resolve_hook_threshold()
        max_daily_km = self._resolve_max_daily_km()
        station_coords = self._load_station_coords()
        holidays = (
            frozenset(self._holidays)
            if self._holidays is not None
            else frozenset(load_holidays())
        )
        extra_workdays = (
            frozenset(self._extra_workdays)
            if self._extra_workdays is not None
            else frozenset(load_extra_workdays())
        )

        result = generate(
            transactions=transactions,
            routes=routes,
            hooks={},
            driver_id=driver_id,
            tank_volume_liters=float(vehicle["tank_volume_liters"]),
            norm_summer=float(vehicle["norm_summer"]),
            norm_winter=float(vehicle["norm_winter"]),
            season_switches=season_switches,
            fuel_start=start_fuel,
            odometer_start=start_odo,
            hook_threshold_km=hook_threshold,
            holidays=holidays,
            extra_workdays=extra_workdays,
            max_daily_km=max_daily_km,
            station_coords=station_coords,
            own_vehicle_id=vehicle_id,
        )

        if force:
            self._repo.delete_waybills_in_period(
                vehicle_id=vehicle_id,
                period_from=period_from,
                period_to=period_to,
            )
        else:
            self._repo.delete_waybills_in_period(
                vehicle_id=vehicle_id,
                period_from=period_from,
                period_to=period_to,
                statuses=("draft",),
            )

        waybills: list[WaybillOut] = []
        problems_by_date = {problem.date: problem for problem in result.problematic_days}
        for day in result.days:
            route_json = _serialize_waybill_route(day)
            warnings_json = _serialize_warnings_json(day, problems_by_date.get(day.date))
            waybill_id = self._repo.upsert_waybill(
                vehicle_id=vehicle_id,
                date=day.date,
                driver_id=day.driver_id,
                status="draft",
                source=day.source,
                odometer_start=day.tank.odometer_start,
                odometer_end=day.tank.odometer_end,
                fuel_start=day.tank.fuel_start,
                fuel_issued=day.tank.fuel_issued,
                fuel_end=day.tank.fuel_end,
                route_json=route_json,
                warnings_json=warnings_json,
            )
            waybills.append(
                self._waybill_out(
                    {
                        "id": waybill_id,
                        "vehicle_id": vehicle_id,
                        "date": day.date.isoformat(),
                        "driver_id": day.driver_id,
                        "status": "draft",
                        "source": day.source,
                        "odometer_start": day.tank.odometer_start,
                        "odometer_end": day.tank.odometer_end,
                        "fuel_start": day.tank.fuel_start,
                        "fuel_issued": day.tank.fuel_issued,
                        "fuel_end": day.tank.fuel_end,
                        "route_json": route_json,
                        "warnings_json": warnings_json,
                    }
                )
            )

        problematic_days = [
            ProblematicDayOut(
                date=problem.date.isoformat(),
                reason=problem.reason,
                detail=problem.detail,
                fuel_before=problem.fuel_before,
                fuel_to_issue=problem.fuel_to_issue,
                tank_volume=problem.tank_volume,
            )
            for problem in result.problematic_days
        ]
        return WaybillGenerateResult(
            waybills=waybills,
            warnings=list(result.warnings),
            days_created=len(waybills),
            problematic_days=problematic_days,
            manual_days=len(problematic_days),
        )

    def generate_bulk(
        self,
        *,
        vehicle_ids: list[int],
        period_from: date,
        period_to: date,
        force: bool = False,
    ) -> WaybillBulkGenerateResult:
        if period_to < period_from:
            raise GsmGenerationError(
                "period_to must be >= period_from",
                code="gsm_invalid_period",
            )
        results: list[WaybillBulkVehicleResult] = []
        for vehicle_id in vehicle_ids:
            try:
                generated = self.generate(
                    vehicle_id=vehicle_id,
                    period_from=period_from,
                    period_to=period_to,
                    force=force,
                )
                results.append(
                    WaybillBulkVehicleResult(
                        vehicle_id=vehicle_id,
                        ok=True,
                        result=generated,
                    )
                )
            except GsmGenerationError as exc:
                results.append(
                    WaybillBulkVehicleResult(
                        vehicle_id=vehicle_id,
                        ok=False,
                        error=WaybillBulkVehicleError(
                            code=exc.code,
                            message=str(exc),
                        ),
                    )
                )
        return WaybillBulkGenerateResult(results=results)

    def list_waybills(
        self,
        *,
        period_from: date,
        period_to: date,
        vehicle_id: int | None = None,
    ) -> list[WaybillOut]:
        if period_to < period_from:
            raise GsmGenerationError(
                "period_to must be >= period_from",
                code="gsm_invalid_period",
            )
        if vehicle_id is not None and self._repo.get_vehicle(vehicle_id) is None:
            raise GsmGenerationError(
                f"vehicle #{vehicle_id} not found",
                code="gsm_vehicle_not_found",
            )
        rows = self._repo.list_waybills(
            vehicle_id=vehicle_id,
            period_from=period_from,
            period_to=period_to,
        )
        return [self._waybill_out(row) for row in rows]

    def patch_waybill(
        self,
        waybill_id: int,
        *,
        driver_id: int | None = None,
        km: int | None = None,
        route: list[WaybillRouteLeg] | list[dict[str, Any]] | None = None,
    ) -> WaybillOut:
        row = self._repo.get_waybill_by_id(waybill_id)
        if row is None:
            raise GsmGenerationError(
                f"waybill #{waybill_id} not found",
                code="gsm_waybill_not_found",
            )
        status = str(row.get("status") or "draft")
        if status in _PROTECTED_STATUSES:
            raise GsmGenerationError(
                "waybill is locked (confirmed/exported)",
                code="gsm_waybill_locked",
            )
        vehicle_id = int(row["vehicle_id"])
        vehicle = self._repo.get_vehicle(vehicle_id)
        if vehicle is None:
            raise GsmGenerationError(
                f"vehicle #{vehicle_id} not found",
                code="gsm_vehicle_not_found",
            )

        day = _as_date(row["date"])
        new_driver_id = int(driver_id) if driver_id is not None else int(row["driver_id"])
        if self._repo.get_driver(new_driver_id) is None:
            raise GsmGenerationError(
                f"driver #{new_driver_id} not found",
                code="gsm_driver_not_found",
            )

        legs = self._normalize_route_input(route)
        if legs is None:
            legs = _parse_route_json(row.get("route_json"))
        if km is not None:
            legs = _apply_km_to_legs(legs, int(km))
        total_km = sum(leg.km for leg in legs)

        fuel_issued = float(row.get("fuel_issued") or 0.0)
        fuel_start = float(row["fuel_start"]) if row.get("fuel_start") is not None else 0.0
        odometer_start = (
            int(row["odometer_start"]) if row.get("odometer_start") is not None else 0
        )
        tank = self._apply_day_balance(
            day=day,
            vehicle=vehicle,
            fuel_start=fuel_start,
            fuel_issued=fuel_issued,
            km=total_km,
            odometer_start=odometer_start,
        )

        downstream = self._repo.list_waybills_after(
            vehicle_id=vehicle_id,
            after_date=day,
        )
        if any(
            str(later.get("status") or "draft") in _PROTECTED_STATUSES
            for later in downstream
        ):
            raise GsmGenerationError(
                "cannot edit waybill: later confirmed/exported waybill exists",
                code="gsm_chain_locked",
            )

        route_json = _route_legs_to_json(legs)
        warnings_json = row.get("warnings_json")

        self._repo.upsert_waybill(
            vehicle_id=vehicle_id,
            date=day,
            driver_id=new_driver_id,
            status=status,
            source="manual",
            odometer_start=tank.odometer_start,
            odometer_end=tank.odometer_end,
            fuel_start=tank.fuel_start,
            fuel_issued=tank.fuel_issued,
            fuel_end=tank.fuel_end,
            route_json=route_json,
            warnings_json=warnings_json,
        )
        rechained = self._rechain_downstream(
            vehicle_id=vehicle_id,
            after_date=day,
            fuel=tank.fuel_end,
            odometer=tank.odometer_end,
            vehicle=vehicle,
        )
        updated = self._repo.get_waybill_by_id(waybill_id)
        assert updated is not None
        return self._waybill_out(updated, rechained_draft_days=rechained)

    def create_waybill(
        self,
        *,
        vehicle_id: int,
        day: date,
        driver_id: int,
        route: list[WaybillRouteLeg] | list[dict[str, Any]],
        fuel_issued: float = 0.0,
        fuel_start: float | None = None,
        odometer_start: int | None = None,
    ) -> WaybillOut:
        vehicle = self._repo.get_vehicle(vehicle_id)
        if vehicle is None:
            raise GsmGenerationError(
                f"vehicle #{vehicle_id} not found",
                code="gsm_vehicle_not_found",
            )
        if self._repo.get_driver(driver_id) is None:
            raise GsmGenerationError(
                f"driver #{driver_id} not found",
                code="gsm_driver_not_found",
            )
        existing = self._repo.get_waybill(vehicle_id, day)
        if existing is not None:
            raise GsmGenerationError(
                f"waybill for vehicle #{vehicle_id} on {day.isoformat()} already exists",
                code="gsm_waybill_conflict",
                details={"date": day.isoformat(), "id": existing.get("id")},
            )

        legs = self._normalize_route_input(route)
        if not legs:
            raise GsmGenerationError(
                "route must have at least one leg",
                code="gsm_invalid_route",
            )
        total_km = sum(leg.km for leg in legs)

        start_fuel, start_odo = self._resolve_chain_start(
            vehicle_id=vehicle_id,
            day=day,
            fuel_start=fuel_start,
            odometer_start=odometer_start,
        )
        tank = self._apply_day_balance(
            day=day,
            vehicle=vehicle,
            fuel_start=start_fuel,
            fuel_issued=float(fuel_issued),
            km=total_km,
            odometer_start=start_odo,
        )
        route_json = _route_legs_to_json(legs)
        waybill_id = self._repo.upsert_waybill(
            vehicle_id=vehicle_id,
            date=day,
            driver_id=driver_id,
            status="draft",
            source="manual",
            odometer_start=tank.odometer_start,
            odometer_end=tank.odometer_end,
            fuel_start=tank.fuel_start,
            fuel_issued=tank.fuel_issued,
            fuel_end=tank.fuel_end,
            route_json=route_json,
            warnings_json=None,
        )
        self._rechain_downstream(
            vehicle_id=vehicle_id,
            after_date=day,
            fuel=tank.fuel_end,
            odometer=tank.odometer_end,
            vehicle=vehicle,
        )
        created = self._repo.get_waybill_by_id(waybill_id)
        assert created is not None
        return self._waybill_out(created)

    def confirm_waybill(self, waybill_id: int) -> WaybillOut:
        row = self._repo.get_waybill_by_id(waybill_id)
        if row is None:
            raise GsmGenerationError(
                f"waybill #{waybill_id} not found",
                code="gsm_waybill_not_found",
            )
        day = _as_date(row["date"])
        self._repo.upsert_waybill(
            vehicle_id=int(row["vehicle_id"]),
            date=day,
            driver_id=int(row["driver_id"]),
            status="confirmed",
            source=str(row.get("source") or "auto"),
            odometer_start=_as_optional_int(row.get("odometer_start")),
            odometer_end=_as_optional_int(row.get("odometer_end")),
            fuel_start=_as_optional_float(row.get("fuel_start")),
            fuel_issued=_as_optional_float(row.get("fuel_issued")),
            fuel_end=_as_optional_float(row.get("fuel_end")),
            route_json=str(row.get("route_json") or "[]"),
            warnings_json=row.get("warnings_json"),
        )
        updated = self._repo.get_waybill_by_id(waybill_id)
        assert updated is not None
        return self._waybill_out(updated)

    def _resolve_chain_start(
        self,
        *,
        vehicle_id: int,
        day: date,
        fuel_start: float | None,
        odometer_start: int | None,
    ) -> tuple[float, int]:
        prev = self._repo.get_last_waybill(vehicle_id, before=day)
        if prev is not None:
            fuel = prev.get("fuel_end")
            odo = prev.get("odometer_end")
            if fuel is None or odo is None:
                raise GsmGenerationError(
                    "previous waybill missing fuel_end/odometer_end",
                    code="gsm_start_required",
                )
            return float(fuel), int(odo)
        if fuel_start is None or odometer_start is None:
            raise GsmGenerationError(
                "fuel_start and odometer_start required when no previous waybill exists",
                code="gsm_start_required",
            )
        return float(fuel_start), int(odometer_start)

    def _apply_day_balance(
        self,
        *,
        day: date,
        vehicle: dict[str, Any],
        fuel_start: float,
        fuel_issued: float,
        km: int,
        odometer_start: int,
    ):
        season_switches = self._resolve_season_switches()
        norm = norm_for(
            day,
            norm_summer=float(vehicle["norm_summer"]),
            norm_winter=float(vehicle["norm_winter"]),
            switches=season_switches,
        )
        try:
            return apply_day(
                day,
                fuel_start=fuel_start,
                fuel_issued=fuel_issued,
                km=km,
                odometer_start=odometer_start,
                norm_per_100km=norm,
                tank_volume_liters=float(vehicle["tank_volume_liters"]),
            )
        except BalanceViolation as exc:
            raise GsmGenerationError(
                str(exc),
                code="gsm_balance_violation",
                details={"date": day.isoformat()},
            ) from exc

    def _rechain_downstream(
        self,
        *,
        vehicle_id: int,
        after_date: date,
        fuel: float,
        odometer: int,
        vehicle: dict[str, Any],
    ) -> int:
        """Recompute fuel/odo chain for DRAFT days after ``after_date``.

        Confirmed/exported days are left untouched; the chain continues from
        their stored fuel_end/odometer_end. Returns the number of recomputed
        draft days.
        """
        fuel_cur = float(fuel)
        odo_cur = int(odometer)
        rechained = 0
        downstream = self._repo.list_waybills_after(
            vehicle_id=vehicle_id,
            after_date=after_date,
        )
        for row in downstream:
            day = _as_date(row["date"])
            status = str(row.get("status") or "draft")
            if status in _PROTECTED_STATUSES:
                fuel_end = row.get("fuel_end")
                odo_end = row.get("odometer_end")
                if fuel_end is None or odo_end is None:
                    raise GsmGenerationError(
                        f"protected waybill on {day.isoformat()} missing fuel_end/odometer_end",
                        code="gsm_start_required",
                    )
                fuel_cur = float(fuel_end)
                odo_cur = int(odo_end)
                continue

            legs = _parse_route_json(row.get("route_json"))
            km = sum(leg.km for leg in legs)
            fuel_issued = float(row.get("fuel_issued") or 0.0)
            tank = self._apply_day_balance(
                day=day,
                vehicle=vehicle,
                fuel_start=fuel_cur,
                fuel_issued=fuel_issued,
                km=km,
                odometer_start=odo_cur,
            )
            self._repo.upsert_waybill(
                vehicle_id=vehicle_id,
                date=day,
                driver_id=int(row["driver_id"]),
                status=status,
                source=str(row.get("source") or "auto"),
                odometer_start=tank.odometer_start,
                odometer_end=tank.odometer_end,
                fuel_start=tank.fuel_start,
                fuel_issued=tank.fuel_issued,
                fuel_end=tank.fuel_end,
                route_json=str(row.get("route_json") or "[]"),
                warnings_json=row.get("warnings_json"),
            )
            fuel_cur = tank.fuel_end
            odo_cur = tank.odometer_end
            rechained += 1
        return rechained

    @staticmethod
    def _normalize_route_input(
        route: list[WaybillRouteLeg] | list[dict[str, Any]] | None,
    ) -> list[WaybillRouteLeg] | None:
        if route is None:
            return None
        legs: list[WaybillRouteLeg] = []
        for item in route:
            if isinstance(item, WaybillRouteLeg):
                legs.append(item)
                continue
            if not isinstance(item, dict):
                raise GsmGenerationError("invalid route leg", code="gsm_invalid_route")
            from_addr = item.get("from") or item.get("from_addr") or item.get("addr_a") or ""
            to_addr = item.get("to") or item.get("to_addr") or item.get("addr_b") or ""
            legs.append(
                WaybillRouteLeg(
                    **{
                        "from": str(from_addr),
                        "to": str(to_addr),
                        "km": int(item.get("km") or 0),
                        "route_id": item.get("route_id"),
                        "station_id": item.get("station_id"),
                        "dep_time": item.get("dep_time"),
                        "arr_time": item.get("arr_time"),
                    }
                )
            )
        return legs

    def _resolve_start(
        self,
        *,
        vehicle_id: int,
        period_from: date,
        fuel_start: float | None,
        odometer_start: int | None,
    ) -> tuple[float, int]:
        last = self._repo.get_last_confirmed_waybill(vehicle_id, before=period_from)
        if last is not None:
            fuel = last.get("fuel_end")
            odo = last.get("odometer_end")
            if fuel is None or odo is None:
                raise GsmGenerationError(
                    "last confirmed waybill missing fuel_end/odometer_end",
                    code="gsm_start_required",
                )
            return float(fuel), int(odo)

        if fuel_start is None or odometer_start is None:
            raise GsmGenerationError(
                "fuel_start and odometer_start required when no confirmed waybill exists",
                code="gsm_start_required",
            )
        return float(fuel_start), int(odometer_start)

    def _resolve_season_switches(self) -> tuple[SeasonSwitch, ...]:
        raw = self._repo.get_setting("season_switches")
        try:
            return parse_season_switches(raw)
        except ValueError as exc:
            raise GsmGenerationError(
                f"invalid season_switches setting: {raw!r}",
                code="gsm_settings_invalid",
            ) from exc

    def _resolve_hook_threshold(self) -> float:
        raw = self._repo.get_setting("hook_threshold_km")
        if raw is None:
            return DEFAULT_HOOK_THRESHOLD_KM
        try:
            value = float(raw)
        except (TypeError, ValueError) as exc:
            raise GsmGenerationError(
                f"invalid hook_threshold_km setting: {raw!r}",
                code="gsm_settings_invalid",
            ) from exc
        if value <= 0:
            raise GsmGenerationError(
                "hook_threshold_km must be > 0",
                code="gsm_settings_invalid",
            )
        return value

    def _resolve_max_daily_km(self) -> int:
        raw = self._repo.get_setting("max_daily_km")
        if raw is None:
            return DEFAULT_MAX_DAILY_KM
        try:
            value = int(raw)
        except (TypeError, ValueError) as exc:
            raise GsmGenerationError(
                f"invalid max_daily_km setting: {raw!r}",
                code="gsm_settings_invalid",
            ) from exc
        if value <= 0:
            raise GsmGenerationError(
                "max_daily_km must be > 0",
                code="gsm_settings_invalid",
            )
        return value

    def _load_station_coords(self) -> dict[int, GeoPoint]:
        coords: dict[int, GeoPoint] = {}
        for row in self._repo.list_stations():
            lat = row.get("lat")
            lon = row.get("lon")
            if lat is None or lon is None:
                continue
            coords[int(row["id"])] = GeoPoint(lat=float(lat), lon=float(lon))
        return coords

    def _load_routes(self) -> tuple[LibraryRoute, ...]:
        rows = self._repo.list_routes()
        routes: list[LibraryRoute] = []
        for row in rows:
            typical = self._parse_typical_station_ids(row.get("typical_station_ids"))
            payload: dict[str, Any] = {
                "route_id": int(row["id"]),
                "addr_a": str(row["addr_a"]),
                "addr_b": str(row["addr_b"]),
                "km": int(row["km"]),
                "frequency": int(row.get("frequency") or 1),
                "typical_station_ids": typical,
            }
            if "vehicle_id" in _LIBRARY_ROUTE_FIELDS:
                raw_vid = row.get("vehicle_id")
                payload["vehicle_id"] = int(raw_vid) if raw_vid is not None else 0
            if "point_a" in _LIBRARY_ROUTE_FIELDS:
                payload["point_a"] = None
            if "point_b" in _LIBRARY_ROUTE_FIELDS:
                payload["point_b"] = None
            routes.append(LibraryRoute(**payload))
        return tuple(routes)

    @staticmethod
    def _parse_typical_station_ids(raw: Any) -> tuple[int, ...]:
        if raw is None or raw == "":
            return ()
        if isinstance(raw, (list, tuple)):
            return tuple(int(x) for x in raw)
        try:
            parsed = json.loads(str(raw))
        except (TypeError, ValueError, json.JSONDecodeError):
            return ()
        if not isinstance(parsed, list):
            return ()
        return tuple(int(x) for x in parsed)

    @staticmethod
    def _to_transaction(row: dict[str, Any]) -> Transaction:
        ts_raw = row["ts"]
        if isinstance(ts_raw, datetime):
            ts = ts_raw
        else:
            ts = datetime.fromisoformat(str(ts_raw))
        qty = row.get("qty_liters")
        station_id = row.get("station_id")
        return Transaction(
            card_id=int(row["card_id"]),
            ts=ts,
            service_type=str(row["service_type"]),
            qty_liters=float(qty) if qty is not None else None,
            amount=float(row["amount"]),
            station_id=int(station_id) if station_id is not None else None,
            raw_address=str(row.get("raw_address") or ""),
        )

    @staticmethod
    def _waybill_out(row: dict[str, Any], *, rechained_draft_days: int = 0) -> WaybillOut:
        route_legs = _parse_route_json(row.get("route_json"))
        warnings, warning_details = _parse_warnings_json(row.get("warnings_json"))
        km = sum(leg.km for leg in route_legs)
        day = row["date"]
        day_str = day.isoformat() if isinstance(day, date) else str(day)
        return WaybillOut(
            id=int(row["id"]),
            vehicle_id=int(row["vehicle_id"]),
            date=day_str,
            driver_id=int(row["driver_id"]),
            status=str(row["status"]),
            source=str(row.get("source") or "auto"),
            odometer_start=_as_optional_int(row.get("odometer_start")),
            odometer_end=_as_optional_int(row.get("odometer_end")),
            fuel_start=_as_optional_float(row.get("fuel_start")),
            fuel_issued=_as_optional_float(row.get("fuel_issued")),
            fuel_end=_as_optional_float(row.get("fuel_end")),
            km=km,
            route=route_legs,
            warnings=warnings,
            warning_details=warning_details,
            rechained_draft_days=rechained_draft_days,
        )


def _serialize_waybill_route(day: WaybillDay) -> str:
    """Persist generator legs; fall back to a single library route if legs empty."""
    if day.legs:
        payload = [
            {
                "from": leg.addr_a,
                "to": leg.addr_b,
                "km": leg.km,
                "route_id": leg.route_id,
            }
            for leg in day.legs
        ]
    else:
        payload = [
            {
                "from": day.route.addr_a,
                "to": day.route.addr_b,
                "km": day.route.km,
                "route_id": day.route.route_id,
            }
        ]
    return json.dumps(payload, ensure_ascii=False)


def _parse_route_json(raw: Any) -> list[WaybillRouteLeg]:
    if raw is None or raw == "":
        return []
    try:
        parsed = json.loads(raw) if isinstance(raw, str) else raw
    except (TypeError, ValueError, json.JSONDecodeError):
        return []
    if not isinstance(parsed, list):
        return []
    legs: list[WaybillRouteLeg] = []
    for item in parsed:
        if not isinstance(item, dict):
            continue
        from_addr = item.get("from") or item.get("from_addr") or item.get("addr_a") or ""
        to_addr = item.get("to") or item.get("to_addr") or item.get("addr_b") or ""
        km = int(item.get("km") or 0)
        legs.append(
            WaybillRouteLeg(
                **{
                    "from": str(from_addr),
                    "to": str(to_addr),
                    "km": km,
                    "route_id": item.get("route_id"),
                    "station_id": item.get("station_id"),
                    "dep_time": item.get("dep_time"),
                    "arr_time": item.get("arr_time"),
                }
            )
        )
    return legs


def _serialize_warnings_json(day: WaybillDay, problem: Any | None) -> str | None:
    if not day.warnings:
        return None
    if problem is None:
        payload: list[Any] = list(day.warnings)
    else:
        payload = [
            {"code": code, "detail": problem.detail} for code in day.warnings
        ]
    return json.dumps(payload, ensure_ascii=False)


def _parse_warnings_json(raw: Any) -> tuple[list[str], list[WaybillWarningDetail]]:
    if raw is None or raw == "":
        return [], []
    try:
        parsed = json.loads(raw) if isinstance(raw, str) else raw
    except (TypeError, ValueError, json.JSONDecodeError):
        return [], []
    if not isinstance(parsed, list):
        return [], []
    codes: list[str] = []
    details: list[WaybillWarningDetail] = []
    for item in parsed:
        if isinstance(item, str):
            codes.append(item)
            continue
        if isinstance(item, dict) and item.get("code"):
            code = str(item["code"])
            codes.append(code)
            detail = item.get("detail")
            if detail:
                details.append(WaybillWarningDetail(code=code, detail=str(detail)))
            continue
        codes.append(str(item))
    return codes, details


def _as_optional_float(value: Any) -> float | None:
    if value is None:
        return None
    return float(value)


def _as_optional_int(value: Any) -> int | None:
    if value is None:
        return None
    return int(value)


def _as_date(value: Any) -> date:
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value))


def _apply_km_to_legs(legs: list[WaybillRouteLeg], km: int) -> list[WaybillRouteLeg]:
    """Set total day km: single-leg replace, empty → placeholder, multi → adjust first."""
    if not legs:
        return [
            WaybillRouteLeg(
                **{"from": "", "to": "", "km": km}
            )
        ]
    if len(legs) == 1:
        leg = legs[0]
        return [
            WaybillRouteLeg(
                **{
                    "from": leg.from_addr,
                    "to": leg.to_addr,
                    "km": km,
                    "route_id": leg.route_id,
                    "station_id": leg.station_id,
                    "dep_time": leg.dep_time,
                    "arr_time": leg.arr_time,
                }
            )
        ]
    current = sum(leg.km for leg in legs)
    delta = km - current
    first = legs[0]
    new_first_km = max(0, first.km + delta)
    updated = [
        WaybillRouteLeg(
            **{
                "from": first.from_addr,
                "to": first.to_addr,
                "km": new_first_km,
                "route_id": first.route_id,
                "station_id": first.station_id,
                "dep_time": first.dep_time,
                "arr_time": first.arr_time,
            }
        ),
        *legs[1:],
    ]
    # If still short (first was clamped), bump last leg
    total = sum(leg.km for leg in updated)
    if total != km and updated:
        last = updated[-1]
        updated[-1] = WaybillRouteLeg(
            **{
                "from": last.from_addr,
                "to": last.to_addr,
                "km": max(0, last.km + (km - total)),
                "route_id": last.route_id,
                "station_id": last.station_id,
                "dep_time": last.dep_time,
                "arr_time": last.arr_time,
            }
        )
    return updated


def _route_legs_to_json(legs: list[WaybillRouteLeg]) -> str:
    payload = []
    for leg in legs:
        item: dict[str, Any] = {
            "from": leg.from_addr,
            "to": leg.to_addr,
            "km": leg.km,
        }
        if leg.route_id is not None:
            item["route_id"] = leg.route_id
        if leg.station_id is not None:
            item["station_id"] = leg.station_id
        if leg.dep_time is not None:
            item["dep_time"] = leg.dep_time
        if leg.arr_time is not None:
            item["arr_time"] = leg.arr_time
        payload.append(item)
    return json.dumps(payload, ensure_ascii=False)
