"""Task T8: generation service reads max_daily_km and station coords.

Acceptance:
- Default max_daily_km=700 when gsm_setting is absent
- Setting overrides the default
- Stations with lat/lon are passed as station_coords
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

import pytest

from app.repositories.gsm_repository import GsmRepository
from app.services.gsm_generation_service import GsmGenerationService
from core import kp_db_schema
from core.gsm.generator import GenerateResult, ProblematicDay
from core.gsm.geo import GeoPoint


def _fresh_db(tmp_path: Path, name: str = "gsm_gen_svc.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


@pytest.fixture()
def repo(tmp_path: Path) -> GsmRepository:
    return GsmRepository(db_path=_fresh_db(tmp_path))


@pytest.fixture()
def service(repo: GsmRepository) -> GsmGenerationService:
    return GsmGenerationService(
        repo=repo, holidays=frozenset(), extra_workdays=frozenset()
    )


def _seed_vehicle(repo: GsmRepository) -> int:
    driver_id = repo.create_driver(
        full_name="Тестов Водитель",
        license_number="00 00 000000",
    )
    return repo.create_vehicle(
        name="Test Car",
        plate_number="О 000 АА 44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )


def _empty_generate_result() -> GenerateResult:
    return GenerateResult(days=(), unsolvable=None, warnings=())


def _capture_generate(monkeypatch: pytest.MonkeyPatch) -> dict[str, Any]:
    captured: dict[str, Any] = {}

    def fake_generate(**kwargs: Any) -> GenerateResult:
        captured.clear()
        captured.update(kwargs)
        return _empty_generate_result()

    monkeypatch.setattr(
        "app.services.gsm_generation_service.generate", fake_generate
    )
    return captured


def test_generate_passes_default_max_daily_km(
    service: GsmGenerationService,
    repo: GsmRepository,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    vehicle_id = _seed_vehicle(repo)
    assert repo.get_setting("max_daily_km") is None
    captured = _capture_generate(monkeypatch)

    service.generate(
        vehicle_id=vehicle_id,
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        fuel_start=20.0,
        odometer_start=10_000,
    )

    assert captured["max_daily_km"] == 700


def test_generate_passes_override_max_daily_km(
    service: GsmGenerationService,
    repo: GsmRepository,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    vehicle_id = _seed_vehicle(repo)
    repo.set_setting("max_daily_km", "400")
    captured = _capture_generate(monkeypatch)

    service.generate(
        vehicle_id=vehicle_id,
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        fuel_start=20.0,
        odometer_start=10_000,
    )

    assert captured["max_daily_km"] == 400


def test_generate_passes_station_coords(
    service: GsmGenerationService,
    repo: GsmRepository,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    vehicle_id = _seed_vehicle(repo)
    with_coords = repo.create_station(
        address="АЗС с координатами",
        brand="TATNEFT",
        lat=57.76,
        lon=40.92,
    )
    repo.create_station(address="АЗС без координат", brand="ТНК")
    captured = _capture_generate(monkeypatch)

    service.generate(
        vehicle_id=vehicle_id,
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        fuel_start=20.0,
        odometer_start=10_000,
    )

    coords = captured["station_coords"]
    assert with_coords in coords
    assert coords[with_coords] == GeoPoint(lat=57.76, lon=40.92)
    assert len(coords) == 1


def test_generate_maps_problematic_days(
    service: GsmGenerationService,
    repo: GsmRepository,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Core ProblematicDay is mapped onto WaybillGenerateResult (200, not gsm_unsolvable)."""
    vehicle_id = _seed_vehicle(repo)
    problem = ProblematicDay(
        date=date(2025, 4, 7),
        reason="manual_intervention",
        detail="не удалось сжечь 51.2 л: требуется ручная доработка",
        fuel_before=40.1,
        fuel_to_issue=54.57,
        tank_volume=70.0,
    )

    def fake_generate(**kwargs: Any) -> GenerateResult:
        return GenerateResult(
            days=(),
            unsolvable=None,
            warnings=("weekend_anchor",),
            problematic_days=(problem,),
        )

    monkeypatch.setattr("app.services.gsm_generation_service.generate", fake_generate)

    result = service.generate(
        vehicle_id=vehicle_id,
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        fuel_start=20.0,
        odometer_start=10_000,
    )

    assert result.warnings == ["weekend_anchor"]
    assert result.manual_days == 1
    assert len(result.problematic_days) == 1
    out = result.problematic_days[0]
    assert out.date == "2025-04-07"
    assert out.reason == "manual_intervention"
    assert out.detail == problem.detail
    assert out.fuel_before == pytest.approx(40.1)
    assert out.fuel_to_issue == pytest.approx(54.57)
    assert out.tank_volume == pytest.approx(70.0)


def _problem_day(*, driver_id: int, warnings: tuple[str, ...] = ("manual_intervention",)):
    from core.gsm.models import RouteRef, TankState, WaybillDay

    return WaybillDay(
        date=date(2025, 4, 7),
        driver_id=driver_id,
        route=RouteRef(route_id=1, addr_a="Завод", addr_b="Объект", km=190),
        tank=TankState(
            date=date(2025, 4, 7),
            fuel_start=20.0,
            fuel_issued=40.0,
            fuel_end=24.0,
            km=380,
            odometer_start=10_000,
            odometer_end=10_380,
        ),
        source="auto",
        warnings=warnings,
    )


def test_generate_persists_warning_details_for_problematic_day(
    service: GsmGenerationService,
    repo: GsmRepository,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    vehicle_id = _seed_vehicle(repo)
    driver_id = int(repo.get_vehicle(vehicle_id)["primary_driver_id"])
    detail = "не удалось сжечь 51.2 л: требуется ручная доработка"
    problem = ProblematicDay(
        date=date(2025, 4, 7),
        reason="manual_intervention",
        detail=detail,
        fuel_before=40.1,
        fuel_to_issue=54.57,
        tank_volume=70.0,
    )

    def fake_generate(**kwargs: Any) -> GenerateResult:
        return GenerateResult(
            days=(_problem_day(driver_id=driver_id),),
            unsolvable=None,
            warnings=(),
            problematic_days=(problem,),
        )

    monkeypatch.setattr("app.services.gsm_generation_service.generate", fake_generate)
    service.generate(
        vehicle_id=vehicle_id,
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        fuel_start=20.0,
        odometer_start=10_000,
    )

    listed = service.list_waybills(
        vehicle_id=vehicle_id,
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
    )
    monday = next(wb for wb in listed if wb.date == "2025-04-07")
    assert "manual_intervention" in monday.warnings
    assert monday.warning_details
    assert monday.warning_details[0].code == "manual_intervention"
    assert monday.warning_details[0].detail == detail

    listed_again = service.list_waybills(
        vehicle_id=vehicle_id,
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
    )
    monday_again = next(wb for wb in listed_again if wb.date == "2025-04-07")
    assert monday_again.warning_details[0].detail == detail


def test_string_warnings_json_parses_without_details(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    vehicle_id = _seed_vehicle(repo)
    driver_id = int(repo.get_vehicle(vehicle_id)["primary_driver_id"])
    repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date=date(2025, 4, 8),
        driver_id=driver_id,
        warnings_json='["hook_above_threshold"]',
        route_json="[]",
    )
    listed = service.list_waybills(
        vehicle_id=vehicle_id,
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
    )
    assert listed[0].warnings == ["hook_above_threshold"]
    assert listed[0].warning_details == []


def test_clean_day_has_empty_warning_details(
    service: GsmGenerationService,
    repo: GsmRepository,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    vehicle_id = _seed_vehicle(repo)
    driver_id = int(repo.get_vehicle(vehicle_id)["primary_driver_id"])

    def fake_generate(**kwargs: Any) -> GenerateResult:
        return GenerateResult(
            days=(_problem_day(driver_id=driver_id, warnings=()),),
            unsolvable=None,
            warnings=(),
        )

    monkeypatch.setattr("app.services.gsm_generation_service.generate", fake_generate)
    result = service.generate(
        vehicle_id=vehicle_id,
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        fuel_start=20.0,
        odometer_start=10_000,
    )
    assert result.waybills[0].warnings == []
    assert not result.waybills[0].warning_details
