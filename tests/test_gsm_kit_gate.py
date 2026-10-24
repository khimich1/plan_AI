"""Unit tests for GSM month-close kit eligibility (stage 1 gate)."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from app.repositories.gsm_repository import GsmRepository
from app.services.gsm_kit_gate import (
    CODE_BOUNDARY,
    CODE_CHAIN,
    CODE_RED,
    CODE_TAIL,
    evaluate_from_overview_row,
    evaluate_kit_vehicle,
    evaluate_vehicle,
    filter_kit_vehicle_ids,
)
from core import kp_db_schema


def _row(**overrides: object) -> dict[str, object]:
    base: dict[str, object] = {
        "vehicle_id": 1,
        "open_before": 0,
        "open_before_month": None,
        "red_days": 0,
        "chain_prev_fuel_end": None,
        "chain_first_fuel_start": None,
        "chain_prev_odometer_end": None,
        "chain_first_odometer_start": None,
    }
    base.update(overrides)
    return base


def test_july_tail_blocks_august_kit() -> None:
    elig = evaluate_kit_vehicle(
        vehicle_id=7,
        period_from=date(2026, 8, 1),
        open_before=6,
        open_before_month="2026-07",
        chain_broken=False,
        red_days=0,
        purpose="kit",
    )
    assert elig.allowed is False
    assert elig.code == CODE_TAIL
    assert elig.message is not None
    assert "Июль" in elig.message


def test_tail_month_does_not_block_kit() -> None:
    elig = evaluate_kit_vehicle(
        vehicle_id=7,
        period_from=date(2026, 7, 1),
        open_before=6,
        open_before_month="2026-07",
        chain_broken=False,
        red_days=0,
        purpose="kit",
    )
    assert elig.allowed is True
    assert elig.code is None


def test_chain_broken_forbids_kit_allows_generate() -> None:
    kit = evaluate_kit_vehicle(
        vehicle_id=1,
        period_from=date(2026, 8, 1),
        open_before=0,
        open_before_month=None,
        chain_broken=True,
        red_days=0,
        purpose="kit",
    )
    gen = evaluate_kit_vehicle(
        vehicle_id=1,
        period_from=date(2026, 8, 1),
        open_before=0,
        open_before_month=None,
        chain_broken=True,
        red_days=0,
        purpose="generate",
    )
    assert kit.allowed is False
    assert kit.code == CODE_CHAIN
    assert gen.allowed is True
    assert gen.code is None


def test_red_forbids_kit_allows_generate() -> None:
    kit = evaluate_kit_vehicle(
        vehicle_id=1,
        period_from=date(2026, 8, 1),
        open_before=0,
        open_before_month=None,
        chain_broken=False,
        red_days=2,
        purpose="kit",
    )
    gen = evaluate_kit_vehicle(
        vehicle_id=1,
        period_from=date(2026, 8, 1),
        open_before=0,
        open_before_month=None,
        chain_broken=False,
        red_days=2,
        purpose="generate",
    )
    assert kit.allowed is False
    assert kit.code == CODE_RED
    assert gen.allowed is True


def test_july_tail_blocks_august_generate() -> None:
    elig = evaluate_kit_vehicle(
        vehicle_id=7,
        period_from=date(2026, 8, 1),
        open_before=6,
        open_before_month="2026-07",
        chain_broken=True,
        red_days=1,
        purpose="generate",
    )
    assert elig.allowed is False
    assert elig.code == CODE_TAIL


def test_red_preferred_over_tail_for_kit() -> None:
    elig = evaluate_kit_vehicle(
        vehicle_id=1,
        period_from=date(2026, 8, 1),
        open_before=6,
        open_before_month="2026-07",
        chain_broken=True,
        red_days=1,
        purpose="kit",
    )
    assert elig.allowed is False
    assert elig.code == CODE_RED


def test_chain_eps_below_threshold_allows_kit() -> None:
    elig = evaluate_from_overview_row(
        _row(
            chain_prev_fuel_end=10.0,
            chain_first_fuel_start=10.005,
            chain_prev_odometer_end=100,
            chain_first_odometer_start=100,
        ),
        period_from=date(2026, 8, 1),
        purpose="kit",
    )
    assert elig.allowed is True


def test_chain_eps_above_threshold_blocks_kit() -> None:
    elig = evaluate_from_overview_row(
        _row(
            chain_prev_fuel_end=10.0,
            chain_first_fuel_start=10.02,
            chain_prev_odometer_end=100,
            chain_first_odometer_start=100,
        ),
        period_from=date(2026, 8, 1),
        purpose="kit",
    )
    assert elig.allowed is False
    assert elig.code == CODE_CHAIN


def test_boundary_blocks_kit_when_closing_fuel_misses_next_fill() -> None:
    elig = evaluate_kit_vehicle(
        vehicle_id=1,
        period_from=date(2026, 8, 1),
        open_before=0,
        open_before_month=None,
        chain_broken=False,
        red_days=0,
        purpose="kit",
        closing_fuel=51.06,
        tank_volume_liters=70.0,
        next_first_fill=30.0,
    )
    assert elig.allowed is False
    assert elig.code == CODE_BOUNDARY
    assert elig.message is not None
    assert "Август" in elig.message
    assert "первую заправку" in elig.message


def test_boundary_allows_generate_and_stays_silent_without_later_fill() -> None:
    blocked_inputs = dict(
        vehicle_id=1,
        period_from=date(2026, 8, 1),
        open_before=0,
        open_before_month=None,
        chain_broken=False,
        red_days=0,
        closing_fuel=51.06,
        tank_volume_liters=70.0,
        next_first_fill=30.0,
    )
    generate = evaluate_kit_vehicle(**blocked_inputs, purpose="generate")
    assert generate.allowed is True
    assert generate.code is None

    silent = evaluate_kit_vehicle(
        vehicle_id=1,
        period_from=date(2026, 8, 1),
        open_before=0,
        open_before_month=None,
        chain_broken=False,
        red_days=0,
        purpose="kit",
        closing_fuel=51.06,
        tank_volume_liters=70.0,
        next_first_fill=None,
    )
    assert silent.allowed is True
    assert silent.code is None

    fits = evaluate_from_overview_row(
        _row(
            fuel_end_last=39.75,
            tank_volume_liters=70.0,
            next_first_fill_liters=30.0,
        ),
        period_from=date(2026, 8, 1),
        purpose="kit",
    )
    assert fits.allowed is True


def test_red_still_preferred_over_boundary() -> None:
    elig = evaluate_kit_vehicle(
        vehicle_id=1,
        period_from=date(2026, 8, 1),
        open_before=0,
        open_before_month=None,
        chain_broken=False,
        red_days=1,
        purpose="kit",
        closing_fuel=51.06,
        tank_volume_liters=70.0,
        next_first_fill=30.0,
    )
    assert elig.code == CODE_RED


def _fresh_db(tmp_path: Path) -> str:
    db_path = str(tmp_path / "gsm_kit_gate.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _seed_boundary_vehicle(repo: GsmRepository, *, fuel_end: float) -> int:
    driver_id = repo.create_driver(full_name="Водитель", license_number="00 00 000001")
    vehicle_id = repo.create_vehicle(
        name="Hyundai Palisade",
        plate_number="О 521 УХ 44",
        tank_volume_liters=70.0,
        norm_summer=14.5,
        norm_winter=16.0,
        primary_driver_id=driver_id,
    )
    repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date=date(2026, 8, 31),
        driver_id=driver_id,
        status="exported",
        fuel_start=32.8,
        fuel_issued=20.0,
        fuel_end=fuel_end,
        odometer_start=1000,
        odometer_end=1012,
    )
    return vehicle_id


def _insert_september_fills(repo: GsmRepository, vehicle_id: int) -> None:
    card_id = repo.create_card(
        card_number="521",
        vehicle_id=vehicle_id,
        assigned_at="2026-01-01",
    )
    batch_id = repo.create_import_batch(
        filename="sept.xls",
        uploaded_at="2026-09-02T00:00:00",
        period_from="2026-09-01",
        period_to="2026-09-30",
    )
    for ts, qty in (
        (datetime(2026, 9, 1, 9, 40, 57), 30.0),
        (datetime(2026, 9, 1, 9, 41, 44), 26.6),
    ):
        repo.insert_transaction(
            card_id=card_id,
            ts=ts.isoformat(timespec="seconds"),
            service_type="fuel",
            qty_liters=qty,
            amount=1000.0,
            raw_address="АЗС",
            batch_id=batch_id,
        )


def test_server_latch_refuses_kit_for_palisade_closing_fuel(tmp_path: Path) -> None:
    repo = GsmRepository(db_path=_fresh_db(tmp_path))
    vehicle_id = _seed_boundary_vehicle(repo, fuel_end=51.06)
    _insert_september_fills(repo, vehicle_id)
    period_from = date(2026, 8, 1)
    period_to = date(2026, 8, 31)

    allowed, blocked = filter_kit_vehicle_ids(
        repo,
        [vehicle_id],
        period_from=period_from,
        period_to=period_to,
        purpose="kit",
    )
    assert allowed == []
    assert len(blocked) == 1
    assert blocked[0].code == CODE_BOUNDARY
    assert blocked[0].message is not None
    assert "Август" in blocked[0].message

    generate = evaluate_vehicle(
        repo,
        vehicle_id,
        period_from=period_from,
        period_to=period_to,
        purpose="generate",
    )
    assert generate.allowed is True


def test_server_latch_silent_when_no_later_fill(tmp_path: Path) -> None:
    repo = GsmRepository(db_path=_fresh_db(tmp_path))
    vehicle_id = _seed_boundary_vehicle(repo, fuel_end=51.06)
    allowed, blocked = filter_kit_vehicle_ids(
        repo,
        [vehicle_id],
        period_from=date(2026, 8, 1),
        period_to=date(2026, 8, 31),
        purpose="kit",
    )
    assert blocked == []
    assert allowed == [vehicle_id]
