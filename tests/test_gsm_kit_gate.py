"""Unit tests for GSM month-close kit eligibility (stage 1 gate)."""

from __future__ import annotations

from datetime import date

from app.services.gsm_kit_gate import (
    CODE_CHAIN,
    CODE_RED,
    CODE_TAIL,
    evaluate_from_overview_row,
    evaluate_kit_vehicle,
)


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
