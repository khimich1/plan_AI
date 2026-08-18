"""Unit tests for core.gsm.balance + models (pure domain, no app I/O).

TDD / Task 8: these tests are written FIRST and must fail until
``core/gsm/models.py`` and ``core/gsm/balance.py`` exist.

Contract (Decision D5-style raise, like PlanBuildError):
    Violations of the tank corridor ``0 ≤ fuel_end ≤ tank_volume`` raise
    ``BalanceViolation`` (a ``RuntimeError`` subclass). Typed-failure
    return values are NOT used.

Public surface pinned by this file
-----------------------------------
``core.gsm.models``
    frozen+slots DTOs: ``Transaction``, ``Anchor``, ``WaybillDay``,
    ``TankState``, ``RouteRef``.

``core.gsm.balance``
    ``burn_for_km(km, norm_per_100km) -> float``
        ``round(km * norm_per_100km / 100, 2)``
    ``apply_day(day, *, fuel_start, fuel_issued, km, odometer_start,
                norm_per_100km, tank_volume_liters) -> TankState``
        ``fuel_end = fuel_start + fuel_issued - burn``;
        ``odometer_end = odometer_start + km``;
        raises ``BalanceViolation`` if fuel_end outside ``[0, tank]``.
    ``apply_day_chain(days, *, fuel_start, odometer_start,
                      norm_per_100km, tank_volume_liters) -> tuple[TankState, ...]``
        ``days`` is a sequence of ``(date, fuel_issued, km)``;
        each day's ``fuel_end`` / ``odometer_end`` become the next day's
        ``fuel_start`` / ``odometer_start``.
    ``BalanceViolation`` — raised on corridor breach.
"""

from __future__ import annotations

import ast
import dataclasses
from datetime import date, datetime
from pathlib import Path

import pytest

from core.gsm.balance import (
    BalanceViolation,
    apply_day,
    apply_day_chain,
    burn_for_km,
)
from core.gsm.models import (
    Anchor,
    LegPlan,
    RouteRef,
    TankState,
    Transaction,
    WaybillDay,
)

REPO_ROOT = Path(__file__).resolve().parents[1]
GSM_BALANCE_PATH = REPO_ROOT / "core" / "gsm" / "balance.py"
GSM_MODELS_PATH = REPO_ROOT / "core" / "gsm" / "models.py"


# ---------------------------------------------------------------------------
# 1. burn_for_km
# ---------------------------------------------------------------------------


def test_burn_for_km_basic() -> None:
    assert burn_for_km(100, 9.4) == 9.4
    assert burn_for_km(190, 9.4) == 17.86
    assert burn_for_km(0, 9.4) == 0.0


def test_burn_for_km_rounds_to_two_decimals() -> None:
    """Must match ``round(km * norm / 100, 2)`` (Python banker's round)."""
    assert burn_for_km(1, 9.4) == round(1 * 9.4 / 100, 2)  # 0.09
    assert burn_for_km(3, 9.4) == round(3 * 9.4 / 100, 2)  # 0.28
    assert burn_for_km(7, 10.3) == round(7 * 10.3 / 100, 2)


@pytest.mark.parametrize(
    ("km", "norm"),
    [
        (0, 0.0),
        (250, 10.3),
        (150, 9.4),
        (1, 0.01),
    ],
)
def test_burn_for_km_matches_formula(km: int, norm: float) -> None:
    assert burn_for_km(km, norm) == round(km * norm / 100, 2)


# ---------------------------------------------------------------------------
# 2. Day chain: fuel_end / odometer_end
# ---------------------------------------------------------------------------


def test_apply_day_fuel_and_odometer() -> None:
    """fuel_end = start + issued − burn; odometer_end = start + km."""
    day = date(2025, 4, 3)
    fuel_start = 20.0
    fuel_issued = 40.0
    km = 190
    norm = 9.4
    odometer_start = 100_000
    burn = burn_for_km(km, norm)

    state = apply_day(
        day,
        fuel_start=fuel_start,
        fuel_issued=fuel_issued,
        km=km,
        odometer_start=odometer_start,
        norm_per_100km=norm,
        tank_volume_liters=55.0,
    )

    assert isinstance(state, TankState)
    assert state.date == day
    assert state.fuel_start == fuel_start
    assert state.fuel_issued == fuel_issued
    assert state.km == km
    assert state.odometer_start == odometer_start
    assert state.fuel_end == pytest.approx(fuel_start + fuel_issued - burn)
    assert state.odometer_end == odometer_start + km


def test_apply_day_zero_issued_burn_only() -> None:
    """Wash / burn-off day: issued=0, only burn reduces fuel."""
    state = apply_day(
        date(2025, 4, 4),
        fuel_start=30.0,
        fuel_issued=0.0,
        km=200,
        odometer_start=50_000,
        norm_per_100km=9.4,
        tank_volume_liters=55.0,
    )
    assert state.fuel_end == pytest.approx(30.0 - burn_for_km(200, 9.4))
    assert state.odometer_end == 50_200


# ---------------------------------------------------------------------------
# 3. Invariant 0 ≤ fuel_end ≤ tank → BalanceViolation
# ---------------------------------------------------------------------------


def test_apply_day_allows_fuel_end_at_zero_and_at_tank() -> None:
    """Boundaries are inclusive: fuel_end == 0 and == tank are valid."""
    # Exactly empty after burn: start 18.8, issued 0, burn 18.8 (200 km × 9.4)
    empty = apply_day(
        date(2025, 4, 5),
        fuel_start=18.8,
        fuel_issued=0.0,
        km=200,
        odometer_start=1,
        norm_per_100km=9.4,
        tank_volume_liters=55.0,
    )
    assert empty.fuel_end == pytest.approx(0.0)

    # Exactly full: start 15, issued 40, burn 0 → 55
    full = apply_day(
        date(2025, 4, 6),
        fuel_start=15.0,
        fuel_issued=40.0,
        km=0,
        odometer_start=1,
        norm_per_100km=9.4,
        tank_volume_liters=55.0,
    )
    assert full.fuel_end == pytest.approx(55.0)


def test_apply_day_raises_on_negative_fuel_end() -> None:
    """BalanceViolation when burn would push fuel_end below 0."""
    with pytest.raises(BalanceViolation):
        apply_day(
            date(2025, 4, 7),
            fuel_start=5.0,
            fuel_issued=0.0,
            km=200,  # burn ≈ 18.8 > 5
            odometer_start=1,
            norm_per_100km=9.4,
            tank_volume_liters=55.0,
        )


def test_apply_day_raises_on_overflow_above_tank() -> None:
    """BalanceViolation when start + issued − burn exceeds tank volume."""
    assert issubclass(BalanceViolation, RuntimeError)
    with pytest.raises(BalanceViolation):
        apply_day(
            date(2025, 4, 8),
            fuel_start=40.0,
            fuel_issued=30.0,  # 70 before burn; km=0 → 70 > 55
            km=0,
            odometer_start=1,
            norm_per_100km=9.4,
            tank_volume_liters=55.0,
        )


def test_apply_day_chain_raises_mid_sequence() -> None:
    """Corridor check applies to every day in the chain."""
    with pytest.raises(BalanceViolation):
        apply_day_chain(
            [
                (date(2025, 4, 9), 0.0, 100),  # OK: 10 − 9.4 = 0.6
                (date(2025, 4, 10), 0.0, 200),  # burn 18.8 > 0.6 → violation
            ],
            fuel_start=10.0,
            odometer_start=0,
            norm_per_100km=9.4,
            tank_volume_liters=55.0,
        )


# ---------------------------------------------------------------------------
# 4. Carry fuel_end → next fuel_start; odometer continuity
# ---------------------------------------------------------------------------


def test_apply_day_chain_carries_fuel_and_odometer() -> None:
    days = [
        (date(2025, 4, 1), 40.0, 190),
        (date(2025, 4, 2), 0.0, 200),
        (date(2025, 4, 3), 35.0, 180),
    ]
    fuel_start = 12.0
    odometer_start = 80_000
    norm = 9.4
    tank = 55.0

    states = apply_day_chain(
        days,
        fuel_start=fuel_start,
        odometer_start=odometer_start,
        norm_per_100km=norm,
        tank_volume_liters=tank,
    )

    assert isinstance(states, tuple)
    assert len(states) == 3

    # Day 0 matches apply_day with the initial starts.
    first = apply_day(
        days[0][0],
        fuel_start=fuel_start,
        fuel_issued=days[0][1],
        km=days[0][2],
        odometer_start=odometer_start,
        norm_per_100km=norm,
        tank_volume_liters=tank,
    )
    assert states[0] == first

    # Continuity: each next start equals previous end.
    for prev, cur in zip(states[:-1], states[1:], strict=True):
        assert cur.fuel_start == pytest.approx(prev.fuel_end)
        assert cur.odometer_start == prev.odometer_end

    # Explicit arithmetic check on day 1.
    burn0 = burn_for_km(190, norm)
    expected_fuel1_start = fuel_start + 40.0 - burn0
    assert states[1].fuel_start == pytest.approx(expected_fuel1_start)
    assert states[1].odometer_start == odometer_start + 190


def test_apply_day_chain_empty_returns_empty_tuple() -> None:
    states = apply_day_chain(
        [],
        fuel_start=20.0,
        odometer_start=1,
        norm_per_100km=9.4,
        tank_volume_liters=55.0,
    )
    assert states == ()


def test_apply_day_chain_single_day() -> None:
    states = apply_day_chain(
        [(date(2025, 5, 1), 20.0, 150)],
        fuel_start=10.0,
        odometer_start=1000,
        norm_per_100km=9.4,
        tank_volume_liters=55.0,
    )
    assert len(states) == 1
    assert states[0].odometer_end == 1150
    assert states[0].fuel_end == pytest.approx(10.0 + 20.0 - burn_for_km(150, 9.4))


# ---------------------------------------------------------------------------
# 5. DTOs exist and are frozen
# ---------------------------------------------------------------------------


def _assert_frozen_slots_dataclass(cls: type) -> None:
    assert dataclasses.is_dataclass(cls), f"{cls.__name__} must be a dataclass"
    params = getattr(cls, "__dataclass_params__", None)
    assert params is not None and params.frozen, f"{cls.__name__} must be frozen=True"
    assert hasattr(cls, "__slots__"), f"{cls.__name__} must use slots=True"


def test_dto_classes_are_frozen_slots() -> None:
    for cls in (Transaction, Anchor, WaybillDay, TankState, RouteRef, LegPlan):
        _assert_frozen_slots_dataclass(cls)


def test_tank_state_constructible() -> None:
    state = TankState(
        date=date(2025, 4, 3),
        fuel_start=20.0,
        fuel_issued=40.0,
        fuel_end=42.14,
        km=190,
        odometer_start=100_000,
        odometer_end=100_190,
    )
    assert state.fuel_end == 42.14
    with pytest.raises(dataclasses.FrozenInstanceError):
        state.km = 0  # type: ignore[misc]


def test_route_ref_and_transaction_constructible() -> None:
    route = RouteRef(
        route_id=1,
        addr_a="Кострома",
        addr_b="Ярославль",
        km=180,
    )
    assert route.km == 180
    with pytest.raises(dataclasses.FrozenInstanceError):
        route.km = 1  # type: ignore[misc]

    tx = Transaction(
        card_id=1,
        ts=datetime(2025, 4, 3, 10, 15, 0),
        service_type="fuel",
        qty_liters=40.0,
        amount=2500.0,
        station_id=7,
        raw_address="ТАТНЕФТЬ, Кострома",
    )
    assert tx.service_type == "fuel"
    assert tx.qty_liters == 40.0


def test_anchor_and_waybill_day_constructible() -> None:
    tx = Transaction(
        card_id=1,
        ts=datetime(2025, 4, 3, 10, 15, 0),
        service_type="wash",
        qty_liters=None,
        amount=350.0,
        station_id=7,
        raw_address="мойка",
    )
    anchor = Anchor(
        date=date(2025, 4, 3),
        transactions=(tx,),
        station_ids=(7,),
    )
    assert anchor.date == date(2025, 4, 3)
    assert len(anchor.transactions) == 1

    tank = TankState(
        date=date(2025, 4, 3),
        fuel_start=20.0,
        fuel_issued=0.0,
        fuel_end=2.14,
        km=190,
        odometer_start=100_000,
        odometer_end=100_190,
    )
    route = RouteRef(route_id=2, addr_a="А", addr_b="Б", km=190)
    waybill = WaybillDay(
        date=date(2025, 4, 3),
        driver_id=1,
        route=route,
        tank=tank,
        source="auto",
        warnings=("weekend_anchor",),
    )
    assert waybill.source == "auto"
    assert waybill.warnings == ("weekend_anchor",)
    assert waybill.tank is tank
    assert waybill.legs == ()


# ---------------------------------------------------------------------------
# 6. AST guard: no app.* imports in core/gsm/{balance,models}.py
# ---------------------------------------------------------------------------


def _assert_no_app_imports(path: Path) -> None:
    assert path.is_file(), f"missing module file: {path}"
    tree = ast.parse(path.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                assert not alias.name.startswith("app"), alias.name
        elif isinstance(node, ast.ImportFrom):
            module = node.module or ""
            assert not module.startswith("app"), module


def test_gsm_balance_module_has_no_app_imports() -> None:
    _assert_no_app_imports(GSM_BALANCE_PATH)


def test_gsm_models_module_has_no_app_imports() -> None:
    _assert_no_app_imports(GSM_MODELS_PATH)


# ---------------------------------------------------------------------------
# 7. Determinism: same input → same output
# ---------------------------------------------------------------------------


def test_burn_for_km_deterministic() -> None:
    assert burn_for_km(190, 9.4) == burn_for_km(190, 9.4)


def test_apply_day_chain_deterministic() -> None:
    days = [
        (date(2025, 6, 2), 40.0, 190),
        (date(2025, 6, 3), 0.0, 210),
        (date(2025, 6, 4), 30.0, 160),
    ]
    kwargs = dict(
        fuel_start=15.0,
        odometer_start=12_000,
        norm_per_100km=10.3,
        tank_volume_liters=60.0,
    )
    a = apply_day_chain(days, **kwargs)
    b = apply_day_chain(days, **kwargs)
    assert a == b
