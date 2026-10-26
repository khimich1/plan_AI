"""Закупка читает заказ из аргумента или из уже открытого bound()."""

from __future__ import annotations

import threading

import pytest

from core.plate_order_context import PlateOrderContext
from core.plate_runtime_state import new_plate_mutable_runtime_empty
from viz_modules.procurement import (
    build_component_breakdown,
    build_component_breakdown_production,
    build_price_rows,
    build_price_rows_production,
    build_procurement_items,
)
from viz_modules.procurement.ports import ProcurementDeps

_STUB_DEPS = ProcurementDeps(
    db_path=":memory:",
    get_price=lambda length_m, load_code, db_path: 1000.0,
    get_raw_material_cost=lambda plate_name, db_path: 1000.0,
    get_reinforcement=lambda length_m, load_code, source="erm", db_path=":memory:", allow_fallback=True: 5.0,
)


def _on_fresh_thread(fn) -> None:
    thread = threading.Thread(target=fn)
    thread.start()
    thread.join()
    assert not thread.is_alive()


def test_procurement_items_explicit_runtime_in_raise_mode(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PLATE_RUNTIME_STRICT", "raise")
    outcome: dict[str, object] = {}

    def _worker() -> None:
        runtime = new_plate_mutable_runtime_empty()
        runtime.plate_load_details[(6.6, 1.2, 8, "")] = 2
        outcome["items"] = build_procurement_items(runtime=runtime)
        outcome["has_tls"] = False

    _on_fresh_thread(_worker)
    items = outcome["items"]
    assert isinstance(items, list)
    assert len(items) == 1
    assert items[0]["qty"] == 2
    assert items[0]["load_code"] == 8
    assert items[0]["length"] == pytest.approx(6.6)


def test_procurement_items_without_runtime_or_bound_raises_in_raise_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("PLATE_RUNTIME_STRICT", "raise")
    outcome: dict[str, BaseException | None] = {"error": None}

    def _worker() -> None:
        try:
            build_procurement_items()
        except RuntimeError as exc:
            outcome["error"] = exc

    _on_fresh_thread(_worker)
    error = outcome["error"]
    assert isinstance(error, RuntimeError)
    assert "заказ плит не привязан" in str(error)


def test_procurement_items_bound_scope_is_silent_in_raise_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("PLATE_RUNTIME_STRICT", "raise")
    ctx = PlateOrderContext.fresh_empty()
    ctx.plates.plate_load_details[(5.5, 1.2, 10, "")] = 4
    with ctx.bound():
        items = build_procurement_items()
    assert len(items) == 1
    assert items[0]["load_code"] == 10
    assert items[0]["qty"] == 4


def test_price_rows_and_breakdown_accept_explicit_runtime(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("PLATE_RUNTIME_STRICT", "raise")
    outcome: dict[str, object] = {}

    def _worker() -> None:
        runtime = new_plate_mutable_runtime_empty()
        runtime.plate_load_details[(7.3, 1.2, 12, "73")] = 1
        rows, _total = build_price_rows({}, deps=_STUB_DEPS, runtime=runtime)
        rows_prod, _total_prod = build_price_rows_production(
            {},
            deps=_STUB_DEPS,
            runtime=runtime,
        )
        breakdown = build_component_breakdown({}, deps=_STUB_DEPS, runtime=runtime)
        breakdown_prod = build_component_breakdown_production(
            {},
            deps=_STUB_DEPS,
            runtime=runtime,
        )
        outcome["rows"] = rows
        outcome["rows_prod"] = rows_prod
        outcome["breakdown"] = breakdown
        outcome["breakdown_prod"] = breakdown_prod

    _on_fresh_thread(_worker)
    rows = outcome["rows"]
    assert isinstance(rows, list) and len(rows) == 1
    assert "12" in str(rows[0][1])
    assert isinstance(outcome["rows_prod"], list) and len(outcome["rows_prod"]) == 1
    assert isinstance(outcome["breakdown"], list)
    assert isinstance(outcome["breakdown_prod"], list)
