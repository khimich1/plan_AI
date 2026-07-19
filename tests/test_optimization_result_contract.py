from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.optimization.orchestrator import optimize_with_cascading_longitudinal_cuts
from core.optimization.result_contract import (
    ERROR_EMPTY_ORDERS_2D,
    ERROR_NO_INPUT,
    is_optimization_success,
    opt_error,
    opt_ok,
)


def test_no_input_returns_structured_error() -> None:
    r = optimize_with_cascading_longitudinal_cuts()
    assert r.get("_opt_status") == "error"
    assert r.get("_opt_error_code") == ERROR_NO_INPUT
    assert not is_optimization_success(r)


def test_opt_error_and_ok_helpers() -> None:
    err = opt_error("test_code", "human")
    assert err["_opt_status"] == "error"
    assert err["_opt_error_code"] == "test_code"
    assert not is_optimization_success(err)

    ok = opt_ok({"total_plates": 1, "primary_cuts": []})
    assert ok["_opt_status"] == "ok"
    assert is_optimization_success(ok)


def test_empty_orders_2d_from_service_path_is_error() -> None:
    """Пустой 2D вход — не путать с успехом (truthy dict)."""
    assert not is_optimization_success(
        opt_error(ERROR_EMPTY_ORDERS_2D, "empty"),
    )
