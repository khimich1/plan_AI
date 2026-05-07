from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


def test_orders_must_be_dict_not_list() -> None:
    from core.optimization.orchestrator import optimize_with_cascading_longitudinal_cuts

    with pytest.raises(ValueError, match="orders must be a dict"):
        optimize_with_cascading_longitudinal_cuts(orders=[])


def test_orders_2d_must_be_list() -> None:
    from core.optimization.orchestrator import optimize_with_cascading_longitudinal_cuts

    with pytest.raises(ValueError, match="orders_2d must be a list"):
        optimize_with_cascading_longitudinal_cuts(orders_2d={})


def test_orders_2d_row_requires_keys() -> None:
    from core.optimization.orchestrator import optimize_with_cascading_longitudinal_cuts

    with pytest.raises(ValueError, match="missing required key"):
        optimize_with_cascading_longitudinal_cuts(orders_2d=[{"length": 5.0, "width": 320}])


def test_redact_order_strips_sensitive_keys() -> None:
    from core.optimization.logging_utils import redact_order

    d = redact_order({"qty": 2, "length": 5.6, "width": 320, "customer": "Секрет"})
    assert "customer" not in d
    assert d["qty"] == 2


def test_order_line_for_console_never_includes_customer_in_shape() -> None:
    """Строка консоли основана только на qty/length/width — без имён."""
    from core.optimization.logging_utils import order_line_for_console

    s = order_line_for_console({"qty": 2, "length": 5.6, "width": 320, "customer": "Секрет"})
    assert "Секрет" not in s
    assert "5.6" in s
