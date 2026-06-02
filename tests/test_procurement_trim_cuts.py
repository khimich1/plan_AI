from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.config_and_data as _cfg  # noqa: F401

from viz_modules.procurement.trim import (
    _calc_trim_components,
    format_long_cut_calculation,
    format_transverse_remainder_calculation,
)

LONG_CUT_PRICE = 460.0
TRANSVERSE_PRICE = 1200.0
BASE_1_2M_28 = 9217.0


def _cascade_plan() -> dict:
    """2×28-5,3 + 1×28-6,65 + 1×25-6,65 из визуализации."""
    return {
        "primary_cuts": [
            {"width": 530, "rest": 670, "qty": 2, "lengths": [2.8]},
            {"width": 665, "rest": 535, "qty": 1, "lengths": [2.8]},
        ],
        "secondary_cuts": [
            {
                "source": 670,
                "source_lengths": [2.8],
                "lengths": [2.5],
                "cuts": [665],
                "qty": 1,
                "pieces": 1,
                "waste": 5,
            },
            {
                "source": 535,
                "source_lengths": [2.8],
                "lengths": [2.8],
                "cuts": [530],
                "qty": 1,
                "pieces": 1,
                "waste": 5,
            },
        ],
    }


def _cascade_secondary_cuts() -> list:
    return [
        {
            "source": 670,
            "source_lengths": [2.8],
            "lengths": [2.5],
            "cuts": [665],
            "qty": 1,
            "pieces": 1,
            "waste": 5,
        },
        {
            "source": 535,
            "source_lengths": [2.8],
            "lengths": [2.8],
            "cuts": [530],
            "qty": 1,
            "pieces": 1,
            "waste": 5,
        },
    ]


def _cascade_plan_instance_rows() -> dict:
    """Cascade в формате оптимизатора: одна строка primary на instance (qty=1)."""
    return {
        "primary_cuts": [
            {"width": 530, "rest": 670, "qty": 1, "lengths": [2.8]},
            {"width": 530, "rest": 670, "qty": 1, "lengths": [2.8]},
            {"width": 665, "rest": 535, "qty": 1, "lengths": [2.8]},
        ],
        "secondary_cuts": _cascade_secondary_cuts(),
    }


def _trim(length: float, width_mm: int, qty: int, plan: dict | None = None):
    return _calc_trim_components(
        plan if plan is not None else _cascade_plan(),
        length=length,
        width_mm=width_mm,
        qty=qty,
        base_price_1_2m=BASE_1_2M_28,
        base_price=BASE_1_2M_28 * (width_mm / 1200.0),
        load_code=8,
        price_table={},
    )


def test_primary_530_two_plates() -> None:
    t = _trim(2.8, 530, qty=2)
    assert t["total_cuts_for_this_size"] == 2
    assert t["long_cut_meterage"] == pytest.approx(5.6)
    assert t["long_cut_cost"] == pytest.approx(2 * 460.0 * 2.8 / 2)
    assert t["trans_cuts"] == 0
    assert t["waste_cost"] == pytest.approx(0.0)
    assert t["rest_cost"] > 0


def test_primary_530_two_plates_instance_rows() -> None:
    plan = _cascade_plan_instance_rows()
    t = _trim(2.8, 530, qty=2, plan=plan)
    assert t["total_cuts_for_this_size"] == 2
    assert t["long_cut_meterage"] == pytest.approx(5.6)
    assert t["long_cut_cost"] == pytest.approx(2 * 460.0 * 2.8 / 2)
    assert t["trans_cuts"] == 0
    assert t["waste_cost"] == pytest.approx(0.0)
    assert t["rest_cost"] > 0
    assert format_long_cut_calculation(t, 2) == "460 × 2,8 × 2 / 2"


def test_primary_665_one_plate() -> None:
    t = _trim(2.8, 665, qty=1)
    assert t["total_cuts_for_this_size"] == 1
    assert t["long_cut_meterage"] == pytest.approx(2.8)
    assert t["long_cut_cost"] == pytest.approx(460.0 * 2.8)
    assert t["trans_cuts"] == 0
    assert t["waste_cost"] == pytest.approx(0.0)


def test_secondary_665_transverse_at_25() -> None:
    t = _trim(2.5, 665, qty=1)
    assert t["total_cuts_for_this_size"] == 0
    assert t["long_cut_meterage"] == pytest.approx(0.0)
    assert t["long_cut_cost"] == pytest.approx(0.0)
    assert t["trans_cuts"] == pytest.approx(1.0)
    expected_trans_rem = BASE_1_2M_28 * (665 / 1200.0) * (0.3 / 2.5)
    assert t["transverse_remainder_cost"] == pytest.approx(expected_trans_rem)
    assert t["waste_cost"] == pytest.approx(0.0)


def test_primary_720_transverse_at_206() -> None:
    """ПБ 20,6-7,2: primary 720 + поперечный рез 2,54→2,06 на основной полосе."""
    base_1_2m = 7105.0
    plan = {
        "primary_cuts": [{"width": 720, "rest": 480, "qty": 3, "lengths": [2.06]}],
        "secondary_cuts": [
            {
                "source": 720,
                "source_lengths": [2.54],
                "lengths": [2.06],
                "cuts": [720],
                "qty": 3,
                "pieces": 1,
                "type": "transverse",
            },
        ],
    }
    t = _calc_trim_components(
        plan,
        length=2.06,
        width_mm=720,
        qty=3,
        base_price_1_2m=base_1_2m,
        base_price=base_1_2m * (720 / 1200.0),
        load_code=8,
        price_table={},
    )
    assert t["primary_matched"] is True
    assert t["trans_cuts"] == pytest.approx(1.0)
    expected = base_1_2m * (720 / 1200.0) * (0.48 / 2.06)
    assert t["transverse_remainder_cost"] == pytest.approx(expected)
    assert t["transverse_remainder_terms"] == [(0.48, 3)]


def test_transverse_remainder_breakdown_label() -> None:
    t = _trim(2.5, 665, qty=3)
    label = format_transverse_remainder_calculation(
        t,
        3,
        base_price_1_2m=BASE_1_2M_28,
        width_m=665 / 1000.0,
        length_m=2.5,
    )
    assert label is not None
    assert "9 217,00" in label
    assert "(0,67 / 1,2)" in label
    assert "(0,30 / 2,50)" in label
    assert "/ 3" in label


def test_secondary_530_from_remainder() -> None:
    """530 из вторичного реза (665 primary на другой плите) — только secondary path."""
    plan = {
        "primary_cuts": [{"width": 665, "rest": 535, "qty": 1, "lengths": [2.8]}],
        "secondary_cuts": [
            {
                "source": 535,
                "source_lengths": [2.8],
                "lengths": [2.8],
                "cuts": [530],
                "qty": 1,
                "pieces": 1,
                "waste": 5,
            },
        ],
    }
    t = _trim(2.8, 530, qty=1, plan=plan)
    assert t["primary_matched"] is False
    assert t["total_cuts_for_this_size"] == 0
    assert t["long_cut_meterage"] == pytest.approx(0.0)
    assert t["waste_cost"] == pytest.approx(0.0)


def test_same_width_cascade_530_qty2() -> None:
    """ПБ 60-5,3 × 2: primary 530 + secondary 670→530+140 — оба реза и отход в одной строке."""
    plan = {
        "primary_cuts": [{"width": 530, "rest": 670, "qty": 1, "lengths": [6.0]}],
        "secondary_cuts": [
            {
                "source": 670,
                "source_lengths": [6.0],
                "lengths": [6.0],
                "cuts": [530],
                "qty": 1,
                "pieces": 1,
                "waste": 140,
            },
        ],
    }
    t = _trim(6.0, 530, qty=2, plan=plan)
    assert t["primary_matched"] is True
    assert t["rest_used"] is True
    assert t["total_cuts_for_this_size"] == 2
    assert t["long_cut_meterage"] == pytest.approx(12.0)
    assert t["long_cut_cost"] == pytest.approx(2 * LONG_CUT_PRICE * 6.0 / 2)
    assert t["waste_cost"] == pytest.approx((140 / 1200.0) * BASE_1_2M_28 / 2)
    assert format_long_cut_calculation(t, 2) == "460 × 6,0 × 2 / 2"
    assert t["waste_terms"] == [(140, 1)]


def test_same_width_cascade_breakdown_labels() -> None:
    """Строки разбивки КП: «Продольный рез» и «Отходы» для merged same-width."""
    plan = {
        "primary_cuts": [{"width": 530, "rest": 670, "qty": 1, "lengths": [6.0]}],
        "secondary_cuts": [
            {
                "source": 670,
                "source_lengths": [6.0],
                "lengths": [6.0],
                "cuts": [530],
                "qty": 1,
                "pieces": 1,
                "waste": 140,
            },
        ],
    }
    t = _trim(6.0, 530, qty=2, plan=plan)
    long_label = format_long_cut_calculation(t, 2)
    assert long_label is not None
    assert "× 2 / 2" in long_label
    waste_parts = []
    for w_mm, n in t["waste_terms"]:
        waste_parts.append(f"{int(w_mm)}×{n}" if n > 1 else str(int(w_mm)))
    assert waste_parts == ["140"]
    assert t["waste_cost"] > 0


def test_same_width_cascade_no_regression_primary_qty2() -> None:
    """primary 530×2: secondary 535→530 не привязывается (чужой source)."""
    plan = {
        "primary_cuts": [
            {"width": 530, "rest": 670, "qty": 2, "lengths": [2.8]},
            {"width": 665, "rest": 535, "qty": 1, "lengths": [2.8]},
        ],
        "secondary_cuts": [
            {
                "source": 535,
                "source_lengths": [2.8],
                "lengths": [2.8],
                "cuts": [530],
                "qty": 1,
                "pieces": 1,
                "waste": 5,
            },
        ],
    }
    t = _trim(2.8, 530, qty=2, plan=plan)
    assert t["total_cuts_for_this_size"] == 2
    assert t["long_cut_meterage"] == pytest.approx(5.6)
    assert t["waste_cost"] == pytest.approx(0.0)


def test_same_width_cascade_qty3_primary2_secondary1() -> None:
    """order qty=3: primary=2, secondary same-width=1."""
    plan = {
        "primary_cuts": [{"width": 530, "rest": 670, "qty": 2, "lengths": [6.0]}],
        "secondary_cuts": [
            {
                "source": 670,
                "source_lengths": [6.0],
                "lengths": [6.0],
                "cuts": [530],
                "qty": 1,
                "pieces": 1,
                "waste": 140,
            },
        ],
    }
    t = _trim(6.0, 530, qty=3, plan=plan)
    assert t["total_cuts_for_this_size"] == 3
    assert t["long_cut_meterage"] == pytest.approx(18.0)
    assert t["waste_cost"] == pytest.approx((140 / 1200.0) * BASE_1_2M_28 / 3)


def test_no_double_waste_between_primary_and_secondary_pair() -> None:
    """530 primary и 665 secondary из одной связки 670 — отход только у secondary."""
    plan = {
        "primary_cuts": [{"width": 530, "rest": 670, "qty": 1, "lengths": [2.8]}],
        "secondary_cuts": [
            {
                "source": 670,
                "source_lengths": [2.8],
                "lengths": [2.8],
                "cuts": [665],
                "qty": 1,
                "pieces": 1,
                "waste": 5,
            },
        ],
    }
    t530 = _trim(2.8, 530, qty=1, plan=plan)
    t665 = _trim(2.8, 665, qty=1, plan=plan)
    assert t530["waste_cost"] == pytest.approx(0.0)
    assert t665["waste_cost"] == pytest.approx(0.0)


def test_order_total_cut_costs() -> None:
    plan = _cascade_plan()
    rows = [
        (2.8, 530, 2),
        (2.8, 665, 1),
        (2.5, 665, 1),
    ]
    total_long = sum(_trim(l, w, q, plan)["long_cut_cost"] * q for l, w, q in rows)
    total_trans = sum(
        _trim(l, w, q, plan)["trans_cuts"] * TRANSVERSE_PRICE for l, w, q in rows
    )
    assert total_long == pytest.approx(3 * 460.0 * 2.8)
    assert total_trans == pytest.approx(TRANSVERSE_PRICE)


def test_order_total_cut_costs_instance_rows() -> None:
    plan = _cascade_plan_instance_rows()
    rows = [
        (2.8, 530, 2),
        (2.8, 665, 1),
        (2.5, 665, 1),
    ]
    total_long = sum(_trim(l, w, q, plan)["long_cut_cost"] * q for l, w, q in rows)
    total_trans = sum(
        _trim(l, w, q, plan)["trans_cuts"] * TRANSVERSE_PRICE for l, w, q in rows
    )
    assert total_long == pytest.approx(3 * 460.0 * 2.8)
    assert total_trans == pytest.approx(TRANSVERSE_PRICE)


def test_plan_exists_no_match_no_fallback_cuts() -> None:
    from viz_modules.procurement.trim import resolve_long_cut_pricing

    plan = _cascade_plan()
    trim = _trim(3.0, 320, qty=1, plan=plan)
    cost, cuts, _ = resolve_long_cut_pricing(
        trim,
        qty=1,
        length=3.0,
        width_m=0.32,
        current_plan=plan,
        fallback_long_cuts=1,
    )
    assert cost == 0.0
    assert cuts == 0


BASE_1_2M_73 = 27574.0


def _apply_factory(**kwargs):
    from viz_modules.procurement.trim import apply_factory_strip_waste

    defaults = dict(
        width_mm=1020,
        base_price_1_2m=BASE_1_2M_73,
        rest_cost=0.0,
        rest_used=False,
        waste_cost=0.0,
        waste_terms=[],
        qty=6,
    )
    defaults.update(kwargs)
    return apply_factory_strip_waste(**defaults)


def test_factory_waste_skipped_when_rest_cost_positive() -> None:
    rest_cost = (180 / 1200.0) * BASE_1_2M_73
    waste_cost, waste_terms = _apply_factory(rest_cost=rest_cost)
    assert waste_cost == pytest.approx(0.0)
    assert waste_terms == []


def test_factory_waste_skipped_when_rest_used() -> None:
    waste_cost, waste_terms = _apply_factory(rest_used=True)
    assert waste_cost == pytest.approx(0.0)
    assert waste_terms == []


def test_factory_waste_applied_without_plan() -> None:
    expected = (180 / 1200.0) * BASE_1_2M_73
    waste_cost, waste_terms = _apply_factory()
    assert waste_cost == pytest.approx(expected)
    assert waste_terms == [(180, 6)]


def test_factory_waste_partial_rest_not_doubled() -> None:
    partial_rest_cost = (540 / 1200.0) * BASE_1_2M_73 / 6
    waste_cost, waste_terms = _apply_factory(rest_cost=partial_rest_cost)
    assert waste_cost == pytest.approx(0.0)
    assert waste_terms == []


def test_factory_waste_not_applied_outside_range() -> None:
    waste_cost, waste_terms = _apply_factory(width_mm=530)
    assert waste_cost == pytest.approx(0.0)
    assert waste_terms == []


def test_1020_primary_unused_rest_no_factory_waste_double_count() -> None:
    """6× ПБ 73-10,2: rest 180мм × 6 не использован — только rest_cost, без factory waste."""
    plan = {
        "primary_cuts": [
            {"width": 1020, "rest": 180, "qty": 6, "lengths": [7.3]},
        ],
        "secondary_cuts": [],
    }
    t = _calc_trim_components(
        plan,
        length=7.3,
        width_mm=1020,
        qty=6,
        base_price_1_2m=BASE_1_2M_73,
        base_price=BASE_1_2M_73 * (1.02 / 1.2),
        load_code=8,
        price_table={},
    )
    per_plate_rest = (180 / 1200.0) * BASE_1_2M_73
    assert t["rest_cost"] == pytest.approx(per_plate_rest)
    assert t["rest_used"] is False

    waste_cost, waste_terms = _apply_factory(
        rest_cost=t["rest_cost"],
        rest_used=t["rest_used"],
    )
    assert waste_cost == pytest.approx(0.0)
    assert waste_terms == []
