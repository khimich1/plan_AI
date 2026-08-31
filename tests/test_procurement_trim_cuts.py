from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.config_and_data as _cfg  # noqa: F401

from viz_modules.procurement.trim import (
    _calc_trim_components,
    _is_crossload_rest_secondary,
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


def test_plan_exists_unmatched_width_uses_fallback_long_cut() -> None:
    """BUG-4: при плане без match trim использует fallback_long_cuts."""
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
    assert cuts == 1
    assert cost == pytest.approx(LONG_CUT_PRICE * 3.0)


BASE_1_2M_63 = 21515.0
BASE_1_2M_45 = 14510.0
BASE_1_2M_108 = 23468.0


def _case3_cross_width_plan() -> dict:
    """ПБ 63-7,2-8п (720) + ПБ 45-3,2-8п (320) с rest 480 → отход 160 только на primary."""
    return {
        "primary_cuts": [{"width": 720, "rest": 480, "qty": 1, "lengths": [6.3]}],
        "secondary_cuts": [
            {
                "source": 480,
                "source_lengths": [6.3],
                "lengths": [6.3],
                "cuts": [320],
                "qty": 1,
                "pieces": 1,
                "waste": 160,
            },
        ],
    }


def test_waste_160_only_on_primary_720_320() -> None:
    """BUG-2 кейс 3: отход 160 мм только на ПБ 63-7,2, не на ПБ 45-3,2."""
    plan = _case3_cross_width_plan()
    t720 = _calc_trim_components(
        plan,
        length=6.3,
        width_mm=720,
        qty=1,
        base_price_1_2m=BASE_1_2M_63,
        base_price=BASE_1_2M_63 * (720 / 1200.0),
        load_code=8,
        price_table={},
    )
    t320 = _calc_trim_components(
        plan,
        length=6.3,
        width_mm=320,
        qty=1,
        base_price_1_2m=BASE_1_2M_45,
        base_price=BASE_1_2M_45 * (320 / 1200.0),
        load_code=8,
        price_table={},
    )
    expected_primary_waste = (160 / 1200.0) * BASE_1_2M_63
    assert t720["waste_cost"] == pytest.approx(expected_primary_waste)
    assert t320["waste_cost"] == pytest.approx(0.0)


def test_secondary_from_rest_no_strip_waste() -> None:
    """BUG-2: secondary с полосы primary rest не тарифицирует отход полосы."""
    plan = {
        "primary_cuts": [{"width": 640, "rest": 560, "qty": 1, "lengths": [5.0]}],
        "secondary_cuts": [
            {
                "source": 560,
                "source_lengths": [5.0],
                "lengths": [5.0],
                "cuts": [320],
                "qty": 1,
                "pieces": 1,
                "waste": 240,
            },
        ],
    }
    t320 = _calc_trim_components(
        plan,
        length=5.0,
        width_mm=320,
        qty=1,
        base_price_1_2m=BASE_1_2M_28,
        base_price=BASE_1_2M_28 * (320 / 1200.0),
        load_code=8,
        price_table={},
    )
    assert t320["primary_matched"] is False
    assert t320["waste_cost"] == pytest.approx(0.0)


def _pb_422_665_321_530_plan() -> dict:
    """ПБ 42,2-6,65-8п (primary) + ПБ 32,1-5,3-8п (secondary с rest + поперечный)."""
    return {
        "primary_cuts": [{"width": 665, "rest": 535, "qty": 1, "lengths": [4.2]}],
        "secondary_cuts": [
            {
                "source": 535,
                "source_lengths": [4.2],
                "lengths": [3.2],
                "cuts": [530],
                "qty": 1,
                "pieces": 1,
                "waste": 5,
            },
        ],
    }


def test_pb_422_665_321_530_no_double_long_cut() -> None:
    """Каскад 665→530: один физический продольный рез — только на primary."""
    from viz_modules.procurement.trim import resolve_long_cut_pricing

    plan = _pb_422_665_321_530_plan()
    t665 = _calc_trim_components(
        plan,
        length=4.2,
        width_mm=665,
        qty=1,
        base_price_1_2m=BASE_1_2M_28,
        base_price=BASE_1_2M_28 * (665 / 1200.0),
        load_code=8,
        price_table={},
    )
    t530 = _calc_trim_components(
        plan,
        length=3.2,
        width_mm=530,
        qty=1,
        base_price_1_2m=BASE_1_2M_28,
        base_price=BASE_1_2M_28 * (530 / 1200.0),
        load_code=8,
        price_table={},
    )

    assert t665["primary_matched"] is True
    assert t665["long_cut_meterage"] == pytest.approx(4.2)
    assert t530["primary_matched"] is False
    assert t530["long_cut_meterage"] == pytest.approx(0.0)
    assert t530["trans_cuts"] == pytest.approx(1.0)

    _, long_cuts_665, _ = resolve_long_cut_pricing(
        t665,
        qty=1,
        length=4.2,
        width_m=0.665,
        current_plan=plan,
        fallback_long_cuts=0,
    )
    long_cost_530, long_cuts_530, _ = resolve_long_cut_pricing(
        t530,
        qty=1,
        length=3.2,
        width_m=0.53,
        current_plan=plan,
        fallback_long_cuts=0,
    )

    assert long_cuts_665 == 1
    assert long_cuts_530 == 0
    assert long_cost_530 == pytest.approx(0.0)

    total_meterage = t665["long_cut_meterage"] + t530["long_cut_meterage"]
    assert total_meterage == pytest.approx(4.2)


def test_crossload_secondary_long_cut_in_trim_not_fallback() -> None:
    """Кросс-нагрузка 10п→8п: продольный рез в trim, не через fallback."""
    from viz_modules.procurement.trim import resolve_long_cut_pricing

    plan = _crossload_plan_10p_rest_535()
    t665 = _calc_trim_components(
        plan,
        length=2.8,
        width_mm=665,
        qty=1,
        base_price_1_2m=BASE_1_2M_28,
        base_price=BASE_1_2M_28 * (665 / 1200.0),
        load_code=8,
        price_table={},
    )
    assert t665["primary_matched"] is False
    assert t665["long_cut_meterage"] == pytest.approx(2.8)

    long_cost, long_cuts, _ = resolve_long_cut_pricing(
        t665,
        qty=1,
        length=2.8,
        width_m=0.665,
        current_plan=plan,
        fallback_long_cuts=0,
    )
    assert long_cuts == 1
    assert long_cost == pytest.approx(LONG_CUT_PRICE * 2.8)


def _case4_multi_secondary_plan() -> dict:
    """BUG-3 кейс 4: rest 880, два secondary 320, waste 240 только в данных оптимизатора."""
    return {
        "primary_cuts": [{"width": 320, "rest": 880, "qty": 1, "lengths": [5.63]}],
        "secondary_cuts": [
            {
                "source": 880,
                "source_lengths": [5.63],
                "lengths": [5.63],
                "cuts": [320],
                "qty": 1,
                "pieces": 1,
                "waste": 0,
            },
            {
                "source": 880,
                "source_lengths": [5.63],
                "lengths": [4.64],
                "cuts": [320],
                "qty": 1,
                "pieces": 1,
                "waste": 240,
            },
        ],
    }


def test_waste_560_primary_secondary_240_zero() -> None:
    """BUG-3 кейс 4: secondary с waste=240 в плане не получает строку отхода."""
    plan = _case4_multi_secondary_plan()
    t_primary = _calc_trim_components(
        plan,
        length=5.63,
        width_mm=320,
        qty=1,
        base_price_1_2m=BASE_1_2M_28,
        base_price=BASE_1_2M_28 * (320 / 1200.0),
        load_code=8,
        price_table={},
    )
    t_secondary = _calc_trim_components(
        plan,
        length=4.64,
        width_mm=320,
        qty=1,
        base_price_1_2m=BASE_1_2M_28,
        base_price=BASE_1_2M_28 * (320 / 1200.0),
        load_code=8,
        price_table={},
    )
    assert t_primary["waste_cost"] > 0
    assert t_secondary["waste_cost"] == pytest.approx(0.0)


def test_1080_factory_strip_includes_long_cut() -> None:
    """BUG-1 кейсы 1–2: 1080 мм — factory strip + продольный рез."""
    from viz_modules.procurement.trim import apply_factory_strip_waste, resolve_long_cut_pricing

    plan = {
        "primary_cuts": [{"width": 530, "rest": 670, "qty": 1, "lengths": [3.05]}],
        "secondary_cuts": [],
    }
    length = 3.05
    width_mm = 1080
    trim = _calc_trim_components(
        plan,
        length=length,
        width_mm=width_mm,
        qty=1,
        base_price_1_2m=BASE_1_2M_108,
        base_price=BASE_1_2M_108 * (1.08 / 1.2),
        load_code=12,
        price_table={},
    )
    waste_cost, waste_terms = apply_factory_strip_waste(
        width_mm=width_mm,
        base_price_1_2m=BASE_1_2M_108,
        rest_cost=trim["rest_cost"],
        rest_used=trim["rest_used"],
        waste_cost=trim["waste_cost"],
        waste_terms=trim["waste_terms"],
        qty=1,
    )
    long_cut_cost, long_cuts, _ = resolve_long_cut_pricing(
        trim,
        qty=1,
        length=length,
        width_m=1.08,
        current_plan=plan,
        fallback_long_cuts=0,
    )
    assert long_cuts == 1
    assert long_cut_cost == pytest.approx(LONG_CUT_PRICE * length)
    assert waste_cost == pytest.approx((120 / 1200.0) * BASE_1_2M_108)
    assert waste_terms == [(120, 1)]


def test_1080_qty2_waste_formula_120x2() -> None:
    """BUG-1: qty=2 — waste (120×2) и продольный рез на каждую плиту."""
    from viz_modules.procurement.trim import apply_factory_strip_waste, resolve_long_cut_pricing

    length = 3.21
    width_mm = 1080
    trim = _calc_trim_components(
        None,
        length=length,
        width_mm=width_mm,
        qty=2,
        base_price_1_2m=BASE_1_2M_108,
        base_price=BASE_1_2M_108 * (1.08 / 1.2),
        load_code=12,
        price_table={},
    )
    waste_cost, waste_terms = apply_factory_strip_waste(
        width_mm=width_mm,
        base_price_1_2m=BASE_1_2M_108,
        rest_cost=trim["rest_cost"],
        rest_used=trim["rest_used"],
        waste_cost=trim["waste_cost"],
        waste_terms=trim["waste_terms"],
        qty=2,
    )
    long_cut_cost, _, _ = resolve_long_cut_pricing(
        trim,
        qty=2,
        length=length,
        width_m=1.08,
        current_plan=None,
        fallback_long_cuts=1,
    )
    assert waste_terms == [(120, 2)]
    assert waste_cost == pytest.approx((120 / 1200.0) * BASE_1_2M_108)
    assert long_cut_cost == pytest.approx(LONG_CUT_PRICE * length)


BASE_1_2M_66_8 = 22821.0


def _pb_66_108_plan() -> dict:
    """ПБ 66-10,8-8п: primary 1080 мм, rest=0 (factory strip из 1200)."""
    return {
        "orders_requested": [
            {"length": 6.62, "width": 1080, "qty": 1, "load_code": 8},
        ],
        "primary_cuts": [
            {
                "width": 1080,
                "rest": 0,
                "qty": 1,
                "lengths": [6.62],
                "load_code": 8,
            },
        ],
        "secondary_cuts": [],
    }


def test_1080_plan_rest_zero_includes_long_cut() -> None:
    """Регрессия: план с rest=0 матчит плиту, но продольный рез всё равно нужен."""
    from viz_modules.procurement.trim import apply_factory_strip_waste, resolve_long_cut_pricing

    plan = _pb_66_108_plan()
    length = 6.62
    width_mm = 1080
    trim = _calc_trim_components(
        plan,
        length=length,
        width_mm=width_mm,
        qty=1,
        base_price_1_2m=BASE_1_2M_66_8,
        base_price=BASE_1_2M_66_8 * (1.08 / 1.2),
        load_code=8,
        price_table={},
    )
    assert trim["total_plates_from_cuts"] == 1
    assert trim["long_cut_meterage"] == pytest.approx(0.0)
    assert trim["long_cut_cost"] == pytest.approx(0.0)

    long_cut_cost, long_cuts, _ = resolve_long_cut_pricing(
        trim,
        qty=1,
        length=length,
        width_m=1.08,
        current_plan=plan,
        fallback_long_cuts=0,
    )
    assert long_cuts == 1
    assert long_cut_cost == pytest.approx(LONG_CUT_PRICE * length)

    waste_cost, waste_terms = apply_factory_strip_waste(
        width_mm=width_mm,
        base_price_1_2m=BASE_1_2M_66_8,
        rest_cost=trim["rest_cost"],
        rest_used=trim["rest_used"],
        waste_cost=trim["waste_cost"],
        waste_terms=trim["waste_terms"],
        qty=1,
    )
    assert waste_terms == [(120, 1)]
    assert waste_cost == pytest.approx((120 / 1200.0) * BASE_1_2M_66_8)

    base_price = BASE_1_2M_66_8 * (1.08 / 1.2)
    total_without_long_cut = base_price + waste_cost
    assert total_without_long_cut == pytest.approx(BASE_1_2M_66_8)
    assert base_price + waste_cost + long_cut_cost > BASE_1_2M_66_8


def test_pb_66_108_breakdown_includes_long_cut_line(monkeypatch) -> None:
    """ПБ 66-10,8-8п: в разбивке есть строка «Продольный рез»."""
    import core.optimization as optimization
    from viz_modules.procurement.breakdown import build_component_breakdown
    from viz_modules.procurement.ports import ProcurementDeps

    plan = _pb_66_108_plan()
    monkeypatch.setattr(optimization, "OPT_CASCADING_PLAN_BY_LOAD", {8: plan})
    monkeypatch.setattr(optimization, "LOAD_TO_REINFORCEMENT_MAP", {8: [8]})
    monkeypatch.setattr(optimization, "OPT_CASCADING_PLAN", {})

    deps = ProcurementDeps(
        db_path=":memory:",
        get_price=lambda length_m, load_code, db_path: BASE_1_2M_66_8,
        get_raw_material_cost=lambda plate_name, db_path: BASE_1_2M_66_8,
        get_reinforcement=lambda length_m, load_code, source="erm", db_path=":memory:", allow_fallback=True: 20.0,
    )

    tables = build_component_breakdown({}, deps=deps, reinforcement_code=8)
    assert len(tables) == 1
    row_labels = [row[0] for row in tables[0]["rows"]]
    assert "Продольный рез" in row_labels
    assert any("Отходы" in label for label in row_labels)

    long_row = next(row for row in tables[0]["rows"] if row[0] == "Продольный рез")
    expected_long = LONG_CUT_PRICE * 6.62
    cost_str = long_row[2].replace(" руб", "").replace(" ", "").replace(",", ".")
    assert float(cost_str) == pytest.approx(expected_long)

    total_row = next(row for row in tables[0]["rows"] if row[0] == "ИТОГО за 1 плиту")
    total_str = total_row[2].replace(" руб", "").replace(" ", "").replace(",", ".")
    total_per_unit = float(total_str)
    assert total_per_unit == pytest.approx(BASE_1_2M_66_8 + expected_long)
    assert total_per_unit > BASE_1_2M_66_8


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


BASE_1_2M_69_8 = 23468.0
BASE_1_2M_69_10 = 24903.0


def _mixed_load_plan_69_395() -> dict:
    """План оптимизатора: слэб 8п (3 плиты, 2 реза на полосе) + слэб 10п (10п+8п, отход 410)."""
    return {
        "primary_cuts": [
            {
                "width": 395,
                "rest": 805,
                "qty": 1,
                "lengths": [6.9],
                "load_code": 8,
                "primary_instance_id": "prim-1",
            },
            {
                "width": 395,
                "rest": 805,
                "qty": 1,
                "lengths": [6.9],
                "load_code": 10,
                "primary_instance_id": "prim-2",
            },
        ],
        "secondary_cuts": [
            {
                "source": 805,
                "source_lengths": [6.9],
                "lengths": [6.9],
                "cuts": [395],
                "qty": 1,
                "pieces": 1,
                "waste": 15,
                "load_code": 8,
                "parent_instance_id": "prim-1",
            },
            {
                "source": 805,
                "source_lengths": [6.9],
                "lengths": [6.9],
                "cuts": [395],
                "qty": 1,
                "pieces": 1,
                "waste": 15,
                "load_code": 8,
                "parent_instance_id": "prim-1",
            },
            {
                "source": 805,
                "source_lengths": [6.9],
                "lengths": [6.9],
                "cuts": [395],
                "qty": 1,
                "pieces": 1,
                "waste": 15,
                "load_code": 8,
                "parent_instance_id": "prim-2",
            },
        ],
    }


def test_mixed_load_10p_strip_waste_410() -> None:
    """10п: отход 410мм на владельце слэба (805−395), не полный остаток 805."""
    plan = _mixed_load_plan_69_395()
    t = _calc_trim_components(
        plan,
        length=6.9,
        width_mm=395,
        qty=1,
        base_price_1_2m=BASE_1_2M_69_10,
        base_price=BASE_1_2M_69_10 * (0.395 / 1.2),
        load_code=10,
        price_table={},
    )
    expected_waste = (410 / 1200.0) * BASE_1_2M_69_10
    assert t["waste_cost"] == pytest.approx(expected_waste)
    assert t["rest_cost"] == pytest.approx(0.0)
    assert any(w == 410 for w, _ in t["waste_terms"])
    assert t["total_cuts_for_this_size"] == 1


def test_mixed_load_8p_three_cuts_no_410_waste() -> None:
    """8п: 2 реза на полосе 8п + 1 кросс-рез с слэба 10п; отход 410 только на 10п."""
    plan = _mixed_load_plan_69_395()
    t = _calc_trim_components(
        plan,
        length=6.9,
        width_mm=395,
        qty=4,
        base_price_1_2m=BASE_1_2M_69_8,
        base_price=BASE_1_2M_69_8 * (0.395 / 1.2),
        load_code=8,
        price_table={},
    )
    assert t["rest_used"] is True
    assert t["rest_cost"] == pytest.approx(0.0)
    assert t["waste_cost"] == pytest.approx(0.0)
    assert t["total_cuts_for_this_size"] == 3
    assert not any(w == 410 for w, _ in t["waste_terms"])


def _crossload_plan_10p_rest_535() -> dict:
    """8п secondary с полосы остатка 10п primary (rest 535 ≠ rest 8п 670)."""
    return {
        "primary_cuts": [
            {
                "width": 665,
                "rest": 535,
                "qty": 1,
                "lengths": [2.8],
                "load_code": 10,
            },
        ],
        "secondary_cuts": [
            {
                "source": 535,
                "source_lengths": [2.8],
                "lengths": [2.8],
                "cuts": [665],
                "qty": 1,
                "pieces": 1,
                "waste": 5,
                "load_code": 8,
            },
        ],
    }


def test_is_crossload_rest_secondary_true_when_source_on_other_load_rest() -> None:
    """Регрессия TypeError: source_lengths — list, target_len — float."""
    plan = _crossload_plan_10p_rest_535()
    sec_cut = plan["secondary_cuts"][0]
    rest_groups = {(670, 2.8): 1}
    assert _is_crossload_rest_secondary(sec_cut, rest_groups, load_key=8, current_plan=plan)


def test_is_crossload_rest_secondary_false_when_source_not_on_other_load_rest() -> None:
    plan = _crossload_plan_10p_rest_535()
    sec_cut = {**plan["secondary_cuts"][0], "source": 999}
    rest_groups = {(670, 2.8): 1}
    assert not _is_crossload_rest_secondary(
        sec_cut, rest_groups, load_key=8, current_plan=plan
    )


REAL_ORDER_8P_TEXT = """
ПБ 72-10,8-8 — 2 шт
ПБ 72-9,2-8 — 3 шт
ПБ 66-10,8-8 — 1 шт
ПБ 63-7,2-8 +доб — 4 шт
ПБ 63-6,65-8 + доб — 4 шт
ПБ 52-5,3-8 + доб — 1 шт
ПБ 31-6,65-8 + доб — 1 шт
"""


def _breakdown_labels(tables: list, *, length_dm: str, width_dm: str) -> list[str]:
    needle = f"ПБ {length_dm}-{width_dm}-"
    table = next(t for t in tables if needle in t.get("name", ""))
    return [row[0] for row in table.get("rows", [])]


def test_real_order_8p_1080_plates_have_long_cut_and_factory_waste() -> None:
    """Интеграция: реальный заказ 8п — 10,8 м с планом и без плана."""
    from core.domain.plate_order import normalize_load_code
    from core.plate_line_parser import parse_line
    from core.plate_order_context import PlateOrderContext
    from app.domain.models.plate_order import PlateOrder as AppPlateOrder
    from app.services.optimization_service import OptimizationService
    from viz_modules.procurement.breakdown import build_component_breakdown
    from viz_modules.procurement.ports import ProcurementDeps

    orders_2d = []
    for raw in REAL_ORDER_8P_TEXT.strip().splitlines():
        r = parse_line(raw.strip())
        assert r.parsed, r.reason_text
        orders_2d.append(
            {
                "length": r.length_m,
                "width": int(round(r.width_m * 1000)),
                "qty": r.qty,
                "load_code": normalize_load_code(r.load_code),
            }
        )

    plate_order = AppPlateOrder.from_orders_2d(orders_2d)
    poc = PlateOrderContext.fresh_empty()
    poc.hydrate_from_order(plate_order)
    opt_ctx = OptimizationService().optimize(
        plate_order, orders_2d=orders_2d, plate_order_ctx=poc
    )
    poc.load_optimization_snapshot(
        optimization_result=opt_ctx.optimization_result,
        plan_by_load=opt_ctx.plan_by_load,
        load_to_reinforcement_map=opt_ctx.load_to_reinforcement_map,
    )

    deps = ProcurementDeps(
        db_path=":memory:",
        get_price=lambda length_m, load_code, db_path: BASE_1_2M_66_8,
        get_raw_material_cost=lambda plate_name, db_path: BASE_1_2M_66_8,
        get_reinforcement=lambda **kwargs: 20.0,
    )

    with poc.bound():
        tables = build_component_breakdown({}, deps=deps, reinforcement_code=8)

    for length_dm, width_dm in (("66", "10,8"), ("72", "10,8")):
        labels = _breakdown_labels(tables, length_dm=length_dm, width_dm=width_dm)
        assert "Продольный рез" in labels
        assert any("Отходы" in label and "120" in label for label in labels), labels

    labels_720 = _breakdown_labels(tables, length_dm="63", width_dm="7,2")
    assert "Продольный рез" in labels_720
    assert any("Отходы" in label for label in labels_720)

    labels_665 = _breakdown_labels(tables, length_dm="63", width_dm="6,65")
    assert "Продольный рез" in labels_665

    labels_530 = _breakdown_labels(tables, length_dm="52", width_dm="5,3")
    # Secondary с rest primary: поперечный + остаток, без второго продольного реза.
    assert "Продольный рез" not in labels_530
    assert "Поперечный рез" in labels_530


def test_real_order_1080_without_plan_uses_waste_not_rest() -> None:
    """Без плана: 10,8 м — «Отходы (120мм)», не «Остаток (120мм)»."""
    from core.plate_order_context import PlateOrderContext
    from core.plate_line_parser import parse_line
    from core.domain.plate_order import normalize_load_code
    from app.domain.models.plate_order import PlateOrder as AppPlateOrder
    from viz_modules.procurement.breakdown import build_component_breakdown
    from viz_modules.procurement.ports import ProcurementDeps

    r = parse_line("ПБ 66-10,8-8 — 1 шт")
    plate_order = AppPlateOrder.from_orders_2d(
        [
            {
                "length": r.length_m,
                "width": int(round(r.width_m * 1000)),
                "qty": r.qty,
                "load_code": normalize_load_code(r.load_code),
            }
        ]
    )
    poc = PlateOrderContext.fresh_empty()
    poc.hydrate_from_order(plate_order)
    deps = ProcurementDeps(
        db_path=":memory:",
        get_price=lambda length_m, load_code, db_path: BASE_1_2M_66_8,
        get_raw_material_cost=lambda plate_name, db_path: BASE_1_2M_66_8,
        get_reinforcement=lambda **kwargs: 20.0,
    )
    with poc.bound():
        tables = build_component_breakdown({}, deps=deps, reinforcement_code=8)

    labels = [row[0] for row in tables[0]["rows"]]
    assert "Продольный рез" in labels
    assert any("Отходы" in label and "120" in label for label in labels)
    assert not any(label.startswith("Остаток (120") for label in labels)


BASE_1_2M_43 = 14498.0


def _foreign_sameload_rest_plan(*, secondary_type: str = "multiple_transverse", waste: int = 175) -> dict:
    """Лента 8.6м (300+900) + свои primary 725×2; secondary 725 из rest 900 той же нагрузки."""
    return {
        "primary_cuts": [
            {"width": 300, "rest": 900, "qty": 1, "lengths": [8.6], "load_code": 8},
            {"width": 725, "rest": 475, "qty": 2, "lengths": [4.3], "load_code": 8},
        ],
        "secondary_cuts": [
            {
                "source": 900,
                "source_lengths": [8.6],
                "lengths": [4.3],
                "cuts": [725],
                "qty": 1,
                "pieces": 1,
                "waste": waste,
                "type": secondary_type,
                "load_code": 8,
            },
        ],
    }


def test_foreign_sameload_rest_secondary_counts_longitudinal_cut() -> None:
    """Secondary 725 из rest 900 чужой primary 8.6м (та же нагрузка):
    продольный рез 8.6м должен быть учтён у позиции со своими primary."""
    plan = _foreign_sameload_rest_plan()
    t = _calc_trim_components(
        plan,
        length=4.3,
        width_mm=725,
        qty=3,
        base_price_1_2m=BASE_1_2M_43,
        base_price=BASE_1_2M_43 * (725 / 1200.0),
        load_code=8,
        price_table={},
    )
    # свои primary: 2 × 4.3м + secondary: 1 × 8.6м
    assert t["long_cut_meterage"] == pytest.approx(2 * 4.3 + 8.6)
    assert t["transverse_remainder_cost"] == pytest.approx(
        BASE_1_2M_43 * (725 / 1200.0) * (4.3 / 4.3) / 3
    )
    assert t["trans_cuts"] == pytest.approx(1.0 / 3)


def test_foreign_sameload_pure_transverse_no_extra_long_cut() -> None:
    """Pure transverse (waste=0, pieces=1) на чужом rest: продольный метраж не растёт."""
    plan = _foreign_sameload_rest_plan(secondary_type="transverse", waste=0)
    t = _calc_trim_components(
        plan,
        length=4.3,
        width_mm=725,
        qty=3,
        base_price_1_2m=BASE_1_2M_43,
        base_price=BASE_1_2M_43 * (725 / 1200.0),
        load_code=8,
        price_table={},
    )
    # только 2 primary-реза по 4.3м; secondary не добавляет продольный рез
    assert t["long_cut_meterage"] == pytest.approx(2 * 4.3)
    assert t["transverse_remainder_cost"] > 0
