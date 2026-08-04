"""Unit-тесты движка укладки propose v2 (golden cases + rules)."""

from __future__ import annotations

import pytest

from core.config_and_data import approximate_weight_kg
from core.shipment_packing import PlateCandidate, VehicleLimits, pack_shipment
from core.shipment_packing.marking import marking_length_m
from core.shipment_packing.reasons import NotFitReason, WarningCode
from core.shipment_packing.rules import (
    gost_stack_count,
    is_piece,
    markings_compatible,
)


T20 = VehicleLimits(max_weight_kg=19_800.0, body_length_m=13.2, max_tiers=4)


def _candidate(
    cp_id: int,
    plate_name: str,
    *,
    length_m: float,
    width_m: float = 1.2,
    qty: int,
    kp_id: int = 1,
    unit_weight: float | None = None,
    completed_date: str = "01.01.2026",
) -> PlateCandidate:
    unit = unit_weight if unit_weight is not None else approximate_weight_kg(length_m, width_m)
    return PlateCandidate(
        completed_plate_id=cp_id,
        kp_id=kp_id,
        plate_name=plate_name,
        length_m=length_m,
        width_m=width_m,
        load_class=800,
        qty=qty,
        unit_weight_kg=unit,
        completed_date=completed_date,
    )


def _total_qty(result) -> int:
    return sum(line.qty for line in result.items)


# ---------------------------------------------------------------------------
# marking + rules
# ---------------------------------------------------------------------------


def test_marking_from_plate_name_pb64() -> None:
    marking, fallback = marking_length_m("ПБ 64-12-8", 6.4)
    assert marking == pytest.approx(6.4)
    assert fallback is False


def test_marking_fallback_rounds_length() -> None:
    marking, fallback = marking_length_m("свободный текст", 5.83)
    assert marking == pytest.approx(5.8)
    assert fallback is True


@pytest.mark.parametrize(
    ("marking", "expected"),
    [
        (3.0, 4),
        (3.3, 4),
        (4.0, 3),
        (5.8, 2),
        (6.5, 2),
        (7.3, 1),
        (9.0, 1),
    ],
)
def test_gost_stack_count(marking: float, expected: int) -> None:
    assert gost_stack_count(marking) == expected


def test_piece_detection() -> None:
    assert is_piece(1.19) is True
    assert is_piece(1.2) is False


def test_markings_compatible_cluster_64_74() -> None:
    assert markings_compatible(6.4, 7.4) is True
    assert markings_compatible(6.4, 7.5) is False


# ---------------------------------------------------------------------------
# golden cases
# ---------------------------------------------------------------------------


def test_golden_pb58_ten_plates() -> None:
    result = pack_shipment(
        [_candidate(1, "ПБ 58-12-8", length_m=5.8, qty=15)],
        limits=T20,
    )
    assert _total_qty(result) == 10
    assert result.total_weight_kg <= T20.max_weight_kg
    assert sum(line.qty for line in result.not_fit) == 5


def test_golden_pb635_nine_plates() -> None:
    result = pack_shipment(
        [_candidate(1, "ПБ 63,5-12-8", length_m=6.35, qty=15)],
        limits=T20,
    )
    assert _total_qty(result) == 9
    assert result.total_weight_kg <= T20.max_weight_kg


def test_golden_pb73_eight_plates() -> None:
    unit = approximate_weight_kg(7.3, 1.2)
    # Подгонка под ровно 8× в лимит 19800 (формула даёт чуть больше)
    unit = 19_800 / 8
    result = pack_shipment(
        [_candidate(1, "ПБ 73-12-8", length_m=7.3, qty=12, unit_weight=unit)],
        limits=T20,
    )
    assert _total_qty(result) == 8
    assert result.total_weight_kg == pytest.approx(19_800)


def test_golden_pb90_six_plates() -> None:
    result = pack_shipment(
        [_candidate(1, "ПБ 90-12-8", length_m=9.0, qty=10)],
        limits=T20,
    )
    assert _total_qty(result) == 6
    assert result.total_weight_kg <= T20.max_weight_kg


def test_golden_cluster_64_74_one_stack() -> None:
    result = pack_shipment(
        [
            _candidate(1, "ПБ 64-12-8", length_m=6.4, qty=4, completed_date="01.01.2026"),
            _candidate(2, "ПБ 74-12-8", length_m=7.4, qty=4, completed_date="02.01.2026"),
        ],
        limits=T20,
    )
    assert _total_qty(result) == 8
    assert len(result.not_fit) == 0


def test_golden_piece_same_length_fill() -> None:
    result = pack_shipment(
        [
            _candidate(1, "ПБ 58-12-8", length_m=5.8, qty=7),
            _candidate(2, "ПБ 58-0,6-8", length_m=5.8, width_m=0.6, qty=2),
        ],
        limits=T20,
    )
    piece_line = next((line for line in result.items if line.completed_plate_id == 2), None)
    assert piece_line is not None
    assert piece_line.qty >= 1
    assert _total_qty(result) >= 8


def test_qty_invariant() -> None:
    candidates = [
        _candidate(1, "ПБ 58-12-8", length_m=5.8, qty=12),
        _candidate(2, "ПБ 73-12-8", length_m=7.3, qty=5),
    ]
    result = pack_shipment(candidates, limits=T20)
    for cand in candidates:
        fitted = sum(line.qty for line in result.items if line.completed_plate_id == cand.completed_plate_id)
        rejected = sum(line.qty for line in result.not_fit if line.completed_plate_id == cand.completed_plate_id)
        remainder = next(
            (
                r.qty_remaining
                for r in result.order_remainder
                if r.completed_plate_id == cand.completed_plate_id
            ),
            cand.qty - fitted,
        )
        assert fitted + remainder == cand.qty
        assert rejected == remainder


def test_body_length_hard_not_fit() -> None:
    """Несовместимые длины — два штабеля вдоль кузова, сумма > 13,2 м."""
    result = pack_shipment(
        [
            _candidate(1, "ПБ 70-12-8", length_m=7.0, qty=1, unit_weight=500.0),
            _candidate(2, "ПБ 85-12-8", length_m=8.5, qty=1, unit_weight=500.0),
        ],
        limits=T20,
    )
    assert _total_qty(result) == 1
    assert result.not_fit[0].reason_code == NotFitReason.BODY_LENGTH.value


def test_kp_mix_warning() -> None:
    result = pack_shipment(
        [
            _candidate(1, "ПБ 58-12-8", length_m=5.8, qty=2, kp_id=1),
            _candidate(2, "ПБ 58-12-8", length_m=5.8, qty=2, kp_id=2),
        ],
        limits=T20,
    )
    assert any(w.code == WarningCode.KP_MIX for w in result.warnings)


def test_pieces_only_rejected() -> None:
    result = pack_shipment(
        [_candidate(1, "ПБ 58-0,6-8", length_m=5.8, width_m=0.6, qty=3)],
        limits=T20,
    )
    assert _total_qty(result) == 0
    assert result.not_fit[0].reason_code == NotFitReason.PIECE_PRIORITY.value


# ---------------------------------------------------------------------------
# layout metadata
# ---------------------------------------------------------------------------


def test_layout_metadata_golden() -> None:
    """Эталонный рейс из спеки: 2 штабеля (8,9 + 4,3 м), 5 шагов погрузки."""
    candidates = [
        _candidate(1, "ПБ 42,6-5,3-10п", length_m=4.26, width_m=0.53, qty=1, unit_weight=640.0),
        _candidate(2, "ПБ 42-3,0-8п", length_m=4.2, width_m=0.3, qty=1, unit_weight=357.0),
        _candidate(3, "ПБ 43-12-8п", length_m=4.3, qty=1, unit_weight=1462.0),
        _candidate(4, "ПБ 89-12-8п", length_m=8.9, qty=3, unit_weight=3026.0),
        _candidate(5, "ПБ 80-12-8п", length_m=8.0, qty=3, unit_weight=2720.0),
    ]
    result = pack_shipment(candidates, limits=T20)

    layout = result.layout
    assert layout is not None
    assert layout.body_length_m == pytest.approx(13.2)
    assert layout.body_used_m == pytest.approx(13.2)
    assert len(layout.stacks) == 2

    s1, s2 = layout.stacks
    assert s1.index == 1
    assert s1.marking_length_m == pytest.approx(8.9)
    assert s2.index == 2
    assert s2.marking_length_m == pytest.approx(4.3)

    assert len(s1.tiers) == 3
    assert [u.plate_name for u in s1.tiers[0].units] == ["ПБ 89-12-8п", "ПБ 89-12-8п"]
    assert sorted(u.plate_name for u in s1.tiers[1].units) == ["ПБ 80-12-8п", "ПБ 89-12-8п"]
    assert [u.plate_name for u in s1.tiers[2].units] == ["ПБ 80-12-8п", "ПБ 80-12-8п"]

    assert len(s2.tiers) == 2
    assert [u.plate_name for u in s2.tiers[0].units] == ["ПБ 43-12-8п", "ПБ 42,6-5,3-10п"]
    assert [u.plate_name for u in s2.tiers[1].units] == ["ПБ 42-3,0-8п"]

    steps = layout.loading_steps
    assert len(steps) == 5
    assert [s.step for s in steps] == [1, 2, 3, 4, 5]
    assert [s.stack_index for s in steps] == [1, 1, 1, 2, 2]
    assert [s.tier_index for s in steps] == [1, 2, 3, 1, 2]
    assert [s.description for s in steps] == [
        "ПБ 89-12-8п ×2",
        "ПБ 89-12-8п + ПБ 80-12-8п",
        "ПБ 80-12-8п ×2",
        "ПБ 43-12-8п + ПБ 42,6-5,3-10п",
        "ПБ 42-3,0-8п",
    ]

    items_qty = sum(line.qty for line in result.items)
    layout_qty = sum(len(tier.units) for stack in layout.stacks for tier in stack.tiers)
    assert items_qty == layout_qty == 9
