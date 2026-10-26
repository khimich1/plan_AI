"""Чистые функции отчёта пула пар: остаток реза, цена, нарезка пачек."""

from __future__ import annotations

from collections import Counter

from scripts.build_kz_pool_pair_report_2026 import (
    _against,
    catalog_sale_price,
    cut_accounts,
    rescued_secondaries,
    rub_delta,
    split_batches,
    tail_tracks,
)


def test_split_batches_keeps_short_tail():
    assert split_batches(8, 5) == [(0, 5), (5, 8)]
    assert tail_tracks(8, 5) == 3
    assert tail_tracks(10, 5) == 0
    assert tail_tracks(3, 5) == 3


def test_unpaired_rest_is_primary_without_secondary_parent():
    primary = [
        {"rest": 0, "lengths": [6.0], "primary_instance_id": "prim-1"},
        {"rest": 500, "lengths": [7.2], "primary_instance_id": "prim-2"},
        {"rest": 400, "lengths": [5.0], "primary_instance_id": "prim-3"},
    ]
    secondary = [
        {"parent_instance_id": "prim-2", "source": 500, "source_lengths": [7.2], "waste": 0, "type": "multiple"},
    ]
    accounts = cut_accounts(primary, secondary)
    assert accounts["unpaired"] == 1
    assert accounts["keys"][(5.0, 400)] == 1
    assert abs(accounts["rest_m2"] - 0.4 * 5.0) < 1e-9


def test_rescued_plates_share_one_unused_rest():
    available = Counter({(7.2, 500): 1})
    secondary = [
        {
            "parent_instance_id": "prim-2",
            "source": 500,
            "source_lengths": [7.2],
            "cuts": [300],
            "lengths": [7.2],
            "load_code": 8,
        },
        {
            "parent_instance_id": "prim-2",
            "source": 500,
            "source_lengths": [7.2],
            "cuts": [200],
            "lengths": [7.2],
            "load_code": 8,
        },
        {
            "parent_instance_id": "prim-9",
            "source": 500,
            "source_lengths": [7.2],
            "cuts": [500],
            "lengths": [7.2],
            "load_code": 8,
        },
    ]
    rescued = rescued_secondaries(available, secondary)
    assert len(rescued) == 2
    assert [plate["width_mm"] for plate in rescued] == [300, 200]


def test_plate_cost_is_full_blank_regardless_of_width():
    from scripts.build_kz_pool_pair_report_2026 import plate_cost

    costs = {(63, 8.0): 8000.0, (63, 12.5): 9000.0}
    assert plate_cost(costs, 6.3, 600, 8) == 8000.0
    assert plate_cost(costs, 6.3, 1200, 13) == 9000.0
    assert plate_cost(costs, 6.3, 1200, 6) == 8000.0


def test_catalog_price_scales_width_and_maps_load_13():
    prices = {(72, 12): 10000.0, (50, 8): 0.0}
    full = catalog_sale_price(prices, 7.2, 1200, 13)
    half = catalog_sale_price(prices, 7.2, 600, 12)
    assert full == 10000.0
    assert half == 5000.0
    assert catalog_sale_price(prices, 5.0, 1200, 8) is None


def test_delta_is_negative_when_current_is_cheaper():
    assert rub_delta(80, 100) == -20


def test_split_piece_keeps_ordered_width():
    from scripts.build_kz_pool_pair_report_2026 import order_pieces_from_layout

    track = {
        "items": [{
            "mode": "split",
            "length": 6.3,
            "main_w": 1.07,
            "rest_w": 0.13,
            "load_code": 8,
            "reinforcement": 18,
            "plate_name": "ПБ 63-11-8п",
            "secondary_cuts": [{
                "width": 0.46,
                "target_length": 6.3,
                "load_code": 8,
                "reinforcement": 16,
                "label": "ПБ 63-4,6-8п",
            }],
        }],
    }
    pieces, _scrap = order_pieces_from_layout(track, lambda _ld, _load: 18)
    assert [piece["width_mm"] for piece in pieces] == [1070, 460]
    assert pieces[1]["wires"] == 16


def test_against_sums_only_batches_with_base():
    batches = [
        {"ok": True, "meters": 100, "wire_rub": 50, "concrete_rub": 10, "base_map": {"meters": 110, "wire_rub": 80, "concrete_rub": 10}},
        {"ok": False, "meters": 0, "wire_rub": 0, "concrete_rub": 0},
    ]
    d_m, d_wire, _d_wire_m, d_concrete, _d_concrete_m = _against(batches, "base_map")
    assert d_m == -10
    assert d_wire == -30
    assert d_concrete == 0
