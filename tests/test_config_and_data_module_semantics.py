# -*- coding: utf-8 -*-
"""``get_config()`` и отсутствие PEP 562-прокси заказа плит."""

import pytest

import core.config_and_data as cfg


def test_get_config_matches_constants() -> None:
    c = cfg.get_config()
    assert c.track_length_m == cfg.TRACK_LENGTH_M
    assert c.track_width_m == cfg.TRACK_WIDTH_M
    assert c.long_cut_price_per_m == cfg.LONG_CUT_PRICE_PER_M
    assert c.transverse_cut_price == cfg.TRANSVERSE_CUT_PRICE
    assert c.weight_kg_per_dm2 == cfg.WEIGHT_KG_PER_DM2


def test_legacy_plate_names_are_not_module_attributes() -> None:
    for name in (
        "PLATE_LOAD_DETAILS",
        "PLATES_1_2",
        "PLATES_1_0",
        "PLATES_1_08",
        "PLATES_0_32",
        "PLATES_0_46",
        "PLATE_EXACT_WIDTHS",
        "PLATE_LENGTH_DM_RAW",
    ):
        with pytest.raises(AttributeError):
            getattr(cfg, name)


def test_bound_runtime_plate_load_details_is_the_order_dict() -> None:
    from core.plate_order_context import PlateOrderContext
    from core.plate_runtime_state import get_plate_mutable_runtime

    ctx = PlateOrderContext.fresh_empty()
    with ctx.bound():
        details = get_plate_mutable_runtime().plate_load_details
        assert details is ctx.plates.plate_load_details
        details[(1.0, 1.2, 8, "")] = 1
        assert ctx.plates.plate_load_details[(1.0, 1.2, 8, "")] == 1


def test_unknown_attribute_raises() -> None:
    with pytest.raises(AttributeError):
        getattr(cfg, "NOT_A_REAL_CONFIG_AND_DATA_SYMBOL")


def test_decommissioned_proxy_names_raise() -> None:
    """P4 A4: cuts/strips/metadata caches removed from PEP 562 proxy."""
    for name in (
        "LONGITUDINAL_CUTS",
        "WASTE_AREA_M2",
        "PLATE_METADATA",
        "PLATE_NOMENCLATURE_CACHE",
        "LAST_PARSE_DIAGNOSTICS",
        "PLATES_1_5_TO_1_2",
    ):
        with pytest.raises(AttributeError, match=name):
            getattr(cfg, name)
