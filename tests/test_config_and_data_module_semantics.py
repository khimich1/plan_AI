# -*- coding: utf-8 -*-
"""PEP 562 ``__getattr__`` + ``get_config()`` alignment with constants re-exports."""

import pytest

import core.config_and_data as cfg


def test_get_config_matches_constants() -> None:
    c = cfg.get_config()
    assert c.track_length_m == cfg.TRACK_LENGTH_M
    assert c.track_width_m == cfg.TRACK_WIDTH_M
    assert c.long_cut_price_per_m == cfg.LONG_CUT_PRICE_PER_M
    assert c.transverse_cut_price == cfg.TRANSVERSE_CUT_PRICE
    assert c.weight_kg_per_dm2 == cfg.WEIGHT_KG_PER_DM2


def test_getattr_plate_load_details_is_runtime_dict() -> None:
    rt = cfg.get_plate_mutable_runtime()
    x = cfg.PLATE_LOAD_DETAILS
    assert x is rt.plate_load_details
    before = dict(x)
    try:
        x.clear()
        assert len(cfg.PLATE_LOAD_DETAILS) == 0
    finally:
        x.clear()
        x.update(before)


def test_unknown_attribute_raises() -> None:
    with pytest.raises(AttributeError):
        getattr(cfg, "NOT_A_REAL_CONFIG_AND_DATA_SYMBOL")
