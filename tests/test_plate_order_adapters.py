"""A3-001: app/core plate order boundary adapters."""

from __future__ import annotations

import dataclasses

import pytest

from app.domain.adapters.plate_order import from_core_order, to_core_order
from app.domain.models.plate_order import PlateOrder as AppPlateOrder
from core.domain.plate_order import PlateOrder as CorePlateOrder
from core.domain.plate_order import coerce_core_plate_order


def _sample_app_order() -> AppPlateOrder:
    key = (3.39, 1.2, 8.0, "339")
    order = AppPlateOrder()
    order.plates_1_2 = [3.39, 3.39]
    order.plate_load_details[key] = 2
    order.plate_length_dm_raw[key] = "339"
    order.nomenclature_cache[key] = {"canonical_name": "ПБ 39-12-8", "plate_name": "ПБ 39"}
    return order


def test_to_core_order_strips_nomenclature_cache() -> None:
    app = _sample_app_order()

    core = to_core_order(app)

    assert type(core) is CorePlateOrder
    assert not hasattr(core, "nomenclature_cache")
    assert core.plates_1_2 == app.plates_1_2
    assert core.plate_load_details == app.plate_load_details


def test_to_core_order_returns_core_instance_unchanged() -> None:
    core = CorePlateOrder()
    core.plates_1_2 = [1.0]

    assert to_core_order(core) is core


def test_from_core_order_builds_app_with_cache_override() -> None:
    core = CorePlateOrder()
    key = (5.0, 1.0, 8.0, "")
    cache = {key: {"sku": "x"}}

    app = from_core_order(core, nomenclature_cache=cache)

    assert isinstance(app, AppPlateOrder)
    assert app.nomenclature_cache == cache


def test_from_core_order_preserves_app_cache_when_same_instance() -> None:
    app = _sample_app_order()

    assert from_core_order(app) is app


def test_roundtrip_to_core_from_core_preserves_fields_and_cache() -> None:
    app = _sample_app_order()
    key = next(iter(app.nomenclature_cache))

    core = to_core_order(app)
    restored = from_core_order(core, nomenclature_cache=dict(app.nomenclature_cache))

    assert restored.plates_1_2 == app.plates_1_2
    assert restored.plate_load_details == app.plate_load_details
    assert restored.plate_length_dm_raw == app.plate_length_dm_raw
    assert restored.nomenclature_cache[key] == app.nomenclature_cache[key]


def test_from_dict_roundtrip_preserves_nomenclature_cache() -> None:
    app = _sample_app_order()

    restored = AppPlateOrder.from_dict(app.to_dict())

    assert restored.plates_1_2 == app.plates_1_2
    assert restored.nomenclature_cache == app.nomenclature_cache


def test_coerce_core_plate_order_is_generic_no_strip() -> None:
    app = _sample_app_order()

    coerced = coerce_core_plate_order(app)

    assert type(coerced) is CorePlateOrder
    assert coerced.plate_load_details == app.plate_load_details


def test_from_core_order_copies_core_fields() -> None:
    core = CorePlateOrder.from_orders_2d(
        [{"length": 3.39, "width": 1200, "qty": 1, "load_code": 8}]
    )

    app = from_core_order(core)

    for field in dataclasses.fields(CorePlateOrder):
        assert getattr(app, field.name) == getattr(core, field.name)
    assert app.nomenclature_cache == {}
