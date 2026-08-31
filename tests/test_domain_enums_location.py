"""A1: app.domain.enums re-exports the same objects as core.domain.enums."""

from __future__ import annotations

from app.domain import enums as app_enums
from core.domain import enums as core_enums


def test_domain_enums_same_object_identity() -> None:
    assert app_enums.PlateStatus is core_enums.PlateStatus
    assert app_enums.KpStatus is core_enums.KpStatus
    assert app_enums.PlateTransitionReason is core_enums.PlateTransitionReason
    assert app_enums.ShipmentStatus is core_enums.ShipmentStatus
    assert app_enums.DeliveryType is core_enums.DeliveryType
    assert app_enums.ShipmentItemType is core_enums.ShipmentItemType
