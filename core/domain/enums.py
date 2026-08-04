"""Доменные перечисления для статусов плит и КП.

Эти Enum-ы — единая точка правды для значений, которые исторически
хранились в БД как строки на русском. Унаследованы от ``str``, поэтому
пригодны для прямой подстановки в SQL и сериализации без ``.value``,
но в коде предпочтителен явный ``PlateStatus.IN_PRODUCTION.value`` —
чтобы IDE подсвечивала опечатки.

Реальный набор значений в БД:
- ``kp_plates.status``: ``'в производстве'``, ``'в плане'``
- ``kp_meta.status``:   ``'в архиве'``, ``'в работе'``, ``'выполнено'``, ``'На СГП'``

Виртуальные («в логе аудита, но не в столбце») значения:
- ``PlateStatus.ON_SGP`` — после отправки дня на склад готовой продукции
- ``PlateStatus.COMPLETED`` — deprecated для новых записей; читать как on_sgp
"""
from __future__ import annotations

from enum import Enum


class PlateStatus(str, Enum):
    """Статусы плиты в её жизненном цикле."""

    IN_PRODUCTION = "в производстве"
    IN_PLAN = "в плане"
    # ON_SGP — псевдо-статус: плита лежит в completed_plates (СГП).
    ON_SGP = "on_sgp"
    # SHIPPED — псевдо-статус в audit: плита отгружена рейсом (уже не на складе).
    SHIPPED = "shipped"
    # COMPLETED — deprecated; совместимость со старым audit (читать как on_sgp).
    COMPLETED = "completed"


class KpStatus(str, Enum):
    """Статусы коммерческого предложения (KP)."""

    ARCHIVED = "в архиве"
    IN_WORK = "в работе"
    ON_SGP = "На СГП"
    # DONE — отгрузка (OUT of MVP); не выставлять из send_to_sgp.
    DONE = "выполнено"


class PlateTransitionReason(str, Enum):
    """Причина перехода статуса в audit-логе."""

    PLANNED = "planned"
    COMPLETED = "completed"  # deprecated; новые записи — SGP_SEND
    REJECTED = "rejected"
    PLAN_ROLLBACK = "plan_rollback"
    SGP_SEND = "sgp_send"
    SGP_UNLINK = "sgp_unlink"
    SGP_RELINK = "sgp_relink"
    SGP_RESERVE = "sgp_reserve"
    SGP_SHIP = "sgp_ship"


class ShipmentStatus(str, Enum):
    """Статусы рейса отгрузки (раздел «Логистика»)."""

    IN_WORK = "in_work"
    DONE = "done"


class DeliveryType(str, Enum):
    """Тип выдачи: доставка или самовывоз."""

    DELIVERY = "delivery"
    PICKUP = "pickup"


class ShipmentItemType(str, Enum):
    """Тип строки состава рейса: плита со СГП или свободная позиция (сваи)."""

    PLATE = "plate"
    FREE = "free"
