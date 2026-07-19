"""Доменные перечисления для статусов плит и КП.

Эти Enum-ы — единая точка правды для значений, которые исторически
хранились в БД как строки на русском. Унаследованы от ``str``, поэтому
пригодны для прямой подстановки в SQL и сериализации без ``.value``,
но в коде предпочтителен явный ``PlateStatus.IN_PRODUCTION.value`` —
чтобы IDE подсвечивала опечатки.

Реальный набор значений в БД:
- ``kp_plates.status``: ``'в производстве'``, ``'в плане'``
- ``kp_meta.status``:   ``'в архиве'``, ``'в работе'``, ``'выполнено'``

Виртуальные («в логе аудита, но не в столбце») значения:
- ``PlateStatus.COMPLETED`` — после переноса в ``completed_plates``
"""
from __future__ import annotations

from enum import Enum


class PlateStatus(str, Enum):
    """Статусы плиты в её жизненном цикле."""

    IN_PRODUCTION = "в производстве"
    IN_PLAN = "в плане"
    # COMPLETED — псевдо-статус: реальной строки в kp_plates с этим статусом
    # нет, но в plate_status_log это полезное значение для записи
    # завершения (move_plates_to_completed).
    COMPLETED = "completed"


class KpStatus(str, Enum):
    """Статусы коммерческого предложения (KP)."""

    ARCHIVED = "в архиве"
    IN_WORK = "в работе"
    DONE = "выполнено"


class PlateTransitionReason(str, Enum):
    """Причина перехода статуса в audit-логе."""

    PLANNED = "planned"
    COMPLETED = "completed"
    REJECTED = "rejected"
    PLAN_ROLLBACK = "plan_rollback"
