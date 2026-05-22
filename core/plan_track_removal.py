"""Сбор возвратов плит при удалении дорожки из производственного плана.

Модуль не зависит от FastAPI и aiogram. Используется
:func:`app.planning.plan_manager.remove_track_from_plan` вместе с
:func:`core.kp_db.return_plate_rows_for_plan`.

Физические плиты дорожки (root items + ``secondary_cuts``) считаются так же,
как при коммите плана в :mod:`core.plan_commit` (P9).
"""
from __future__ import annotations

from collections import Counter
from typing import Any

from core.plan_commit import _identity_for_track_item, _iter_physical_items

LegacyIdentity = tuple[int, str]


class TrackRemovalError(Exception):
    """Доменная ошибка при удалении дорожки из плана."""

    def __init__(self, message: str, code: str | None = None) -> None:
        super().__init__(message)
        self.message = message
        self.code = code


def collect_plate_returns_from_track(
    track: dict[str, Any],
) -> tuple[Counter[int], Counter[LegacyIdentity]]:
    """Собирает счётчики плит для возврата в БД при удалении дорожки.

    Для каждого физического item (root + ``secondary_cuts``):

    - если задан ``kp_plate_id`` — увеличивает ``id_qty[kp_plate_id]``;
    - иначе, если есть ``kp_id`` и каноническое имя плиты — увеличивает
      ``legacy_identity_qty[(kp_id, canonical_name)]`` (legacy-планы без id).

    Args:
        track: словарь дорожки из ``plan['days'][date]['tracks'][i]``.

    Returns:
        Кортеж ``(id_qty, legacy_identity_qty)``.
    """
    id_qty: Counter[int] = Counter()
    legacy_identity_qty: Counter[LegacyIdentity] = Counter()

    for physical in _iter_physical_items(track.get("items")):
        kp_plate_id = physical.get("kp_plate_id")
        if kp_plate_id is not None:
            id_qty[int(kp_plate_id)] += 1
            continue

        identity = _identity_for_track_item(physical)
        if identity is not None:
            legacy_identity_qty[identity] += 1

    return id_qty, legacy_identity_qty


def collect_returns_from_track(
    track: dict[str, Any],
) -> tuple[Counter[int], Counter[LegacyIdentity]]:
    """Алиас :func:`collect_plate_returns_from_track` (имя из plan_manager)."""
    return collect_plate_returns_from_track(track)


__all__ = [
    "LegacyIdentity",
    "TrackRemovalError",
    "collect_plate_returns_from_track",
    "collect_returns_from_track",
]
