"""Тесты распределения дорожек по дням с учётом глобальной занятости.

Цель — убедиться, что ``plan_manager.distribute_tracks_by_days`` теперь
корректно пропускает переполненные дни и не превышает ``max_per_day``.
Патчим праздничный календарь на пустой, чтобы тесты были детерминированными.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from bot.handlers import plan_manager


@pytest.fixture(autouse=True)
def _patch_calendar(monkeypatch):
    """Все дни 21-30 апреля 2026 трактуются по умолчанию (пн-пт рабочие)."""
    monkeypatch.setattr(plan_manager, "load_holidays", lambda: set())
    monkeypatch.setattr(plan_manager, "load_extra_workdays", lambda: set())


def _tracks(n: int) -> list[dict]:
    return [{"id": i} for i in range(n)]


def test_distribute_respects_max_per_day_without_occupancy():
    """Без занятых слотов ``tracks_per_day`` используется как верхняя граница."""
    result = plan_manager.distribute_tracks_by_days(
        _tracks(7),
        start_date="2026-04-21",  # вторник
        tracks_per_day=3,
    )

    # 3 + 3 + 1 по трём последовательным рабочим дням
    assert result == {
        "2026-04-21": _tracks(3),
        "2026-04-22": _tracks(3)[0:0] + [{"id": 3}, {"id": 4}, {"id": 5}],
        "2026-04-23": [{"id": 6}],
    }


def test_distribute_skips_overbooked_day_and_pushes_forward():
    """Если день переполнен другими планами — пропускаем и идём дальше."""
    occupancy = {"2026-04-21": 5}  # MAX_TRACKS_PER_DAY

    result = plan_manager.distribute_tracks_by_days(
        _tracks(6),
        start_date="2026-04-21",
        tracks_per_day=3,
        global_occupancy=occupancy,
    )

    assert "2026-04-21" not in result
    assert result == {
        "2026-04-22": [{"id": 0}, {"id": 1}, {"id": 2}],
        "2026-04-23": [{"id": 3}, {"id": 4}, {"id": 5}],
    }


def test_distribute_caps_chunk_to_available_free_slots():
    """``tracks_per_day`` кэпится остатком ``max_per_day - occupancy``."""
    occupancy = {
        "2026-04-21": 3,  # свободно 2
        "2026-04-22": 0,  # свободно 5
    }

    result = plan_manager.distribute_tracks_by_days(
        _tracks(5),
        start_date="2026-04-21",
        tracks_per_day=4,
        global_occupancy=occupancy,
    )

    # 21.04: занято 3 + max 2 = 5; 22.04: занято 0, хотели 4, кладём остаток 3
    assert result == {
        "2026-04-21": [{"id": 0}, {"id": 1}],
        "2026-04-22": [{"id": 2}, {"id": 3}, {"id": 4}],
    }


def test_distribute_uses_custom_max_per_day():
    """Параметр ``max_per_day`` может быть больше 5 (например, для тестов)."""
    occupancy = {"2026-04-21": 5}

    result = plan_manager.distribute_tracks_by_days(
        _tracks(3),
        start_date="2026-04-21",
        tracks_per_day=10,
        global_occupancy=occupancy,
        max_per_day=10,  # свободно 10-5=5
    )

    assert result == {"2026-04-21": _tracks(3)}
