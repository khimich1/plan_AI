"""Тесты согласованности подсчёта дорожек на дату.

Закрывают баг, когда календарь показывал «3/5 занято», а Drawer — «на этот
день не запланировано дорожек»: два агрегатора в
``app.planning.plan_manager`` опирались на разные источники truth
(``saved_tracks_count`` vs фактический массив ``tracks``). После фикса
оба используют :func:`app.planning.plan_manager.count_day_tracks`.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from app.planning import plan_manager


def test_count_day_tracks_returns_len_of_tracks():
    day_data = {
        "tracks": [{"items": []}, {"items": []}],
        "saved_tracks_count": 2,
    }
    assert plan_manager.count_day_tracks(day_data) == 2


def test_count_day_tracks_ignores_stale_saved_tracks_count(caplog):
    """Рассинхрон: saved_tracks_count=3, но tracks пустой → считаем 0."""
    day_data = {"tracks": [], "saved_tracks_count": 3}

    with caplog.at_level("WARNING", logger="app.planning.plan_manager"):
        result = plan_manager.count_day_tracks(day_data)

    assert result == 0
    assert any("Рассинхрон" in rec.message for rec in caplog.records), (
        "Ожидали WARNING в логе при расхождении saved_tracks_count и len(tracks)"
    )


def test_count_day_tracks_handles_missing_fields():
    assert plan_manager.count_day_tracks({}) == 0
    assert plan_manager.count_day_tracks({"tracks": None}) == 0


def _patch_plans(monkeypatch, plans_by_id: dict[str, dict]) -> None:
    """Подменяет загрузку метаданных и планов заранее подготовленными данными."""
    metadata = {"plans": [{"id": plan_id} for plan_id in plans_by_id]}
    monkeypatch.setattr(plan_manager, "load_plans_metadata", lambda: metadata)
    monkeypatch.setattr(plan_manager, "load_plan", lambda plan_id: plans_by_id.get(plan_id))


def test_calendar_and_detail_agree_on_broken_day(monkeypatch):
    """Главный инвариант: если в плане tracks=[] и saved_tracks_count>0,
    календарь и детали дня должны вернуть одно и то же число (0)."""
    broken_plan = {
        "id": "plan_broken",
        "name": "Битый план",
        "days": {
            "2026-04-22": {
                "date": "2026-04-22",
                "day_number": 1,
                "tracks": [],
                "saved_tracks_count": 3,
                "total_tracks_count": 5,
                "completed": False,
            },
        },
    }
    _patch_plans(monkeypatch, {"plan_broken": broken_plan})

    occupancy = plan_manager.get_global_day_occupancy()
    calendar = plan_manager.get_global_calendar_info()
    multi = plan_manager.get_tracks_for_date_from_all_plans("2026-04-22")

    assert occupancy.get("2026-04-22", 0) == 0, (
        "Календарь не должен доверять saved_tracks_count при пустом tracks"
    )
    assert calendar is not None
    assert calendar["days_info"]["2026-04-22"]["occupied"] == 0
    assert multi is not None
    assert len(multi["tracks"]) == 0


def test_calendar_and_detail_agree_on_normal_day(monkeypatch):
    """На здоровом плане оба агрегатора возвращают одинаковое число."""
    healthy_plan = {
        "id": "plan_ok",
        "name": "Рабочий план",
        "days": {
            "2026-04-22": {
                "date": "2026-04-22",
                "day_number": 1,
                "tracks": [{"items": []}, {"items": []}, {"items": []}],
                "saved_tracks_count": 3,
                "total_tracks_count": 5,
                "completed": False,
            },
        },
    }
    _patch_plans(monkeypatch, {"plan_ok": healthy_plan})

    occupancy = plan_manager.get_global_day_occupancy()
    multi = plan_manager.get_tracks_for_date_from_all_plans("2026-04-22")

    assert occupancy["2026-04-22"] == 3
    assert multi is not None
    assert len(multi["tracks"]) == 3
