"""Cross-surface parity: bot adapter и API path → согласованная структура плана (WP2 / A2)."""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from fastapi.testclient import TestClient

from app.services.production_planning_service import ProductionPlanningService
from bot.services.production_planning_adapter import (
    build_plan_preview,
    plan_structure_signature,
)
from tests.helpers import production_api_fixtures as paf

API_PREFIX = paf.API_PREFIX


def _service_for_db(db_path: str) -> ProductionPlanningService:
    return ProductionPlanningService(plita_db_path=db_path, pb_db_path=db_path)


def test_bot_adapter_matches_api_build_plan_structure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Одинаковые fixtures + mocked optimizer → одинаковая структура плана."""
    api_dir = tmp_path / "api"
    api_dir.mkdir()
    db_api = paf.configure_production_api_env(api_dir, monkeypatch)
    api_result = _service_for_db(db_api).build_plan(
        start_date=paf.DATE_KEY,
        tracks_count=1,
        filter_method="all",
    )

    bot_dir = tmp_path / "bot"
    bot_dir.mkdir()
    db_bot = paf.configure_production_api_env(bot_dir, monkeypatch)
    bot_result = build_plan_preview(
        _service_for_db(db_bot),
        start_date=paf.DATE_KEY,
        tracks_count=1,
        filter_method="all",
    )

    api_sig = plan_structure_signature(api_result["plan"])
    bot_sig = plan_structure_signature(bot_result["plan"])

    assert bot_sig == api_sig
    total_items = sum(
        len(track["items"])
        for day in api_sig["days"].values()
        for track in day["tracks"]
    )
    assert total_items == 3
    assert len(api_sig["days"]) >= 1


def test_bot_adapter_matches_api_via_http(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_api_db: str,
) -> None:
    """На свежей БД bot preview (без persist) и HTTP build дают одинаковую структуру."""
    service = _service_for_db(production_api_db)

    bot_payload = build_plan_preview(
        service,
        start_date=paf.DATE_KEY,
        tracks_count=1,
        filter_method="all",
    )

    api_payload = paf.build_plan_via_api(
        production_api_client,
        production_admin_cookie,
        tracks_count=1,
    )

    assert plan_structure_signature(bot_payload["plan"]) == plan_structure_signature(
        api_payload["plan"]
    )


def test_production_execution_imports_core_pipeline() -> None:
    """Guard: bot handler делегирует в ProductionPlanningService.run_planning_pipeline."""
    import inspect

    from bot.handlers import production_execution

    source = inspect.getsource(production_execution.load_and_plan_production)
    assert "run_planning_pipeline" in source
    assert "load_plates_for_production" in source
    assert "build_layout_sequence" not in source
    assert "split_sequence_into_tracks" not in source
