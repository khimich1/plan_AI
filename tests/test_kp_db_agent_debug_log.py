"""Tests for gated agent NDJSON logs in kp_db hot paths."""

from __future__ import annotations

import json

import pytest

from core.debug_paths import append_agent_debug_log, get_debug_log_path, kp_db_agent_debug_active


@pytest.fixture(autouse=True)
def _clear_settings_cache() -> None:
    from app.core.settings import get_settings

    get_settings.cache_clear()
    yield
    get_settings.cache_clear()


def test_append_skips_when_debug_disabled(
    tmp_path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.delenv("KP_DB_AGENT_DEBUG", raising=False)
    monkeypatch.delenv("APP_DEBUG", raising=False)
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")

    log_path = tmp_path / "agent-test.log"
    append_agent_debug_log(log_path, {"message": "secret", "kp_id": 1})
    assert not log_path.exists()


def test_append_writes_when_kp_db_agent_debug_enabled(
    tmp_path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("KP_DB_AGENT_DEBUG", "1")
    log_path = tmp_path / "agent-test.log"
    append_agent_debug_log(log_path, {"message": "ok", "kp_id": 42})
    assert log_path.exists()
    line = json.loads(log_path.read_text(encoding="utf-8").strip())
    assert line["kp_id"] == 42


def test_kp_db_agent_debug_active_respects_app_debug(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("KP_DB_AGENT_DEBUG", raising=False)
    monkeypatch.setenv("APP_DEBUG", "true")
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    assert kp_db_agent_debug_active() is True


def test_find_kp_plate_row_writes_no_log_when_debug_disabled(
    tmp_path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Matching hot path must not touch debug_logs when agent debug is off."""
    import sqlite3

    from core.domain.plate_completion_matching import find_kp_plate_row
    from tests.helpers import kp_db_fixtures as fx

    monkeypatch.delenv("KP_DB_AGENT_DEBUG", raising=False)
    monkeypatch.delenv("APP_DEBUG", raising=False)
    db = fx.make_iso_db(tmp_path)
    fx.seed_kp_offer(db, 1)
    fx.seed_plate(
        db,
        kp_id=1,
        plate_name="ПБ 60-12-8п",
        length_m=6.0,
        width_m=1.2,
        qty=1,
        status="в производстве",
    )
    log_path = get_debug_log_path("debug.log")
    if log_path.exists():
        log_path.unlink()

    with sqlite3.connect(db) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        find_kp_plate_row(cur, "ПБ 60-12-8п", 6.0, 1.2, 800, 1)

    assert not log_path.exists()
