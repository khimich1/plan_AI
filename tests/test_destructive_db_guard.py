"""Tests for production guards on destructive kp_db resets."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from app.core.settings import get_settings
from app.services.admin_service import AdminService
from core import kp_db
from core.destructive_db_guard import (
    DestructiveDbFlagsMisconfigured,
    DestructiveDbOperationBlocked,
    destructive_db_reset_allowed,
    fail_fast_if_destructive_flags_in_production,
    require_destructive_db_reset,
)


@pytest.fixture(autouse=True)
def _clear_settings_cache() -> None:
    get_settings.cache_clear()
    yield
    get_settings.cache_clear()


def test_destructive_blocked_in_development_without_flag(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_ENV", "development")
    monkeypatch.delenv("ALLOW_DESTRUCTIVE_DB_RESET", raising=False)
    monkeypatch.delenv("DESTRUCTIVE_DB_RESET_BREAK_GLASS", raising=False)
    assert destructive_db_reset_allowed() is False
    with pytest.raises(DestructiveDbOperationBlocked):
        require_destructive_db_reset()


def test_destructive_allowed_in_development_with_allow(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_ENV", "development")
    monkeypatch.setenv("ALLOW_DESTRUCTIVE_DB_RESET", "1")
    monkeypatch.delenv("DESTRUCTIVE_DB_RESET_BREAK_GLASS", raising=False)
    assert destructive_db_reset_allowed() is True
    require_destructive_db_reset()


def test_destructive_blocked_in_production(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.delenv("ALLOW_DESTRUCTIVE_DB_RESET", raising=False)
    assert destructive_db_reset_allowed() is False
    with pytest.raises(DestructiveDbOperationBlocked):
        require_destructive_db_reset()


def test_destructive_blocked_in_staging(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_ENV", "staging")
    monkeypatch.delenv("ALLOW_DESTRUCTIVE_DB_RESET", raising=False)
    assert destructive_db_reset_allowed() is False
    with pytest.raises(DestructiveDbOperationBlocked):
        require_destructive_db_reset()


def test_destructive_blocked_in_production_with_allow_only(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("ALLOW_DESTRUCTIVE_DB_RESET", "1")
    monkeypatch.delenv("DESTRUCTIVE_DB_RESET_BREAK_GLASS", raising=False)
    assert destructive_db_reset_allowed() is False
    with pytest.raises(DestructiveDbOperationBlocked):
        require_destructive_db_reset()


def test_destructive_blocked_in_staging_with_allow_only(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_ENV", "staging")
    monkeypatch.setenv("ALLOW_DESTRUCTIVE_DB_RESET", "1")
    monkeypatch.delenv("DESTRUCTIVE_DB_RESET_BREAK_GLASS", raising=False)
    assert destructive_db_reset_allowed() is False
    with pytest.raises(DestructiveDbOperationBlocked):
        require_destructive_db_reset()


def test_destructive_allowed_in_staging_with_break_glass(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_ENV", "staging")
    monkeypatch.setenv("ALLOW_DESTRUCTIVE_DB_RESET", "1")
    monkeypatch.setenv("DESTRUCTIVE_DB_RESET_BREAK_GLASS", "1")
    assert destructive_db_reset_allowed() is True
    require_destructive_db_reset()


def test_destructive_allowed_in_production_with_break_glass(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("ALLOW_DESTRUCTIVE_DB_RESET", "1")
    monkeypatch.setenv("DESTRUCTIVE_DB_RESET_BREAK_GLASS", "1")
    assert destructive_db_reset_allowed() is True
    require_destructive_db_reset()


def test_clear_all_kp_blocked_in_production(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.delenv("ALLOW_DESTRUCTIVE_DB_RESET", raising=False)
    db_path = str(tmp_path / "plita.db")
    kp_db.init_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO KP_offers (creation_date, customer_name) VALUES (?, ?)",
            ("01.01.2026", "Test"),
        )
        conn.commit()

    with pytest.raises(DestructiveDbOperationBlocked):
        kp_db.clear_all_kp(db_path)

    with sqlite3.connect(db_path) as conn:
        count = conn.execute("SELECT COUNT(*) FROM KP_offers").fetchone()[0]
    assert int(count) == 1


@pytest.fixture()
def _admin_settings(tmp_path: Path):
    from app.core.settings import Settings

    plans_dir = tmp_path / "plans"
    plans_dir.mkdir()
    return Settings(
        app_secret_key="test-secret-key-for-pytest-must-be-32-chars-min",
        plita_db_path=tmp_path / "plita.db",
        plans_dir=plans_dir,
        plans_metadata_path=tmp_path / "plans_metadata.json",
        current_plan_path=tmp_path / "current_plan.json",
        work_calendar_path=tmp_path / "work_calendar.json",
    )


@pytest.fixture()
def _populated_db(_admin_settings):
    from app.repositories.auth_repository import AuthRepository

    db_path = str(_admin_settings.plita_db_path)
    kp_db.init_schema(db_path)
    AuthRepository(db_path=db_path).create_or_update_user(
        username="root_admin",
        password="AdminTestPass12!",
        role="admin",
    )
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO KP_offers (creation_date, customer_name) VALUES (?, ?)",
            ("01.01.2026", "ООО Тест"),
        )
        conn.commit()
    return _admin_settings.plita_db_path


def test_admin_reset_full_blocked_in_production(
    _admin_settings,
    _populated_db: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.delenv("ALLOW_DESTRUCTIVE_DB_RESET", raising=False)
    service = AdminService(settings=_admin_settings)

    with pytest.raises(DestructiveDbOperationBlocked):
        service.reset_full()

    assert _table_count(_populated_db, "KP_offers") > 0


def _table_count(db_path: Path | str, table: str) -> int:
    with sqlite3.connect(str(db_path)) as conn:
        cur = conn.cursor()
        cur.execute(f"SELECT COUNT(*) FROM {table}")
        return int(cur.fetchone()[0])


def test_fail_fast_when_both_destructive_flags_in_production(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("ALLOW_DESTRUCTIVE_DB_RESET", "1")
    monkeypatch.setenv("DESTRUCTIVE_DB_RESET_BREAK_GLASS", "1")
    with pytest.raises(DestructiveDbFlagsMisconfigured):
        fail_fast_if_destructive_flags_in_production("production")


def test_fail_fast_skips_development(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_ENV", "development")
    monkeypatch.setenv("ALLOW_DESTRUCTIVE_DB_RESET", "1")
    monkeypatch.setenv("DESTRUCTIVE_DB_RESET_BREAK_GLASS", "1")
    fail_fast_if_destructive_flags_in_production("development")


def test_fail_fast_logs_partial_flags_in_production(
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    import logging

    monkeypatch.setenv("ALLOW_DESTRUCTIVE_DB_RESET", "1")
    monkeypatch.delenv("DESTRUCTIVE_DB_RESET_BREAK_GLASS", raising=False)
    caplog.set_level(logging.ERROR, logger="core.destructive_db_guard")
    fail_fast_if_destructive_flags_in_production("production")
    assert any("ALLOW_DESTRUCTIVE_DB_RESET" in record.message for record in caplog.records)


def test_app_lifespan_fails_in_production_with_both_flags(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from fastapi.testclient import TestClient

    from app.main import create_app

    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("BOT_AUTH_ENABLED", "true")
    monkeypatch.setenv("BOT_TELEGRAM_ALLOWLIST", "1:admin")
    monkeypatch.setenv("ALLOW_DESTRUCTIVE_DB_RESET", "1")
    monkeypatch.setenv("DESTRUCTIVE_DB_RESET_BREAK_GLASS", "1")
    get_settings.cache_clear()

    with pytest.raises(DestructiveDbFlagsMisconfigured):
        with TestClient(create_app()):
            pass
