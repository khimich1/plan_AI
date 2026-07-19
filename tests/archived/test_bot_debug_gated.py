"""Bot debug helpers must not write NDJSON when agent debug is disabled."""

from __future__ import annotations

import pytest

from bot.handlers.debug_util import write_agent_debug, write_agent_debug_session
from core.debug_paths import kp_db_agent_debug_active


@pytest.fixture(autouse=True)
def _disable_agent_debug(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("KP_DB_AGENT_DEBUG", raising=False)
    monkeypatch.delenv("APP_DEBUG", raising=False)
    monkeypatch.setenv("APP_ENV", "test")
    try:
        from core.config.settings import get_settings

        monkeypatch.setattr(get_settings(), "app_debug", False)
    except Exception:
        pass


def test_kp_db_agent_debug_inactive() -> None:
    assert kp_db_agent_debug_active() is False


def test_write_agent_debug_does_not_create_file(tmp_path, monkeypatch: pytest.MonkeyPatch) -> None:
    log_path = tmp_path / "bot-debug.log"
    monkeypatch.chdir(tmp_path)
    write_agent_debug(log_path, {"message": "secret", "kp_id": 1})
    assert not log_path.exists()


def test_write_agent_debug_session_does_not_create_file(tmp_path) -> None:
    log_path = tmp_path / "bot-session.log"
    write_agent_debug_session(
        log_path,
        run_id="run1",
        hypothesis_id="H1",
        location="bot:test",
        message="test",
        data={"kp_id": 2},
    )
    assert not log_path.exists()
