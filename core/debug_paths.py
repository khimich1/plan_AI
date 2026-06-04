from __future__ import annotations

import json
import os
import time
from pathlib import Path

_TRUTHY = frozenset({"1", "true", "yes", "on"})

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEBUG_LOGS_DIR = PROJECT_ROOT / "debug_logs"


def get_debug_log_path(filename: str) -> Path:
    """Return path inside debug_logs and ensure directory exists."""
    DEBUG_LOGS_DIR.mkdir(parents=True, exist_ok=True)
    return DEBUG_LOGS_DIR / filename


def kp_db_agent_debug_active() -> bool:
    """
    Whether agent NDJSON logs for kp_db hot paths should be written.

    Enabled when APP_DEBUG / KP_DB_AGENT_DEBUG is truthy, or when
    ``core.config.settings.Settings.app_debug`` is True (if settings load).
    """
    if os.environ.get("KP_DB_AGENT_DEBUG", "").strip().lower() in _TRUTHY:
        return True
    if os.environ.get("APP_DEBUG", "").strip().lower() in _TRUTHY:
        return True
    try:
        from core.config.settings import get_settings

        if get_settings().app_debug:
            return True
    except Exception:
        pass
    return False


def append_agent_debug_log(path: Path | str, payload: dict) -> None:
    """Append one NDJSON line when :func:`kp_db_agent_debug_active` is True."""
    if not kp_db_agent_debug_active():
        return
    try:
        line = json.dumps(payload, ensure_ascii=False) + "\n"
        with open(path, "a", encoding="utf-8") as handle:
            handle.write(line)
    except Exception:
        pass


def append_agent_debug_session(
    path: Path | str,
    *,
    run_id: str,
    hypothesis_id: str,
    location: str,
    message: str,
    data: dict,
    session_id: str = "d7e22e",
) -> None:
    """Session-style NDJSON row (legacy ``_debug_session_write`` format)."""
    append_agent_debug_log(
        path,
        {
            "sessionId": session_id,
            "runId": run_id,
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data,
            "timestamp": int(time.time() * 1000),
        },
    )
