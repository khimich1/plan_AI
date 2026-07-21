"""Gated NDJSON debug helpers for Telegram handlers."""

from __future__ import annotations

from pathlib import Path

from core.debug_paths import append_agent_debug_log, append_agent_debug_session


def write_agent_debug(path: Path | str, payload: dict) -> None:
    """Append one NDJSON line when agent debug is enabled."""
    append_agent_debug_log(path, payload)


def write_agent_debug_session(
    path: Path | str,
    *,
    run_id: str,
    hypothesis_id: str,
    location: str,
    message: str,
    data: dict,
    session_id: str = "d7e22e",
) -> None:
    append_agent_debug_session(
        path,
        run_id=run_id,
        hypothesis_id=hypothesis_id,
        location=location,
        message=message,
        data=data,
        session_id=session_id,
    )
