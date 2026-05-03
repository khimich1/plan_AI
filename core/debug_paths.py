from __future__ import annotations

from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEBUG_LOGS_DIR = PROJECT_ROOT / "debug_logs"


def get_debug_log_path(filename: str) -> Path:
    """Return path inside debug_logs and ensure directory exists."""
    DEBUG_LOGS_DIR.mkdir(parents=True, exist_ok=True)
    return DEBUG_LOGS_DIR / filename
