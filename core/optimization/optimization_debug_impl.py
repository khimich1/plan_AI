#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Debug sinks used by ``_implementation.py`` (paths + session writers).

Separated from ``debug_log.py`` so shared helpers stay lightweight and nothing
here imports ``_implementation`` (avoids cycles with the orchestrator package).
"""

from __future__ import annotations

from pathlib import Path

from core.debug_paths import PROJECT_ROOT, get_debug_log_path

from core.optimization.debug_log import _opt_debug_enabled

_CORE_DIR = Path(__file__).resolve().parent.parent

# Historical paths (must stay stable for existing log collectors / agents)
_DEBUG_LOG_7E420E = _CORE_DIR / "debug-7e420e.log"
_DEBUG_LOG_EF42AE = PROJECT_ROOT / "debug-ef42ae.log"

_DEBUG_AGENT_LOG_EBB546 = get_debug_log_path("debug-ebb546.log")
_DEBUG_RUNTIME_LOG_648532 = get_debug_log_path("debug-648532.log")
_DEBUG_LOG_2D5C43 = get_debug_log_path("debug-2d5c43.log")

_DEBUG_RUNTIME_SESSION_ID_648532 = "648532"


def _debug_runtime_write_648532(
    run_id: str,
    hypothesis_id: str,
    location: str,
    message: str,
    data: dict,
) -> None:
    if not _opt_debug_enabled():
        return
    try:
        import json as _json
        import time as _time

        with open(_DEBUG_RUNTIME_LOG_648532, "a", encoding="utf-8") as _f:
            _f.write(
                _json.dumps(
                    {
                        "sessionId": _DEBUG_RUNTIME_SESSION_ID_648532,
                        "runId": run_id,
                        "hypothesisId": hypothesis_id,
                        "location": location,
                        "message": message,
                        "data": data,
                        "timestamp": int(_time.time() * 1000),
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
    except Exception:
        pass
