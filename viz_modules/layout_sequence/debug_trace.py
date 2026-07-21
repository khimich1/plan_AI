# -*- coding: utf-8 -*-
"""Пути NDJSON-трасс и запись строк; контекст для _agent_seq_debug (contextvars для потоков/async)."""

from __future__ import annotations

import json
import logging
import time
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterator

from core.debug_paths import get_debug_log_path, kp_db_agent_debug_active

_log = logging.getLogger(__name__)

@dataclass(frozen=True)
class LayoutSequenceTracePaths:
    """Пути файлов отладочных логов layout_sequence (совместимо с монолитом)."""

    debug_log: Path
    debug_95694e: Path
    debug_2d5c43: Path
    debug_7e420e: Path
    debug_ef42ae: Path

    @classmethod
    def default(cls) -> LayoutSequenceTracePaths:
        return cls(
            debug_log=get_debug_log_path("debug.log"),
            debug_95694e=get_debug_log_path("debug-95694e.log"),
            debug_2d5c43=get_debug_log_path("debug-2d5c43.log"),
            debug_7e420e=get_debug_log_path("debug-7e420e.log"),
            debug_ef42ae=get_debug_log_path("debug-ef42ae.log"),
        )


_DEFAULT_TRACES: LayoutSequenceTracePaths = LayoutSequenceTracePaths.default()
_trace_paths_ctx: ContextVar[LayoutSequenceTracePaths] = ContextVar(
    "layout_sequence_traces",
    default=_DEFAULT_TRACES,
)


@contextmanager
def layout_sequence_trace_context(traces: LayoutSequenceTracePaths) -> Iterator[None]:
    """Временная подстановка путей трасс (build_layout_sequence / _build_sequence_from_plan)."""

    token = _trace_paths_ctx.set(traces)
    try:
        yield
    finally:
        _trace_paths_ctx.reset(token)


def append_json_line(path: Path, payload_dict: dict[str, Any], *, ensure_ascii: bool = False) -> None:
    if not kp_db_agent_debug_active():
        return
    line = json.dumps(payload_dict, ensure_ascii=ensure_ascii, default=str) + "\n"
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(line)


def _agent_seq_debug(hypothesis_id: str, message: str, data: dict[str, Any]) -> None:
    # #region agent log
    if not kp_db_agent_debug_active():
        return
    try:
        payload = {
            "sessionId": "7e420e",
            "hypothesisId": hypothesis_id,
            "location": "layout_sequence._build_sequence_from_plan",
            "message": message,
            "data": data,
            "timestamp": int(time.time() * 1000),
        }
        append_json_line(_trace_paths_ctx.get().debug_7e420e, payload)
    except OSError as exc:
        _log.debug("layout_sequence trace skip: %s", exc, exc_info=False)
    # #endregion
