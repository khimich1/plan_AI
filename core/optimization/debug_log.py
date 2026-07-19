#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Shared debug logging helpers for the optimization package (no ILP / PuLP).

Implementation-specific paths and session writers live in
``optimization_debug_impl`` (imported only from ``_implementation``).
"""

from __future__ import annotations

from core.config.logging import optimization_debug_active
from core.debug_paths import get_debug_log_path

_DEBUG_LOG_COMMON = get_debug_log_path("debug.log")
_DEBUG_LOG_5b5324 = get_debug_log_path("debug-5b5324.log")


def _opt_debug_enabled() -> bool:
    """
    Включены ли подробные debug-логи оптимизатора.

    См. :func:`core.config.logging.optimization_debug_active` (OPT_DEBUG_LOG или DEBUG
    для логгера ``core.optimization``).
    """
    return optimization_debug_active()


class _DbgNullFile:
    """
    No-op file handle: когда отладка оптимизатора выключена
    (см. :func:`optimization_debug_active`). Поддерживает контекст-менеджер и `.write`.
    """

    def write(self, *_args, **_kwargs):
        return 0

    def __enter__(self):
        return self

    def __exit__(self, *_args, **_kwargs):
        return False


_DBG_NULL_FILE = _DbgNullFile()


def _dbg_open_append(path):
    """
    Append-handle для debug-логов оптимизатора.
    Когда :func:`optimization_debug_active` даёт False — no-op handle,
    без открытия файла. При ошибке открытия — тоже no-op.
    """
    if not _opt_debug_enabled():
        return _DBG_NULL_FILE
    try:
        return open(path, "a", encoding="utf-8")
    except Exception:
        return _DBG_NULL_FILE
