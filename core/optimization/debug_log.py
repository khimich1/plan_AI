#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Shared debug logging helpers for the optimization package (no ILP / PuLP).

Implementation-specific paths and session writers live in
``optimization_debug_impl`` (imported only from ``_implementation``).
"""

from __future__ import annotations

import os as _os

from core.debug_paths import get_debug_log_path

_DEBUG_LOG_COMMON = get_debug_log_path("debug.log")
_DEBUG_LOG_5b5324 = get_debug_log_path("debug-5b5324.log")


def _opt_debug_enabled() -> bool:
    """
    Включены ли подробные debug-логи оптимизатора.
    По умолчанию выключены: дебаг-регионы пишут в файлы только при OPT_DEBUG_LOG=1.
    Это нужно для честных замеров и чтобы prod не засорял диск.
    """
    return _os.environ.get("OPT_DEBUG_LOG", "").strip() in ("1", "true", "True", "yes", "on")


class _DbgNullFile:
    """
    No-op file handle: используется как заглушка для debug-логов,
    когда OPT_DEBUG_LOG выключен. Поддерживает и контекст-менеджер,
    и прямой `.write(...)` без `with`.
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
    При OPT_DEBUG_LOG=0 возвращает no-op handle, поэтому никакая запись не идёт.
    Никогда не бросает исключений: при ошибке открытия — тоже no-op.
    """
    if not _opt_debug_enabled():
        return _DBG_NULL_FILE
    try:
        return open(path, "a", encoding="utf-8")
    except Exception:
        return _DBG_NULL_FILE
