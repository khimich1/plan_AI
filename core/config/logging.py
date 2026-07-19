# -*- coding: utf-8 -*-
"""
Уровни логирования для пакета оптимизации и связанного отладочного вывода.

Не задавайте уровни «вручную» в модулях оптимизатора: используйте переменные окружения
или вызов ``configure_optimization_logging_from_env`` после ``setup_logging``.
"""

from __future__ import annotations

import logging
import os

# Единый логгер для core/optimization/* (ILP, дорожки, атрибуция КП).
OPTIMIZATION_LOGGER_NAME = "core.optimization"

_TRUE_ENV = frozenset({"1", "true", "yes", "on"})


def get_optimization_logger() -> logging.Logger:
    return logging.getLogger(OPTIMIZATION_LOGGER_NAME)


def optimization_debug_active() -> bool:
    """
    Включена ли подробная отладка оптимизатора (файловые debug-логи + тяжёлые JSON-снимки).

    True, если:
    - ``OPT_DEBUG_LOG`` установлен в truthy-значение (обратная совместимость), или
    - у логгера ``core.optimization`` эффективно включён уровень DEBUG.
    """
    raw = os.environ.get("OPT_DEBUG_LOG", "").strip().lower()
    if raw in _TRUE_ENV:
        return True
    return get_optimization_logger().isEnabledFor(logging.DEBUG)


def configure_optimization_logging_from_env() -> None:
    """
    Выставляет уровень для ``core.optimization`` по окружению.

    Приоритет:
    1. ``OPT_DEBUG_LOG=1`` (и аналоги) → DEBUG;
    2. ``OPTIMIZATION_LOG_LEVEL`` (DEBUG/INFO/WARNING/ERROR), если задан;
    3. без изменений (наследование от root).
    """
    log = get_optimization_logger()
    opt_raw = os.environ.get("OPTIMIZATION_LOG_LEVEL", "").strip().upper()
    dbg_raw = os.environ.get("OPT_DEBUG_LOG", "").strip().lower()

    if dbg_raw in _TRUE_ENV:
        log.setLevel(logging.DEBUG)
        return

    if opt_raw:
        level = getattr(logging, opt_raw, None)
        if isinstance(level, int):
            log.setLevel(level)
