"""Страж неявного заказа плит: PLATE_RUNTIME_STRICT=warn|raise."""

from __future__ import annotations

import inspect
import logging
import os
import threading

import pytest

from core.plate_order_context import PlateOrderContext
from core.plate_runtime_state import (
    _tls,
    get_plate_mutable_runtime,
    new_plate_mutable_runtime_empty,
    plate_mutable_runtime_scope,
)

_LOGGER_NAME = "core.plate_runtime_state"


class _ListHandler(logging.Handler):
    def __init__(self) -> None:
        super().__init__(level=logging.WARNING)
        self.records: list[logging.LogRecord] = []

    def emit(self, record: logging.LogRecord) -> None:
        self.records.append(record)


@pytest.fixture
def strict_mode(monkeypatch: pytest.MonkeyPatch):
    def _set(mode: str) -> None:
        monkeypatch.setenv("PLATE_RUNTIME_STRICT", mode)

    return _set


@pytest.fixture
def runtime_log_handler():
    logger = logging.getLogger(_LOGGER_NAME)
    handler = _ListHandler()
    previous = logger.level
    logger.addHandler(handler)
    logger.setLevel(logging.WARNING)
    try:
        yield handler
    finally:
        logger.removeHandler(handler)
        logger.setLevel(previous)


def _run_on_fresh_thread(fn) -> None:
    thread = threading.Thread(target=fn)
    thread.start()
    thread.join()
    assert not thread.is_alive()


def test_raise_does_not_create_thread_runtime_and_sequential_reads_do_not_share_lists(
    strict_mode,
) -> None:
    strict_mode("raise")
    outcome: dict[str, object] = {}

    def _worker() -> None:
        lists: list[list[float]] = []
        errors: list[BaseException] = []
        for _ in range(2):
            try:
                lists.append(get_plate_mutable_runtime().plates_1_2)
            except RuntimeError as exc:
                errors.append(exc)
        outcome["lists"] = lists
        outcome["errors"] = errors
        outcome["has_runtime"] = hasattr(_tls, "runtime")

    _run_on_fresh_thread(_worker)

    errors = outcome["errors"]
    assert isinstance(errors, list)
    assert outcome["lists"] == []
    assert len(errors) == 2
    assert outcome["has_runtime"] is False
    for exc in errors:
        message = str(exc)
        assert "заказ плит не привязан" in message
        assert __file__ in message


def test_warn_returns_same_thread_runtime_and_logs_caller(
    strict_mode,
    runtime_log_handler: _ListHandler,
) -> None:
    strict_mode("warn")
    outcome: dict[str, object] = {}

    def _worker() -> None:
        caller_line = inspect.currentframe().f_lineno + 1
        first = get_plate_mutable_runtime()
        second = get_plate_mutable_runtime()
        first.plates_1_2.append(4.2)
        outcome["same"] = first is second
        outcome["seen"] = list(second.plates_1_2)
        outcome["has_runtime"] = hasattr(_tls, "runtime")
        outcome["tls_is_first"] = getattr(_tls, "runtime", None) is first
        outcome["caller_line"] = caller_line

    _run_on_fresh_thread(_worker)

    assert outcome["same"] is True
    assert outcome["seen"] == [4.2]
    assert outcome["has_runtime"] is True
    assert outcome["tls_is_first"] is True
    caller_line = outcome["caller_line"]
    assert isinstance(caller_line, int)
    matched = [
        record
        for record in runtime_log_handler.records
        if record.pathname == __file__
        and record.lineno == caller_line
        and f"{__file__}:{caller_line}" in record.getMessage()
        and "заказ плит не привязан" in record.getMessage()
    ]
    assert matched


def test_bound_scope_does_not_warn(
    strict_mode,
    runtime_log_handler: _ListHandler,
) -> None:
    strict_mode("warn")
    ctx = PlateOrderContext.fresh_empty()
    with ctx.bound():
        assert get_plate_mutable_runtime() is ctx.plates
    with plate_mutable_runtime_scope(new_plate_mutable_runtime_empty()) as runtime:
        assert get_plate_mutable_runtime() is runtime
    assert runtime_log_handler.records == []


def test_bound_scope_does_not_raise_when_strict(strict_mode) -> None:
    strict_mode("raise")
    ctx = PlateOrderContext.fresh_empty()
    ctx.plates.plates_1_2.append(1.25)
    with ctx.bound():
        assert get_plate_mutable_runtime().plates_1_2 == [1.25]
    with plate_mutable_runtime_scope(ctx.plates):
        assert get_plate_mutable_runtime() is ctx.plates


def test_default_mode_is_raise(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("PLATE_RUNTIME_STRICT", raising=False)
    assert "PLATE_RUNTIME_STRICT" not in os.environ
    outcome: dict[str, object] = {}

    def _worker() -> None:
        errors: list[BaseException] = []
        for _ in range(2):
            try:
                get_plate_mutable_runtime()
            except RuntimeError as exc:
                errors.append(exc)
        outcome["errors"] = errors
        outcome["has_runtime"] = hasattr(_tls, "runtime")

    _run_on_fresh_thread(_worker)

    errors = outcome["errors"]
    assert isinstance(errors, list)
    assert len(errors) == 2
    assert outcome["has_runtime"] is False
    assert all("заказ плит не привязан" in str(exc) for exc in errors)
