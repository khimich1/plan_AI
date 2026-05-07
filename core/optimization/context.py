#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Потоколокальное хранилище глобального состояния оптимизатора (OPT-005).

Мутабельные «глобали» вида OPT_PLAN раньше были общими на весь процесс и при
параллельных вызовах давали гонки. Здесь у каждого потока свой набор dict/list.
Снаружи по-прежнему используются имена OPT_PLAN, … — см. прокси и
core.optimization._OptimizationModule.__setattr__.
"""

from __future__ import annotations

import copy
import threading
from collections.abc import Iterator, MutableMapping, MutableSequence
from typing import Any


_tls = threading.local()


def _ensure() -> dict[str, Any]:
    if not hasattr(_tls, "state"):
        _tls.state = {
            "opt_plan": {},
            "opt_cascading_plan": {},
            "opt_cascading_plan_by_load": {},
            "opt_width_priority": [],
            "load_to_reinforcement_map": {},
        }
    return _tls.state


def _assign_dict(target: dict, value: Any) -> None:
    target.clear()
    if value is None:
        return
    if isinstance(value, dict):
        target.update(value)
    else:
        target.update(dict(value))


def _assign_list(target: list, value: Any) -> None:
    target.clear()
    if value is None:
        return
    if isinstance(value, list):
        target.extend(value)
    else:
        target.extend(list(value))  # type: ignore[arg-type]


def tls_set(name: str, value: Any) -> None:
    """Установить один из публичных атрибутов модуля оптимизации (для __setattr__)."""
    s = _ensure()
    if name == "OPT_PLAN":
        _assign_dict(s["opt_plan"], value)
    elif name == "OPT_CASCADING_PLAN":
        _assign_dict(s["opt_cascading_plan"], value)
    elif name == "OPT_CASCADING_PLAN_BY_LOAD":
        _assign_dict(s["opt_cascading_plan_by_load"], value)
    elif name == "OPT_WIDTH_PRIORITY":
        _assign_list(s["opt_width_priority"], value)
    elif name == "LOAD_TO_REINFORCEMENT_MAP":
        _assign_dict(s["load_to_reinforcement_map"], value)
    else:
        raise KeyError(name)


class _ThreadLocalDictProxy(MutableMapping):
    """Прокси к потоколокальному dict; один экземпляр на весь процесс."""

    __slots__ = ("_key",)

    def __init__(self, state_key: str) -> None:
        self._key = state_key

    def _d(self) -> dict:
        return _ensure()[self._key]

    def __getitem__(self, key: Any) -> Any:
        return self._d().__getitem__(key)

    def __setitem__(self, key: Any, value: Any) -> None:
        self._d().__setitem__(key, value)

    def __delitem__(self, key: Any) -> None:
        self._d().__delitem__(key)

    def __iter__(self) -> Iterator:
        return iter(self._d())

    def __len__(self) -> int:
        return len(self._d())

    def __bool__(self) -> bool:
        return bool(self._d())

    def __repr__(self) -> str:
        return f"<ThreadLocalDictProxy {self._key} {self._d()!r}>"

    def __copy__(self) -> dict:
        return self._d().copy()

    def __deepcopy__(self, memo: dict) -> dict:
        return copy.deepcopy(self._d(), memo)

    def clear(self) -> None:
        self._d().clear()

    def update(self, *args: Any, **kwargs: Any) -> None:
        self._d().update(*args, **kwargs)

    def get(self, key: Any, default: Any = None) -> Any:
        return self._d().get(key, default)

    def setdefault(self, key: Any, default: Any = None) -> Any:
        return self._d().setdefault(key, default)

    def keys(self):
        return self._d().keys()

    def values(self):
        return self._d().values()

    def items(self):
        return self._d().items()


class _ThreadLocalListProxy(MutableSequence):
    __slots__ = ("_key",)

    def __init__(self, state_key: str) -> None:
        self._key = state_key

    def _lst(self) -> list:
        return _ensure()[self._key]

    def __getitem__(self, index: Any) -> Any:
        return self._lst().__getitem__(index)

    def __setitem__(self, index: Any, value: Any) -> None:
        self._lst().__setitem__(index, value)

    def __delitem__(self, index: Any) -> None:
        self._lst().__delitem__(index)

    def __len__(self) -> int:
        return len(self._lst())

    def __iter__(self) -> Iterator:
        return iter(self._lst())

    def __bool__(self) -> bool:
        return bool(self._lst())

    def insert(self, index: int, value: Any) -> None:
        self._lst().insert(index, value)

    def clear(self) -> None:
        self._lst().clear()

    def extend(self, values: Any) -> None:
        self._lst().extend(values)

    def append(self, value: Any) -> None:
        self._lst().append(value)

    def __repr__(self) -> str:
        return f"<ThreadLocalListProxy {self._key} {self._lst()!r}>"

    def __copy__(self) -> list:
        return list(self._lst())

    def __deepcopy__(self, memo: dict) -> list:
        return copy.deepcopy(self._lst(), memo)


# Единственные экземпляры прокси (идентичность стабильна при импорте)
OPT_PLAN = _ThreadLocalDictProxy("opt_plan")
OPT_CASCADING_PLAN = _ThreadLocalDictProxy("opt_cascading_plan")
OPT_CASCADING_PLAN_BY_LOAD = _ThreadLocalDictProxy("opt_cascading_plan_by_load")
OPT_WIDTH_PRIORITY = _ThreadLocalListProxy("opt_width_priority")
LOAD_TO_REINFORCEMENT_MAP = _ThreadLocalDictProxy("load_to_reinforcement_map")
