#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Мутабельное состояние заказа плит (SSOT на время сессии расчёта).

PLATE-CTX-001: изоляция от гонок между потоками (threading.local) и опционально
между asyncio-задачами (contextvars): при установленном ContextVar он имеет приоритет.

Оркестрация (бот, FastAPI) может обернуть прогон в::

    token = bind_plate_mutable_runtime(runtime)
    try:
        ...
    finally:
        reset_plate_mutable_runtime(token)

Если bind не вызывается, поведение как раньше: отдельное состояние на поток.
"""

from __future__ import annotations

import contextvars
import threading
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import Any, Dict, Iterator, List, Optional, Tuple, Union

PlateLoadKey = Tuple[float, float, int, str]
PlateExactKey = Tuple[float, str]


@dataclass
class PlateMutableRuntime:
    """Все бывшие глобальные списки/карты заказа из config_and_data (мутабельно)."""

    plates_1_2: List[float] = field(default_factory=list)
    plates_1_5_to_1_2: List[float] = field(default_factory=list)
    plates_1_0: List[float] = field(default_factory=list)
    plates_1_08: List[float] = field(default_factory=list)
    plates_0_46: List[float] = field(default_factory=list)
    plates_0_32: List[float] = field(default_factory=list)
    plates_0_72: List[float] = field(default_factory=list)
    plates_0_70: List[float] = field(default_factory=list)
    plates_0_86: List[float] = field(default_factory=list)
    plates_0_74: List[float] = field(default_factory=list)
    plates_0_88: List[float] = field(default_factory=list)
    plates_0_48: List[float] = field(default_factory=list)
    plates_0_50: List[float] = field(default_factory=list)
    plates_0_34: List[float] = field(default_factory=list)

    longitudinal_cuts: int = 0
    length_trims: int = 0
    unused_strips_0_3_m_total: float = 0.0
    scrap_strips_0_2_m_total: float = 0.0
    usable_strips_0_74_m_total: float = 0.0
    usable_strips_0_88_m_total: float = 0.0
    usable_strips_0_48_m_total: float = 0.0
    usable_strips_0_50_m_total: float = 0.0
    usable_strips_0_34_m_total: float = 0.0
    scrap_strips_0_12_m_total: float = 0.0
    waste_area_m2: float = 0.0

    plate_load_details: Dict[PlateLoadKey, int] = field(default_factory=dict)
    plate_length_dm_raw: Dict[PlateLoadKey, str] = field(default_factory=dict)
    plate_exact_widths: Dict[PlateExactKey, float] = field(default_factory=dict)
    plate_metadata: Dict[Tuple[float, int], List[Dict[str, Any]]] = field(default_factory=dict)
    plate_max_reinforcement_map: Dict[Tuple[float, int], float] = field(default_factory=dict)
    plate_nomenclature_cache: Dict[PlateLoadKey, Dict[str, Any]] = field(default_factory=dict)
    last_parse_diagnostics: List[Dict[str, Any]] = field(default_factory=list)

    @classmethod
    def factory_demo_order(cls) -> PlateMutableRuntime:
        """Начальные данные как в legacy-модуле до парсинга (демо КЗ-план)."""
        plates_1_2 = [3.39] * 2
        plates_1_08 = plates_1_2.copy()
        plates_0_32 = [6.63] * 4 + [7.83] * 3
        plates_0_72 = [5.63] * 5
        plates_0_70 = [4.65] * 5
        plates_0_86 = [6.75] * 2 + [4.65] * 5
        rt = cls(
            plates_1_2=plates_1_2,
            plates_1_08=plates_1_08,
            plates_0_32=plates_0_32,
            plates_0_72=plates_0_72,
            plates_0_70=plates_0_70,
            plates_0_86=plates_0_86,
        )
        rt._recompute_totals_from_lists()
        return rt

    def _recompute_totals_from_lists(self) -> None:
        self.longitudinal_cuts = (
            len(self.plates_1_5_to_1_2)
            + len(self.plates_1_0)
            + len(self.plates_1_08)
            + len(self.plates_0_46)
            + len(self.plates_0_32)
            + len(self.plates_0_72)
            + len(self.plates_0_70)
            + len(self.plates_0_86)
        )
        self.length_trims = 0
        self.unused_strips_0_3_m_total = 0.0
        self.scrap_strips_0_2_m_total = 0.0
        self.usable_strips_0_74_m_total = round(sum(self.plates_0_46), 1)
        self.usable_strips_0_88_m_total = round(sum(self.plates_0_32), 1)
        self.usable_strips_0_48_m_total = round(sum(self.plates_0_72), 1)
        self.usable_strips_0_50_m_total = round(sum(self.plates_0_70), 1)
        self.usable_strips_0_34_m_total = round(sum(self.plates_0_86), 1)
        self.scrap_strips_0_12_m_total = round(sum(self.plates_1_08), 1)
        self.waste_area_m2 = round(0.12 * self.scrap_strips_0_12_m_total, 2)

    def clear_plate_lists(self) -> None:
        self.plates_1_2 = []
        self.plates_1_5_to_1_2 = []
        self.plates_1_0 = []
        self.plates_1_08 = []
        self.plates_0_46 = []
        self.plates_0_32 = []
        self.plates_0_72 = []
        self.plates_0_70 = []
        self.plates_0_86 = []
        self.plates_0_74 = []
        self.plates_0_88 = []
        self.plates_0_48 = []
        self.plates_0_50 = []
        self.plates_0_34 = []
        self.plate_load_details.clear()
        self.plate_exact_widths.clear()
        self.plate_length_dm_raw.clear()
        self.plate_max_reinforcement_map.clear()
        self.plate_nomenclature_cache.clear()


_plate_cv: contextvars.ContextVar[Optional[PlateMutableRuntime]] = contextvars.ContextVar(
    "plate_mutable_runtime", default=None
)
_tls = threading.local()


def get_plate_mutable_runtime() -> PlateMutableRuntime:
    """Текущее мутабельное состояние заказа (контекст задачи или поток)."""
    ctx_rt = _plate_cv.get()
    if ctx_rt is not None:
        return ctx_rt
    if not hasattr(_tls, "runtime"):
        # Пустой рантайм по умолчанию: демо-заказ только через
        # ``new_plate_mutable_runtime_from_demo()`` (избегаем утечек «чужого» заказа в пуле потоков / при первом касании TLS).
        _tls.runtime = new_plate_mutable_runtime_empty()
    return _tls.runtime


def bind_plate_mutable_runtime(runtime: PlateMutableRuntime) -> contextvars.Token[Optional[PlateMutableRuntime]]:
    """Привязать состояние к текущей asyncio-логической задаче (вернуть token для reset)."""
    return _plate_cv.set(runtime)


def reset_plate_mutable_runtime(token: contextvars.Token[Optional[PlateMutableRuntime]]) -> None:
    _plate_cv.reset(token)


def new_plate_mutable_runtime_empty() -> PlateMutableRuntime:
    """Пустое состояние (после clear), без демо-заказа."""
    rt = PlateMutableRuntime()
    return rt


def new_plate_mutable_runtime_from_demo() -> PlateMutableRuntime:
    return PlateMutableRuntime.factory_demo_order()


@contextmanager
def plate_mutable_runtime_scope(runtime: PlateMutableRuntime) -> Iterator[PlateMutableRuntime]:
    """Контекстный менеджер для явной изоляции прогона (async/sync)."""
    tok = bind_plate_mutable_runtime(runtime)
    try:
        yield runtime
    finally:
        reset_plate_mutable_runtime(tok)


@contextmanager
def fresh_plate_mutable_request_scope() -> Iterator[PlateMutableRuntime]:
    """Пустое состояние заказа на время одного HTTP-запроса / апдейта бота (S1/PLATE-CTX-001)."""
    with plate_mutable_runtime_scope(new_plate_mutable_runtime_empty()) as rt:
        yield rt


# Имена верхнего уровня legacy-модуля cfg -> атрибут PlateMutableRuntime
MUTABLE_ATTR_MAP: Dict[str, str] = {
    "PLATES_1_2": "plates_1_2",
    "PLATES_1_5_TO_1_2": "plates_1_5_to_1_2",
    "PLATES_1_0": "plates_1_0",
    "PLATES_1_08": "plates_1_08",
    "PLATES_0_46": "plates_0_46",
    "PLATES_0_32": "plates_0_32",
    "PLATES_0_72": "plates_0_72",
    "PLATES_0_70": "plates_0_70",
    "PLATES_0_86": "plates_0_86",
    "PLATES_0_74": "plates_0_74",
    "PLATES_0_88": "plates_0_88",
    "PLATES_0_48": "plates_0_48",
    "PLATES_0_50": "plates_0_50",
    "PLATES_0_34": "plates_0_34",
    "LONGITUDINAL_CUTS": "longitudinal_cuts",
    "LENGTH_TRIMS": "length_trims",
    "UNUSED_STRIPS_0_3_M_TOTAL": "unused_strips_0_3_m_total",
    "SCRAP_STRIPS_0_2_M_TOTAL": "scrap_strips_0_2_m_total",
    "USABLE_STRIPS_0_74_M_TOTAL": "usable_strips_0_74_m_total",
    "USABLE_STRIPS_0_88_M_TOTAL": "usable_strips_0_88_m_total",
    "USABLE_STRIPS_0_48_M_TOTAL": "usable_strips_0_48_m_total",
    "USABLE_STRIPS_0_50_M_TOTAL": "usable_strips_0_50_m_total",
    "USABLE_STRIPS_0_34_M_TOTAL": "usable_strips_0_34_m_total",
    "SCRAP_STRIPS_0_12_M_TOTAL": "scrap_strips_0_12_m_total",
    "WASTE_AREA_M2": "waste_area_m2",
    "PLATE_LOAD_DETAILS": "plate_load_details",
    "PLATE_EXACT_WIDTHS": "plate_exact_widths",
    "PLATE_LENGTH_DM_RAW": "plate_length_dm_raw",
    "PLATE_METADATA": "plate_metadata",
    "PLATE_MAX_REINFORCEMENT_MAP": "plate_max_reinforcement_map",
    "PLATE_NOMENCLATURE_CACHE": "plate_nomenclature_cache",
    "LAST_PARSE_DIAGNOSTICS": "last_parse_diagnostics",
}

MUTABLE_LEGACY_NAMES = frozenset(MUTABLE_ATTR_MAP.keys())

__all__ = [
    "PlateLoadKey",
    "PlateExactKey",
    "PlateMutableRuntime",
    "MUTABLE_ATTR_MAP",
    "MUTABLE_LEGACY_NAMES",
    "get_plate_mutable_runtime",
    "bind_plate_mutable_runtime",
    "reset_plate_mutable_runtime",
    "new_plate_mutable_runtime_empty",
    "new_plate_mutable_runtime_from_demo",
    "plate_mutable_runtime_scope",
    "fresh_plate_mutable_request_scope",
]
