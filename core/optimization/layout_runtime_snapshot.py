#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Снимок окружения для визуализации раскладки (optimize → layout): **замороженная копия**
плана OPT и срез заказа из cfg.

План берётся из текущего контекста (`core.optimization.context`: TLS и/или
``ContextVar``) и при сборке снимка **глубоко копируется** в read-only структуры —
последующие записи в OPT_* не меняют уже переданный в layout ``LayoutRuntimeSnapshot``.

**Остаётся в рантайме заказа** (`core.plate_runtime_state`): списки плит и
`PLATE_LOAD_DETAILS` — потоколокально / ``ContextVar``; ``core.config_and_data``
реэкспортирует те же имена через прокси модуля.

Composition root (бот, CLI, тест) после оптимизации собирает снимок через
`build_layout_runtime_snapshot(...)` и передаёт его в layout (см. DIP-003).
"""

from __future__ import annotations

import copy
import types
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from types import MappingProxyType
from typing import Any

from core.optimization.context import (
    LOAD_TO_REINFORCEMENT_MAP,
    OPT_CASCADING_PLAN,
    OPT_CASCADING_PLAN_BY_LOAD,
    OPT_PLAN,
    OPT_WIDTH_PRIORITY,
)
from core.optimization.optimization_config import OptimizationConfig

PlateLoadDetailsMap = Mapping[tuple[float, float, int, str], int]

def _freeze_plate_load_details(
    src: PlateLoadDetailsMap | dict[tuple[float, float, int, str], int],
) -> Mapping[tuple[float, float, int, str], int]:
    return MappingProxyType(dict(src))


def _tuple_floats(values: list[float] | tuple[float, ...]) -> tuple[float, ...]:
    return tuple(values)


def _make_get_load_code_for_plate(
    plate_load_details: Mapping[tuple[float, float, int, str], int],
) -> Callable[[float, float, int], int]:
    """Логика как в config_and_data.get_load_code_for_plate, но по переданной карте."""

    def get_load_code_for_plate(length_m: float, width_m: float, default: int = 8) -> int:
        try:
            key_base = (round(float(length_m), 3), round(float(width_m), 3))
        except Exception:
            return 6 if (isinstance(width_m, (int, float)) and float(width_m) < 1.0) else default

        matching_loads: list[tuple[int, int]] = []
        for key, qty in plate_load_details.items():
            l_m, w_m, load = key[0], key[1], key[2]
            if abs(l_m - key_base[0]) <= 0.005 and abs(w_m - key_base[1]) <= 0.005:
                matching_loads.append((load, qty))
        if matching_loads:
            return max(matching_loads, key=lambda x: x[1])[0]

        try:
            w_val = float(width_m)
        except Exception:
            w_val = 1.2
        if w_val < 1.0:
            return 6
        return default

    return get_load_code_for_plate


@dataclass(frozen=True, slots=True)
class OptPlanFrozenSnapshot:
    """
    Неизменяемый снимок OPT на момент вызова ``capture_from_context`` / ``build_layout_runtime_snapshot``.
    Не алиасирует TLS: дальнейшие присваивания ``core.optimization.OPT_*`` не видны здесь.
    """

    opt_plan: Mapping[Any, Any]
    opt_cascading_plan: Mapping[Any, Any]
    opt_cascading_plan_by_load: Mapping[Any, Any]
    opt_width_priority: tuple[Any, ...]
    load_to_reinforcement_map: Mapping[Any, Any]

    @classmethod
    def capture_from_context(cls) -> OptPlanFrozenSnapshot:
        return cls(
            opt_plan=MappingProxyType(copy.deepcopy(dict(OPT_PLAN))),
            opt_cascading_plan=MappingProxyType(copy.deepcopy(dict(OPT_CASCADING_PLAN))),
            opt_cascading_plan_by_load=MappingProxyType(copy.deepcopy(dict(OPT_CASCADING_PLAN_BY_LOAD))),
            opt_width_priority=tuple(copy.deepcopy(list(OPT_WIDTH_PRIORITY))),
            load_to_reinforcement_map=MappingProxyType(copy.deepcopy(dict(LOAD_TO_REINFORCEMENT_MAP))),
        )

    @classmethod
    def from_context(cls) -> OptPlanFrozenSnapshot:
        return cls.capture_from_context()


OptPlanTlsHandles = OptPlanFrozenSnapshot


@dataclass(frozen=True, slots=True)
class LayoutPlateListsReadOnly:
    """Кортежи длин по группам ширин (копия на момент снимка)."""

    plates_1_2: tuple[float, ...]
    plates_1_08: tuple[float, ...]
    plates_1_5_to_1_2: tuple[float, ...]
    plates_1_0: tuple[float, ...]
    plates_0_46: tuple[float, ...]
    plates_0_32: tuple[float, ...]
    plates_0_72: tuple[float, ...]
    plates_0_70: tuple[float, ...]
    plates_0_86: tuple[float, ...]
    plates_0_74: tuple[float, ...]
    plates_0_88: tuple[float, ...]
    plates_0_48: tuple[float, ...]
    plates_0_50: tuple[float, ...]
    plates_0_34: tuple[float, ...]

    @classmethod
    def from_config_module(cls, cfg: types.ModuleType) -> LayoutPlateListsReadOnly:
        return cls(
            plates_1_2=_tuple_floats(cfg.PLATES_1_2),
            plates_1_08=_tuple_floats(cfg.PLATES_1_08),
            plates_1_5_to_1_2=_tuple_floats(cfg.PLATES_1_5_TO_1_2),
            plates_1_0=_tuple_floats(cfg.PLATES_1_0),
            plates_0_46=_tuple_floats(cfg.PLATES_0_46),
            plates_0_32=_tuple_floats(cfg.PLATES_0_32),
            plates_0_72=_tuple_floats(cfg.PLATES_0_72),
            plates_0_70=_tuple_floats(cfg.PLATES_0_70),
            plates_0_86=_tuple_floats(cfg.PLATES_0_86),
            plates_0_74=_tuple_floats(cfg.PLATES_0_74),
            plates_0_88=_tuple_floats(cfg.PLATES_0_88),
            plates_0_48=_tuple_floats(cfg.PLATES_0_48),
            plates_0_50=_tuple_floats(cfg.PLATES_0_50),
            plates_0_34=_tuple_floats(cfg.PLATES_0_34),
        )


@dataclass(frozen=True, slots=True)
class LayoutSequenceCfgSlice:
    """
    Срез `cfg`, используемый в layout_sequence: данные заказа (read-only) и
    вызываемые хелперы. `get_load_code_for_plate` привязан к `plate_load_details`
    снимка, а не к глобальному PLATE_LOAD_DETAILS.
    """

    plate_load_details: PlateLoadDetailsMap
    plate_lists: LayoutPlateListsReadOnly
    normalize_load_code: Callable[..., Any]
    make_plate_name: Callable[..., Any]
    format_reinforcement_from_load_code: Callable[[float | int], str]
    get_load_code_for_plate: Callable[[float, float, int], int]

    @classmethod
    def from_config_module(
        cls,
        cfg: types.ModuleType,
        *,
        plate_load_details: PlateLoadDetailsMap | None = None,
        plate_lists: LayoutPlateListsReadOnly | None = None,
    ) -> LayoutSequenceCfgSlice:
        details = plate_load_details if plate_load_details is not None else cfg.PLATE_LOAD_DETAILS
        frozen_details = _freeze_plate_load_details(details)
        lists = plate_lists if plate_lists is not None else LayoutPlateListsReadOnly.from_config_module(cfg)
        return cls(
            plate_load_details=frozen_details,
            plate_lists=lists,
            normalize_load_code=cfg.normalize_load_code,
            make_plate_name=cfg.make_plate_name,
            format_reinforcement_from_load_code=cfg.format_reinforcement_from_load_code,
            get_load_code_for_plate=_make_get_load_code_for_plate(frozen_details),
        )


@dataclass(frozen=True, slots=True)
class LayoutRuntimeSnapshot:
    """Единый объект для передачи в layout после оптимизации (см. DIP-003)."""

    opt_snapshot: OptPlanFrozenSnapshot
    layout_cfg: LayoutSequenceCfgSlice
    optimization_config: OptimizationConfig | None


def build_layout_runtime_snapshot(
    *,
    cfg: types.ModuleType | None = None,
    optimization_config: OptimizationConfig | None = None,
    plate_load_details: PlateLoadDetailsMap | dict[tuple[float, float, int, str], int] | None = None,
    plate_lists: LayoutPlateListsReadOnly | None = None,
    opt_snapshot: OptPlanFrozenSnapshot | None = None,
) -> LayoutRuntimeSnapshot:
    """
    Собрать снимок для раскладки. Вызывать из composition root после заполнения OPT TLS
    и актуального заказа в cfg (или передать переопределения для тестов/изоляции).

    :param cfg: модуль `core.config_and_data`; по умолчанию импортируется.
    :param optimization_config: необязательный `OptimizationConfig` текущего прогона оптимизации.
    :param plate_load_details: переопределение карты нагрузок; иначе берётся из cfg.
    :param plate_lists: переопределение списков длин; иначе из cfg.
    :param opt_snapshot: явный снимок плана (тесты); иначе ``OptPlanFrozenSnapshot.capture_from_context()``.
    """
    if cfg is None:
        import core.config_and_data as _cfg

        cfg = _cfg

    snap = opt_snapshot if opt_snapshot is not None else OptPlanFrozenSnapshot.capture_from_context()
    layout_slice = LayoutSequenceCfgSlice.from_config_module(
        cfg,
        plate_load_details=plate_load_details,
        plate_lists=plate_lists,
    )
    return LayoutRuntimeSnapshot(
        opt_snapshot=snap,
        layout_cfg=layout_slice,
        optimization_config=optimization_config,
    )
