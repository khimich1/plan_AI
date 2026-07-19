# -*- coding: utf-8 -*-
"""Агрегат заказа плит (legacy cfg.PlateOrder) и синхронизация с мутабельным рантаймом."""

from __future__ import annotations

import logging
import warnings
from dataclasses import dataclass, field
from typing import Any, Dict, List, Sequence, Tuple, Union

from ..plate_runtime_state import PlateMutableRuntime, get_plate_mutable_runtime
from ..runtime import NomenclatureCacheFiller, get_default_nomenclature_cache_filler

logger = logging.getLogger(__name__)

_APPLY_TO_GLOBALS_DEPRECATION = (
    "PlateOrder.apply_to_globals() is deprecated (A1-002); "
    "use PlateOrderContext.hydrate_from_order() with ctx.bound() or run_in_order_context()."
)
_GET_CURRENT_PLATE_ORDER_DEPRECATION = (
    "get_current_plate_order() is deprecated (A1-002); "
    "read from PlateOrderContext.plates or hydrate_from_order() instead."
)


def normalize_load_code(value, default: int = 8):
    """
    Нормализует код нагрузки к единому формату (6/8/10/12/12.5/16...).

    Примеры:
    - 800 -> 8
    - 1200 -> 12
    - 12.0 -> 12
    - "8" -> 8
    - None/ошибка -> default
    """
    if value is None:
        return default
    try:
        val = float(value)
    except Exception:
        return default

    if val >= 100:
        val = val / 100.0

    if abs(val - round(val)) < 1e-6:
        return int(round(val))

    return round(val, 1)


def parse_load_key_from_list(k_list: Sequence[Any]) -> Tuple[float, float, Union[int, float], str]:
    """Конвертирует список [length, width, load_code] или [length, width, load_code, ldr] в 4-кортеж."""
    ldr = str(k_list[3]).strip() if len(k_list) > 3 and k_list[3] is not None else ""
    load_value = k_list[2]
    if isinstance(load_value, (int, float)):
        normalized_load: Union[int, float] = float(load_value)
    else:
        normalized_load = int(load_value)
    return (float(k_list[0]), float(k_list[1]), normalized_load, ldr)


def _try_fill_plate_nomenclature_cache(
    fill_nomenclature_cache: NomenclatureCacheFiller | None,
) -> None:
    filler = (
        fill_nomenclature_cache
        if fill_nomenclature_cache is not None
        else get_default_nomenclature_cache_filler()
    )
    if not callable(filler):
        return
    try:
        filler()
    except Exception as _e:
        logger.warning(f"Не удалось заполнить PLATE_NOMENCLATURE_CACHE: {_e}")


@dataclass
class PlateOrder:
    """
    Данные заказа плит: списки по ширине, карта нагрузок, точные ширины, итоги.
    Используется для изоляции заказа по пользователю (хранение в FSM state, передача в оптимизацию/визуализацию).
    """

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
    plate_load_details: Dict[Tuple[float, float, Union[int, float], str], int] = field(default_factory=dict)
    plate_length_dm_raw: Dict[Tuple[float, float, Union[int, float], str], str] = field(default_factory=dict)
    plate_exact_widths: Dict[Tuple[float, str], float] = field(default_factory=dict)
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

    def to_dict(self) -> dict:
        """Сериализация для FSM state (JSON-совместимые ключи)."""
        return {
            "plates_1_2": list(self.plates_1_2),
            "plates_1_5_to_1_2": list(self.plates_1_5_to_1_2),
            "plates_1_0": list(self.plates_1_0),
            "plates_1_08": list(self.plates_1_08),
            "plates_0_46": list(self.plates_0_46),
            "plates_0_32": list(self.plates_0_32),
            "plates_0_72": list(self.plates_0_72),
            "plates_0_70": list(self.plates_0_70),
            "plates_0_86": list(self.plates_0_86),
            "plates_0_74": list(self.plates_0_74),
            "plates_0_88": list(self.plates_0_88),
            "plates_0_48": list(self.plates_0_48),
            "plates_0_50": list(self.plates_0_50),
            "plates_0_34": list(self.plates_0_34),
            "plate_load_details": [[list(k), v] for k, v in self.plate_load_details.items()],
            "plate_length_dm_raw": [[list(k), v] for k, v in self.plate_length_dm_raw.items()],
            "plate_exact_widths": [[list(k), v] for k, v in self.plate_exact_widths.items()],
            "longitudinal_cuts": self.longitudinal_cuts,
            "length_trims": self.length_trims,
            "unused_strips_0_3_m_total": self.unused_strips_0_3_m_total,
            "scrap_strips_0_2_m_total": self.scrap_strips_0_2_m_total,
            "usable_strips_0_74_m_total": self.usable_strips_0_74_m_total,
            "usable_strips_0_88_m_total": self.usable_strips_0_88_m_total,
            "usable_strips_0_48_m_total": self.usable_strips_0_48_m_total,
            "usable_strips_0_50_m_total": self.usable_strips_0_50_m_total,
            "usable_strips_0_34_m_total": self.usable_strips_0_34_m_total,
            "scrap_strips_0_12_m_total": self.scrap_strips_0_12_m_total,
            "waste_area_m2": self.waste_area_m2,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "PlateOrder":
        """Восстановление из FSM state."""
        load_details = {}
        for k_list, v in d.get("plate_load_details", []):
            load_details[parse_load_key_from_list(k_list)] = int(v)
        length_dm_raw = {}
        for k_list, v in d.get("plate_length_dm_raw", []):
            key = parse_load_key_from_list(k_list)
            length_dm_raw[key] = str(v) if v is not None else ""
        exact_widths = {}
        for k_list, v in d.get("plate_exact_widths", []):
            exact_widths[(float(k_list[0]), str(k_list[1]))] = float(v)
        return cls(
            plates_1_2=list(d.get("plates_1_2", [])),
            plates_1_5_to_1_2=list(d.get("plates_1_5_to_1_2", [])),
            plates_1_0=list(d.get("plates_1_0", [])),
            plates_1_08=list(d.get("plates_1_08", [])),
            plates_0_46=list(d.get("plates_0_46", [])),
            plates_0_32=list(d.get("plates_0_32", [])),
            plates_0_72=list(d.get("plates_0_72", [])),
            plates_0_70=list(d.get("plates_0_70", [])),
            plates_0_86=list(d.get("plates_0_86", [])),
            plates_0_74=list(d.get("plates_0_74", [])),
            plates_0_88=list(d.get("plates_0_88", [])),
            plates_0_48=list(d.get("plates_0_48", [])),
            plates_0_50=list(d.get("plates_0_50", [])),
            plates_0_34=list(d.get("plates_0_34", [])),
            plate_load_details=load_details,
            plate_length_dm_raw=length_dm_raw,
            plate_exact_widths=exact_widths,
            longitudinal_cuts=int(d.get("longitudinal_cuts", 0)),
            length_trims=int(d.get("length_trims", 0)),
            unused_strips_0_3_m_total=float(d.get("unused_strips_0_3_m_total", 0)),
            scrap_strips_0_2_m_total=float(d.get("scrap_strips_0_2_m_total", 0)),
            usable_strips_0_74_m_total=float(d.get("usable_strips_0_74_m_total", 0)),
            usable_strips_0_88_m_total=float(d.get("usable_strips_0_88_m_total", 0)),
            usable_strips_0_48_m_total=float(d.get("usable_strips_0_48_m_total", 0)),
            usable_strips_0_50_m_total=float(d.get("usable_strips_0_50_m_total", 0)),
            usable_strips_0_34_m_total=float(d.get("usable_strips_0_34_m_total", 0)),
            scrap_strips_0_12_m_total=float(d.get("scrap_strips_0_12_m_total", 0)),
            waste_area_m2=float(d.get("waste_area_m2", 0)),
        )

    @classmethod
    def from_runtime(cls, rt: PlateMutableRuntime) -> "PlateOrder":
        """Снимок заказа из мутабельного рантайма (``PlateOrderContext.plates``)."""
        return cls(
            plates_1_2=list(rt.plates_1_2),
            plates_1_5_to_1_2=list(rt.plates_1_5_to_1_2),
            plates_1_0=list(rt.plates_1_0),
            plates_1_08=list(rt.plates_1_08),
            plates_0_46=list(rt.plates_0_46),
            plates_0_32=list(rt.plates_0_32),
            plates_0_72=list(rt.plates_0_72),
            plates_0_70=list(rt.plates_0_70),
            plates_0_86=list(rt.plates_0_86),
            plates_0_74=list(rt.plates_0_74),
            plates_0_88=list(rt.plates_0_88),
            plates_0_48=list(rt.plates_0_48),
            plates_0_50=list(rt.plates_0_50),
            plates_0_34=list(rt.plates_0_34),
            plate_load_details=dict(rt.plate_load_details),
            plate_length_dm_raw=dict(rt.plate_length_dm_raw),
            plate_exact_widths=dict(rt.plate_exact_widths),
            longitudinal_cuts=int(rt.longitudinal_cuts),
            length_trims=int(rt.length_trims),
            unused_strips_0_3_m_total=float(rt.unused_strips_0_3_m_total),
            scrap_strips_0_2_m_total=float(rt.scrap_strips_0_2_m_total),
            usable_strips_0_74_m_total=float(rt.usable_strips_0_74_m_total),
            usable_strips_0_88_m_total=float(rt.usable_strips_0_88_m_total),
            usable_strips_0_48_m_total=float(rt.usable_strips_0_48_m_total),
            usable_strips_0_50_m_total=float(rt.usable_strips_0_50_m_total),
            usable_strips_0_34_m_total=float(rt.usable_strips_0_34_m_total),
            scrap_strips_0_12_m_total=float(rt.scrap_strips_0_12_m_total),
            waste_area_m2=float(rt.waste_area_m2),
        )

    @classmethod
    def from_orders_2d(cls, orders_2d: List[Dict]) -> "PlateOrder":
        """Строит заказ из списка dict с ключами length, width (мм), qty, load_code (как в state/плане)."""
        order = cls()
        width_to_list = {
            1200: order.plates_1_2,
            1080: order.plates_1_08,
            1000: order.plates_1_0,
            320: order.plates_0_32,
            460: order.plates_0_46,
            700: order.plates_0_70,
            720: order.plates_0_72,
            860: order.plates_0_86,
            880: order.plates_0_88,
            740: order.plates_0_74,
            480: order.plates_0_48,
            500: order.plates_0_50,
            340: order.plates_0_34,
        }
        for p in orders_2d:
            length = float(p["length"])
            width_mm = int(p["width"])
            width_m = width_mm / 1000.0
            qty = int(p.get("qty", 1))
            load_code = normalize_load_code(p.get("load_code", 8), default=8)
            ldr = (p.get("length_dm_raw") or "").strip()
            key = (round(length, 3), round(width_m, 3), load_code, ldr)
            order.plate_load_details[key] = order.plate_load_details.get(key, 0) + qty
            order.plate_length_dm_raw[key] = ldr
            w = width_mm
            if w in width_to_list:
                lst = width_to_list[w]
            elif 1020 <= w <= 1080:
                lst = order.plates_1_08
            elif 260 <= w <= 320:
                lst = order.plates_0_32
            elif 460 <= w <= 530:
                lst = order.plates_0_46
            elif 660 <= w <= 720:
                lst = order.plates_0_72 if w >= 710 else order.plates_0_70
            elif 860 <= w <= 920:
                lst = order.plates_0_86
            else:
                lst = order.plates_1_2 if abs(width_m - 1.2) < 0.01 else None
            if lst is not None:
                for _ in range(qty):
                    lst.append(length)
        order._recompute_totals()
        return order

    def recompute_totals(self) -> None:
        """Пересчитывает итоговые поля из списков плит."""
        self._recompute_totals()

    def _recompute_totals(self) -> None:
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

    def to_orders_2d(self) -> List[Dict]:
        """Список dict {length, width, qty, load_code, length_dm_raw} для оптимизатора и state."""
        out = []
        for key, qty in self.plate_load_details.items():
            length, width_m, load_code = key[0], key[1], key[2]
            ldr = key[3] if len(key) > 3 else self.plate_length_dm_raw.get(key, "")
            out.append(
                {
                    "length": length,
                    "width": int(round(width_m * 1000)),
                    "qty": qty,
                    "load_code": load_code,
                    "length_dm_raw": ldr,
                }
            )
        return out

    def apply_to_globals(
        self,
        *,
        fill_nomenclature_cache: NomenclatureCacheFiller | None = None,
    ) -> None:
        """Записывает данные заказа в потоколокальное / контекстное состояние cfg.

        .. deprecated:: A1-002
            Prefer ``PlateOrderContext.hydrate_from_order()`` — см. ``core.plate_order_context``.
        """
        warnings.warn(_APPLY_TO_GLOBALS_DEPRECATION, DeprecationWarning, stacklevel=2)
        rt = get_plate_mutable_runtime()
        rt.plate_load_details.clear()
        rt.plate_load_details.update(self.plate_load_details)
        rt.plate_length_dm_raw.clear()
        rt.plate_length_dm_raw.update(self.plate_length_dm_raw)
        rt.plate_nomenclature_cache.clear()
        _try_fill_plate_nomenclature_cache(fill_nomenclature_cache)
        rt.plates_1_2 = list(self.plates_1_2)
        rt.plates_1_5_to_1_2 = list(self.plates_1_5_to_1_2)
        rt.plates_1_0 = list(self.plates_1_0)
        rt.plates_1_08 = list(self.plates_1_08)
        rt.plates_0_46 = list(self.plates_0_46)
        rt.plates_0_32 = list(self.plates_0_32)
        rt.plates_0_72 = list(self.plates_0_72)
        rt.plates_0_70 = list(self.plates_0_70)
        rt.plates_0_86 = list(self.plates_0_86)
        rt.plates_0_74 = list(self.plates_0_74)
        rt.plates_0_88 = list(self.plates_0_88)
        rt.plates_0_48 = list(self.plates_0_48)
        rt.plates_0_50 = list(self.plates_0_50)
        rt.plates_0_34 = list(self.plates_0_34)
        rt.plate_exact_widths.clear()
        rt.plate_exact_widths.update(self.plate_exact_widths)
        rt.longitudinal_cuts = self.longitudinal_cuts
        rt.length_trims = self.length_trims
        rt.unused_strips_0_3_m_total = self.unused_strips_0_3_m_total
        rt.scrap_strips_0_2_m_total = self.scrap_strips_0_2_m_total
        rt.usable_strips_0_74_m_total = self.usable_strips_0_74_m_total
        rt.usable_strips_0_88_m_total = self.usable_strips_0_88_m_total
        rt.usable_strips_0_48_m_total = self.usable_strips_0_48_m_total
        rt.usable_strips_0_50_m_total = self.usable_strips_0_50_m_total
        rt.usable_strips_0_34_m_total = self.usable_strips_0_34_m_total
        rt.scrap_strips_0_12_m_total = self.scrap_strips_0_12_m_total
        rt.waste_area_m2 = self.waste_area_m2


def coerce_core_plate_order(plate_order: Any) -> PlateOrder:
    """Привести app/legacy заказ к каноническому ``PlateOrder`` (без app-only полей)."""
    if type(plate_order) is PlateOrder:
        return plate_order
    if hasattr(plate_order, "to_dict"):
        return PlateOrder.from_dict(dict(plate_order.to_dict()))
    raise TypeError(f"Unsupported plate order type: {type(plate_order)!r}")


def get_current_plate_order() -> PlateOrder:
    """Строит PlateOrder из текущего потоколокального / контекстного состояния.

    .. deprecated:: A1-002
        Prefer ``PlateOrderContext.plates`` or ``hydrate_from_order()`` for explicit context.
    """
    warnings.warn(_GET_CURRENT_PLATE_ORDER_DEPRECATION, DeprecationWarning, stacklevel=2)
    return PlateOrder.from_runtime(get_plate_mutable_runtime())
