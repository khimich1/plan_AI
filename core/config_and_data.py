#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конфигурация и данные проекта:
- Константы (размеры дорожки, цены резов)
- Глобальные списки плит
- Парсинг текста пользователя
"""
import math
import os
import re
import logging
from pathlib import Path
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union

# Импортируем пользовательские исключения
from .exceptions import PlateParseError
from .plate_line_parser import parse_line
from .plate_validation import validate_plate_values

# Настройка логирования
logger = logging.getLogger(__name__)

# ==================== КОНСТАНТЫ ====================

TRACK_LENGTH_M = 101.0
TRACK_WIDTH_M = 1.2

# Пути к прайсам (BASE_DIR указывает на корень проекта, на уровень выше core/)
BASE_DIR = Path(__file__).resolve().parent.parent
PRICE_XLSX_PATH = BASE_DIR / 'банк знаний' / 'Новые цены для прайса с 19.08.24.xlsx'
CUTS_DOCX_PATH = BASE_DIR / 'банк знаний' / 'Письмо Цены с 29.05.2024 цены на резы.docx'
PRICE_DB_PATH = BASE_DIR / 'pb.db'

# Стоимость резов
LONG_CUT_PRICE_PER_M = 460.0  # Продольный рез, руб/пог.м
TRANSVERSE_CUT_PRICE = 1200.0  # Поперечный (или скошенный) рез, руб/шт

# Источник веса для КП:
# - "formula": расчет по формуле (дм * дм * коэффициент)
# - "plate_weights": legacy-режим через таблицу plate_weights
WEIGHT_SOURCE = os.getenv("WEIGHT_SOURCE", "formula").strip().lower()

# Вес 1 дм длины при ширине 1 дм, кг.
WEIGHT_KG_PER_DM2 = 2.83333333

# Длина из марки ПБ: целое без запятой/точки = номинал в дм → длина = номинал / 10 м (без вычета 20 мм)
def length_dm_to_m(Ldm_str: str) -> float:
    """
    Переводит строку длины из марки ПБ в метры.
    - Нет запятой/точки (например "38") → номинал в дм: длина = номинал / 10 м (38 → 3.8 м).
    - Есть запятая/точка ("38,0", "75,5") → точное значение: парсим в целые мм, затем одно деление на 1000.
    """
    s = (Ldm_str or "").strip().replace(" ", "")
    if "," in s or "." in s:
        # Ветка с запятой/точкой: разбираем строку без float, считаем длину в целых мм
        parts = s.replace(",", ".").split(".")
        if len(parts) != 2:
            try:
                return round(float(s.replace(",", ".")) / 10.0, 3)
            except ValueError:
                return 0.0
        try:
            int_part = int(parts[0].strip())
            frac_str = (parts[1].strip() or "0")[:3]  # не более 3 знаков
            frac_part = int(frac_str) if frac_str else 0
            denom = 10 ** len(frac_str)  # 1, 10, 100
            # Длина в мм: value_dm = int_part + frac_part/denom, length_mm = value_dm * 100
            # Целочисленно: (int_part * denom + frac_part) * 100 // denom
            length_mm = (int_part * denom + frac_part) * 100 // denom
        except ValueError:
            return 0.0
        if length_mm < 0:
            return 0.0
        return round(length_mm / 1000.0, 3)
    try:
        nominal_dm = int(float(s))
    except ValueError:
        return 0.0
    # Номинал в дм → длина в метрах без вычета 20 мм (69 → 6.9 м)
    return round(nominal_dm / 10.0, 3)


def normalize_dimension(value_str: str) -> float:
    """
    Нормализует размер из строки в дециметры.

    Примеры:
    - "5,30" -> 5.3
    - "530" -> 5.3 (введены мм без разделителя)
    - "6.65" -> 6.65
    """
    clean = value_str.strip().replace(" ", "").replace(",", ".")
    try:
        val = float(clean)
    except ValueError:
        return 0.0

    # Диапазон значений, которые чаще всего вводят как мм без запятой:
    # 530 -> 5.30 дм, 665 -> 6.65 дм.
    if 20.0 < val < 1000.0:
        val = val / 100.0
        logger.debug(
            "normalize_dimension: применили /100 для значения, похожего на мм: %s -> %s",
            value_str.strip(),
            val,
        )
    return val


def parse_pb_width_to_m(width_str: str) -> float:
    """
    Парсит ширину из марки ПБ/ПК (часть W в L-W-N) в метры.

    Правила:
    - 0.2/0.3 (или 0,2/0,3) считаются уже метрами (спец-ленты).
    - Остальные значения интерпретируются как дециметры и делятся на 10.
    """
    width_raw = normalize_dimension(width_str)
    if width_raw <= 0:
        return 0.0
    if abs(width_raw - 0.3) < 1e-6 or abs(width_raw - 0.2) < 1e-6:
        return round(width_raw, 3)
    return round(width_raw / 10.0, 3)

# ==================== ГЛОБАЛЬНЫЕ СПИСКИ ПЛИТ ====================

# Данные из согласованного КЗ-плана
# 1) Плиты 1.2 м — без резов (новый заказ)
PLATES_1_2 = [3.39] * 2

# Дополнительные целевые ширины, которые получаем продольным резом из 1.2 м
# По умолчанию считаем, что доступен такой же набор плит шириной 1.08 м.
PLATES_1_08 = PLATES_1_2.copy()
PLATES_0_46 = []         # нет 0.46 в этом заказе
PLATES_0_32 = [6.63]*4 + [7.83]*3
PLATES_0_72 = [5.63]*5
PLATES_0_70 = [4.65]*5
PLATES_0_86 = [6.75]*2 + [4.65]*5

# Заказы на вторую половину (если пользователь прислал такие ширины)
PLATES_0_74 = []
PLATES_0_88 = []
PLATES_0_48 = []
PLATES_0_50 = []
PLATES_0_34 = []

# 2) Плиты 1.5 м — используем как 1.2 м (лента 0.3 образуется)
PLATES_1_5_TO_1_2 = []
# 3) Плиты 1.0 м — получаем из 1.2 (остаток 0.2 уходит в обрезки)
PLATES_1_0 = []

# Резы по плану: по одному на каждую плиту, получаемую резом
LONGITUDINAL_CUTS = (
    len(PLATES_1_5_TO_1_2) + len(PLATES_1_0) +
    len(PLATES_1_08) + len(PLATES_0_46) +
    len(PLATES_0_32) + len(PLATES_0_72) + len(PLATES_0_70) + len(PLATES_0_86)
)
LENGTH_TRIMS = 0

# Остатки и отходы
UNUSED_STRIPS_0_3_M_TOTAL = 0.0
SCRAP_STRIPS_0_2_M_TOTAL = 0.0
USABLE_STRIPS_0_74_M_TOTAL = round(sum(PLATES_0_46), 1)
USABLE_STRIPS_0_88_M_TOTAL = round(sum(PLATES_0_32), 1)
USABLE_STRIPS_0_48_M_TOTAL = round(sum(PLATES_0_72), 1)
USABLE_STRIPS_0_50_M_TOTAL = round(sum(PLATES_0_70), 1)
USABLE_STRIPS_0_34_M_TOTAL = round(sum(PLATES_0_86), 1)
SCRAP_STRIPS_0_12_M_TOTAL = round(sum(PLATES_1_08), 1)
WASTE_AREA_M2 = round(0.12 * SCRAP_STRIPS_0_12_M_TOTAL, 2)

# Метаданные плит для визуализации и смет
PLATE_METADATA: Dict[Tuple[float, int], List[Dict[str, Any]]] = {}

# Детальная карта плит с нагрузкой (единственный источник правды по нагрузкам)
# Формат: (длина, ширина, нагрузка, номинал_длины) → количество
# Пример: {(7.3, 1.2, 8, ""): 93, (4.0, 1.2, 8, "40"): 4, (4.0, 1.2, 8, "40,0"): 6}
# Четвёртый элемент — нормализованная строка длины из марки: "40" и "40,0" — разные позиции.
PLATE_LOAD_DETAILS: Dict[Tuple[float, float, int, str], int] = {}

# Карта точных размеров плит: (длина, имя_списка) → точная_ширина_в_метрах
# Сохраняет ТОЧНУЮ ширину каждой плиты при парсинге заказа.
# Пример: {(2.8, 'PLATES_0_46'): 0.53} - плита ПБ 28-5,3-8п имеет ширину 530мм, а не 460мм
PLATE_EXACT_WIDTHS: Dict[Tuple[float, str], float] = {}

# Исходная строка длины из марки для каждой позиции заказа (ключ = как в PLATE_LOAD_DETAILS).
# Нужна для отображения и поиска при списании: "59,81" vs "59,84" не сливаются.
PLATE_LENGTH_DM_RAW: Dict[Tuple[float, float, int, str], str] = {}

# Кэш номенклатуры: canonical_name и nomenclature_id из prays_plity, найденные по имени
# с length_dm_raw (номинальная длина). Ключ = как в PLATE_LOAD_DETAILS.
# Заполняется один раз после парсинга или восстановления заказа.
# Значение: {"canonical_name": str | None, "nomenclature_id": str | None}
PLATE_NOMENCLATURE_CACHE: Dict[Tuple[float, float, int, str], Dict[str, Any]] = {}

# Карта максимального армирования дорожки для каждой плиты
# Формат: (длина, ширина_мм) → максимальное армирование в дорожке, где лежит эта плита
# Заполняется в visualization.py после формирования дорожек
PLATE_MAX_REINFORCEMENT_MAP: Dict[Tuple[float, int], float] = {}

# Последняя диагностика распознавания (по строкам) для отладки и UI.
LAST_PARSE_DIAGNOSTICS: list[dict[str, Any]] = []


# ==================== КЛАСС ЗАКАЗА (PlateOrder) ====================

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
        def _parse_load_key(k_list):
            """Конвертирует список [length, width, load_code] или [length, width, load_code, ldr] в 4-кортеж."""
            ldr = str(k_list[3]).strip() if len(k_list) > 3 and k_list[3] is not None else ""
            return (float(k_list[0]), float(k_list[1]), float(k_list[2]) if isinstance(k_list[2], (int, float)) else int(k_list[2]), ldr)

        load_details = {}
        for k_list, v in d.get("plate_load_details", []):
            load_details[_parse_load_key(k_list)] = int(v)
        length_dm_raw = {}
        for k_list, v in d.get("plate_length_dm_raw", []):
            key = _parse_load_key(k_list)
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
    def from_orders_2d(cls, orders_2d: List[Dict]) -> "PlateOrder":
        """Строит заказ из списка dict с ключами length, width (мм), qty, load_code (как в state/плане)."""
        order = cls()
        # Распределение по ширине (мм) -> список для добавления длин
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
            # Ближайшая стандартная ширина для раскладки по спискам
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

    def _recompute_totals(self) -> None:
        """Пересчитывает итоговые поля из списков плит."""
        self.longitudinal_cuts = (
            len(self.plates_1_5_to_1_2) + len(self.plates_1_0) +
            len(self.plates_1_08) + len(self.plates_0_46) +
            len(self.plates_0_32) + len(self.plates_0_72) + len(self.plates_0_70) + len(self.plates_0_86)
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
            out.append({
                "length": length,
                "width": int(round(width_m * 1000)),
                "qty": qty,
                "load_code": load_code,
                "length_dm_raw": ldr,
            })
        return out

    def apply_to_globals(self) -> None:
        """Записывает данные заказа в глобальные переменные cfg (для обратной совместимости с оптимизацией/визуализацией)."""
        g = globals()
        g["PLATE_LOAD_DETAILS"].clear()
        g["PLATE_LOAD_DETAILS"].update(self.plate_load_details)
        g["PLATE_LENGTH_DM_RAW"].clear()
        g["PLATE_LENGTH_DM_RAW"].update(self.plate_length_dm_raw)
        g["PLATE_NOMENCLATURE_CACHE"].clear()
        try:
            from core.kp_db import fill_plate_nomenclature_cache
            fill_plate_nomenclature_cache()
        except Exception as _e:
            import logging as _logging
            _logging.getLogger(__name__).warning(f"Не удалось заполнить PLATE_NOMENCLATURE_CACHE: {_e}")
        g["PLATES_1_2"] = list(self.plates_1_2)
        g["PLATES_1_5_TO_1_2"] = list(self.plates_1_5_to_1_2)
        g["PLATES_1_0"] = list(self.plates_1_0)
        g["PLATES_1_08"] = list(self.plates_1_08)
        g["PLATES_0_46"] = list(self.plates_0_46)
        g["PLATES_0_32"] = list(self.plates_0_32)
        g["PLATES_0_72"] = list(self.plates_0_72)
        g["PLATES_0_70"] = list(self.plates_0_70)
        g["PLATES_0_86"] = list(self.plates_0_86)
        g["PLATES_0_74"] = list(self.plates_0_74)
        g["PLATES_0_88"] = list(self.plates_0_88)
        g["PLATES_0_48"] = list(self.plates_0_48)
        g["PLATES_0_50"] = list(self.plates_0_50)
        g["PLATES_0_34"] = list(self.plates_0_34)
        g["PLATE_EXACT_WIDTHS"].clear()
        g["PLATE_EXACT_WIDTHS"].update(self.plate_exact_widths)
        g["LONGITUDINAL_CUTS"] = self.longitudinal_cuts
        g["LENGTH_TRIMS"] = self.length_trims
        g["UNUSED_STRIPS_0_3_M_TOTAL"] = self.unused_strips_0_3_m_total
        g["SCRAP_STRIPS_0_2_M_TOTAL"] = self.scrap_strips_0_2_m_total
        g["USABLE_STRIPS_0_74_M_TOTAL"] = self.usable_strips_0_74_m_total
        g["USABLE_STRIPS_0_88_M_TOTAL"] = self.usable_strips_0_88_m_total
        g["USABLE_STRIPS_0_48_M_TOTAL"] = self.usable_strips_0_48_m_total
        g["USABLE_STRIPS_0_50_M_TOTAL"] = self.usable_strips_0_50_m_total
        g["USABLE_STRIPS_0_34_M_TOTAL"] = self.usable_strips_0_34_m_total
        g["SCRAP_STRIPS_0_12_M_TOTAL"] = self.scrap_strips_0_12_m_total
        g["WASTE_AREA_M2"] = self.waste_area_m2


def get_current_plate_order() -> PlateOrder:
    """Строит PlateOrder из текущих глобальных переменных (после set_plate_lists_from_text)."""
    return PlateOrder(
        plates_1_2=list(PLATES_1_2),
        plates_1_5_to_1_2=list(PLATES_1_5_TO_1_2),
        plates_1_0=list(PLATES_1_0),
        plates_1_08=list(PLATES_1_08),
        plates_0_46=list(PLATES_0_46),
        plates_0_32=list(PLATES_0_32),
        plates_0_72=list(PLATES_0_72),
        plates_0_70=list(PLATES_0_70),
        plates_0_86=list(PLATES_0_86),
        plates_0_74=list(PLATES_0_74),
        plates_0_88=list(PLATES_0_88),
        plates_0_48=list(PLATES_0_48),
        plates_0_50=list(PLATES_0_50),
        plates_0_34=list(PLATES_0_34),
        plate_load_details=dict(PLATE_LOAD_DETAILS),
        plate_length_dm_raw=dict(PLATE_LENGTH_DM_RAW),
        plate_exact_widths=dict(PLATE_EXACT_WIDTHS),
        longitudinal_cuts=int(LONGITUDINAL_CUTS),
        length_trims=int(LENGTH_TRIMS),
        unused_strips_0_3_m_total=float(UNUSED_STRIPS_0_3_M_TOTAL),
        scrap_strips_0_2_m_total=float(SCRAP_STRIPS_0_2_M_TOTAL),
        usable_strips_0_74_m_total=float(USABLE_STRIPS_0_74_M_TOTAL),
        usable_strips_0_88_m_total=float(USABLE_STRIPS_0_88_M_TOTAL),
        usable_strips_0_48_m_total=float(USABLE_STRIPS_0_48_M_TOTAL),
        usable_strips_0_50_m_total=float(USABLE_STRIPS_0_50_M_TOTAL),
        usable_strips_0_34_m_total=float(USABLE_STRIPS_0_34_M_TOTAL),
        scrap_strips_0_12_m_total=float(SCRAP_STRIPS_0_12_M_TOTAL),
        waste_area_m2=float(WASTE_AREA_M2),
    )


def get_last_parse_diagnostics() -> list[dict[str, Any]]:
    """Возвращает диагностику последнего запуска set_plate_lists_from_text()."""
    return list(LAST_PARSE_DIAGNOSTICS)


# ==================== ФУНКЦИИ ПАРСИНГА ====================

def _clear_all_plate_lists():
    """Очищает все глобальные списки плит"""
    global PLATES_1_2, PLATES_1_5_TO_1_2, PLATES_1_0, PLATES_1_08
    global PLATES_0_46, PLATES_0_32, PLATES_0_72, PLATES_0_70, PLATES_0_86
    global PLATES_0_74, PLATES_0_88, PLATES_0_48, PLATES_0_50, PLATES_0_34
    global PLATE_LOAD_DETAILS, PLATE_EXACT_WIDTHS, PLATE_LENGTH_DM_RAW, PLATE_MAX_REINFORCEMENT_MAP
    global PLATE_NOMENCLATURE_CACHE
    PLATES_1_2 = []
    PLATES_1_5_TO_1_2 = []
    PLATES_1_0 = []
    PLATES_1_08 = []
    PLATES_0_46 = []
    PLATES_0_32 = []
    PLATES_0_72 = []
    PLATES_0_70 = []
    PLATES_0_86 = []
    PLATES_0_74 = []
    PLATES_0_88 = []
    PLATES_0_48 = []
    PLATES_0_50 = []
    PLATES_0_34 = []
    PLATE_LOAD_DETAILS.clear()
    PLATE_EXACT_WIDTHS.clear()
    PLATE_LENGTH_DM_RAW.clear()
    PLATE_MAX_REINFORCEMENT_MAP.clear()
    PLATE_NOMENCLATURE_CACHE.clear()


def _recompute_totals_from_lists():
    """Пересчитывает глобальные итоговые переменные на основе списков плит"""
    global LONGITUDINAL_CUTS, LENGTH_TRIMS
    global UNUSED_STRIPS_0_3_M_TOTAL, SCRAP_STRIPS_0_2_M_TOTAL
    global USABLE_STRIPS_0_74_M_TOTAL, USABLE_STRIPS_0_88_M_TOTAL
    global USABLE_STRIPS_0_48_M_TOTAL, USABLE_STRIPS_0_50_M_TOTAL
    global USABLE_STRIPS_0_34_M_TOTAL, SCRAP_STRIPS_0_12_M_TOTAL
    global WASTE_AREA_M2

    LONGITUDINAL_CUTS = (
        len(PLATES_1_5_TO_1_2) + len(PLATES_1_0) +
        len(PLATES_1_08) + len(PLATES_0_46) +
        len(PLATES_0_32) + len(PLATES_0_72) + len(PLATES_0_70) + len(PLATES_0_86)
    )
    LENGTH_TRIMS = 0

    UNUSED_STRIPS_0_3_M_TOTAL = 0.0
    SCRAP_STRIPS_0_2_M_TOTAL = 0.0
    USABLE_STRIPS_0_74_M_TOTAL = round(sum(PLATES_0_46), 1)
    USABLE_STRIPS_0_88_M_TOTAL = round(sum(PLATES_0_32), 1)
    USABLE_STRIPS_0_48_M_TOTAL = round(sum(PLATES_0_72), 1)
    USABLE_STRIPS_0_50_M_TOTAL = round(sum(PLATES_0_70), 1)
    USABLE_STRIPS_0_34_M_TOTAL = round(sum(PLATES_0_86), 1)
    SCRAP_STRIPS_0_12_M_TOTAL = round(sum(PLATES_1_08), 1)
    WASTE_AREA_M2 = round(0.12 * SCRAP_STRIPS_0_12_M_TOTAL, 2)


# Ключ для сопоставления строки ввода с позициями заказа / order_data:
# длина и ширина в метрах, код нагрузки (или None для размерных строк без марки), сырой фрагмент длины из марки.
LineContributionKey = Tuple[float, float, Optional[float], str]


def set_plate_lists_from_text(
    user_text: str,
) -> tuple[list[str], list[list[LineContributionKey]], list[dict[tuple, int]]]:
    """Парсит свободный текст пользователя и заполняет списки PLATES_*.

    Поддерживаемые форматы (регистр не важен, пробелы опциональны):
      - Размеры через «x» или «×»: «1.2×3.39 — 2 шт», «0,32x6,63 - 4»
      - Марка ПБ/ПК: «ПБ 78-12-8п 3», «Плиты ПБ 78-12-8п», «ПБ78-12-8п 5», «ПК 80-12-8 7»
        Длина и ширина в дециметрах (78 => 7.8 м, 12 => 1.2 м), нагрузка — после последнего дефиса (8п, 10п и т.д.)
      - Количество: после марки, опционально «шт» («8п 5», «8п 5 шт», «8п — 5»)
    Неизвестные ширины и нераспознанные строки возвращаются в списке нераспознанных.

    Returns:
        (unparsed_lines, line_contributions, line_plate_load_details): нераспознанные строки;
        для каждой строки ``lines`` — список ключей вклада; для каждой строки — словарь
        накопленных количеств по тем же ключам, что ``PLATE_LOAD_DETAILS``, только для этой строки ввода.

    Raises:
        PlateParseError: Если текст пустой или после разбивки не осталось валидных строк.
    """
    if not user_text or not user_text.strip():
        logger.warning("Получен пустой текст заказа")
        raise PlateParseError(
            "Текст заказа пустой. Пожалуйста, введите список плит.\n"
            "Пример: ПБ 78-12-8п 5 шт"
        )
    
    global LAST_PARSE_DIAGNOSTICS
    _clear_all_plate_lists()
    LAST_PARSE_DIAGNOSTICS = []

    # Нормализация: конвертация каталожных марок (ПБ 59.12-8Вр1400-25 → ПБ 59-12-8п)
    # и других нестандартных вариантов записи перед основным парсингом.
    _processing_text = user_text
    try:
        from .plate_text_normalizer import normalize_order_text
        _norm = normalize_order_text(user_text)
        if _norm.warnings:
            for _w in _norm.warnings[:10]:
                logger.info("Нормализатор: %s", _w)
        if _norm.normalized_text.strip():
            _processing_text = _norm.normalized_text
    except Exception as _norm_err:
        logger.warning("Ошибка нормализатора, используем исходный текст: %s", _norm_err)

    # Нормализация: единый символ умножения, неразрывные пробелы как обычные
    text = (_processing_text or '').replace('\u00d7', 'x').replace('×', 'x')
    text = text.replace('\u00a0', ' ')
    lines = [re.sub(r'\s+', ' ', l).strip() for l in re.split(r'[\n;]+', text) if l.strip()]
    line_contributions: list[list[LineContributionKey]] = [[] for _ in lines]
    line_plate_load_details: list[dict[tuple, int]] = [{} for _ in lines]

    def _record_contribution(line_idx: int, length_m: float, width_m: float, load_code: Optional[float], ldr: str) -> None:
        if line_idx < 0 or line_idx >= len(line_contributions):
            return
        line_contributions[line_idx].append(
            (round(float(length_m), 3), round(float(width_m), 3), load_code, (ldr or "").strip())
        )

    # Дополнительная проверка после разбивки на строки
    if not lines:
        logger.warning("После разбивки не осталось валидных строк")
        raise PlateParseError(
            "Не удалось найти ни одной строки с плитами.\n"
            "Проверьте формат ввода."
        )

    def add_items(
        width_m: float,
        length_m: float,
        qty: int,
        load_code: int = None,
        length_dm_raw: str = None,
        line_idx: Optional[int] = None,
    ):
        """
        Добавляет плиты в соответствующий глобальный список по ширине.

        Если ширина не попадает ни в один из «жёстких» диапазонов,
        используем правило: берём ближайший МЕНЬШИЙ допустимый рез
        (включая остаточные ширины) и повторно вызываем add_items
        уже с притянутой шириной.
        
        Args:
            width_m: ширина плиты в метрах
            length_m: длина плиты в метрах
            qty: количество плит
            load_code: нагрузка (6, 8, 10, 12, 13 и т.д.) - опционально
            length_dm_raw: исходная строка длины из марки (например "59,81") для различения плит
            line_idx: индекс строки заказа (для line_contributions); None — не писать вклад
        """
        # ЗАЩИТА: Проверяем адекватность размеров
        # Если размеры слишком большие (вероятно, ошибка OCR распознал мм как дм)
        # то игнорируем эту плиту
        if width_m > 20.0 or length_m > 200.0:
            logger.warning(
                f"Пропущена плита с неадекватными размерами: {length_m}м × {width_m}м. "
                f"Возможно, OCR распознал мм как дм."
            )
            return
        
        if width_m <= 0 or length_m <= 0:
            logger.warning(f"Пропущена плита с нулевыми размерами: {length_m}м × {width_m}м")
            return

        # Защита от зависаний: слишком большое количество может «повесить» бот,
        # потому что ниже мы добавляем в списки по 1 штуке в цикле.
        if qty is None or qty <= 0:
            logger.warning(f"Пропущена плита с некорректным количеством: qty={qty}")
            return
        if qty > 500:
            logger.warning(f"Пропущена строка с слишком большим количеством плит: qty={qty}")
            return
        # Специальная обработка плит 1.5 м → заменяем на 1.2 м + 0.3 м
        if 1.45 <= width_m <= 1.55:  # 1.5 м (диапазон ±50 мм)
            length_rounded = round(float(length_m), 3)
            # Добавляем плиту 1.2 м
            for _ in range(max(0, qty)):
                PLATES_1_2.append(length_rounded)
                # Сохраняем точную ширину 1.2м
                PLATE_EXACT_WIDTHS[(length_rounded, 'PLATES_1_2')] = 1.2
            # Добавляем плиту 0.3 м (записываем в PLATES_0_32)
            for _ in range(max(0, qty)):
                PLATES_0_32.append(length_rounded)
                # Сохраняем точную ширину 0.3м (попадает в диапазон 0.26-0.32)
                PLATE_EXACT_WIDTHS[(length_rounded, 'PLATES_0_32')] = 0.3
            ldr_norm = (length_dm_raw or "").strip()
            if load_code is not None and load_code > 0:
                width_rounded = round(width_m, 3)
                key_new = (length_rounded, width_rounded, load_code, ldr_norm)
                PLATE_LOAD_DETAILS[key_new] = PLATE_LOAD_DETAILS.get(key_new, 0) + qty
                PLATE_LENGTH_DM_RAW[key_new] = ldr_norm
                if line_idx is not None and 0 <= line_idx < len(line_plate_load_details):
                    _ld = line_plate_load_details[line_idx]
                    _ld[key_new] = _ld.get(key_new, 0) + qty
                if line_idx is not None:
                    lc = float(load_code)
                    _record_contribution(line_idx, length_rounded, 1.2, lc, ldr_norm)
                    _record_contribution(line_idx, length_rounded, 0.3, lc, ldr_norm)
            elif line_idx is not None:
                _record_contribution(line_idx, length_rounded, 1.2, None, ldr_norm)
                _record_contribution(line_idx, length_rounded, 0.3, None, ldr_norm)
            return

        target = None
        target_name = None  # Имя списка для сохранения точной ширины
        
        # Стандартные ширины плит
        if 1.15 <= width_m <= 1.25:
            target = PLATES_1_2
            target_name = 'PLATES_1_2'
        elif 0.98 <= width_m <= 1.02:
            target = PLATES_1_0
            target_name = 'PLATES_1_0'
        elif 1.02 <= width_m <= 1.08:   # по таблице завода: рез 1020–1080 мм
            target = PLATES_1_08
            target_name = 'PLATES_1_08'
        # Основные части (по таблице допустимых резов: 260-320, 460-530, 660-720, 860-920):
        elif 0.26 <= width_m <= 0.32:    # 260-320 мм
            target = PLATES_0_32
            target_name = 'PLATES_0_32'
        elif 0.46 <= width_m <= 0.53:    # 460-530 мм
            target = PLATES_0_46
            target_name = 'PLATES_0_46'
        elif 0.66 <= width_m <= 0.71:    # 660-710 мм → PLATES_0_70
            target = PLATES_0_70
            target_name = 'PLATES_0_70'
        elif 0.71 < width_m <= 0.72:     # 710-720 мм → PLATES_0_72
            target = PLATES_0_72
            target_name = 'PLATES_0_72'
        elif 0.86 <= width_m <= 0.92:    # 860-920 мм
            target = PLATES_0_86
            target_name = 'PLATES_0_86'
        # Остатки по таблице завода (добор): остаток от 860–920 = 260–320 (попадает в 0_32 выше)
        # 340 мм по таблице не входит в допустимый остаток — не выделяем отдельно
        elif 0.47 < width_m <= 0.49:     # ~480 мм (остаток от 720)
            target = PLATES_0_48
            target_name = 'PLATES_0_48'
        elif 0.49 < width_m <= 0.51:     # ~500 мм (остаток от 700)
            target = PLATES_0_50
            target_name = 'PLATES_0_50'
        # По таблице остаток от реза 460–530 = 660–720 мм (попадает в 0_70/0_72 выше), 740 не используем
        elif 0.87 < width_m <= 0.89:     # ~880 мм (остаток от 320)
            target = PLATES_0_88
            target_name = 'PLATES_0_88'
        else:
            # Здесь ширина не попала ни в один диапазон.
            # Применяем правило: «берём меньший рез».
            # Допустимые ширины по таблице завода (информ. письмо): резы 260–320, 460–530, 660–720, 860–920, 1020–1080 мм
            STANDARD_WIDTHS = [
                0.20, 0.30,            # специальные ленты
                0.32,                  # рез 300 (-40;+20) = 260–320
                0.46, 0.48,            # рез 500 и остаток
                0.50, 0.53,            # остаток ~500 и рез до 530
                0.70, 0.72,            # рез 700 и остаток (740 по таблице не отдельный остаток)
                0.86, 0.88, 0.92,      # рез 900 и остатки
                1.00, 1.02, 1.08, 1.20,  # рез 1020–1080 (1.02–1.08), целая 1.2
            ]
            # Берём максимальную стандартную ширину, не превышающую фактическую
            candidates = [w for w in STANDARD_WIDTHS if w <= width_m + 1e-6]
            if not candidates:
                # Слишком узкая или совсем нестандартная плита — игнорируем
                return
            snapped_width = max(candidates)

            # Рекурсивный вызов с притянутой шириной, чтобы сработали диапазоны выше.
            add_items(snapped_width, length_m, qty, load_code, length_dm_raw=length_dm_raw, line_idx=line_idx)
            return

        # Добавляем плиты в список и сохраняем точную ширину (длина с точностью 3 знака)
        length_rounded = round(float(length_m), 3)
        for _ in range(max(0, qty)):
            target.append(length_rounded)
            
            # Сохраняем точную ширину для этой плиты
            if target_name:
                key = (length_rounded, target_name)
                PLATE_EXACT_WIDTHS[key] = round(width_m, 3)
        
        # Сохраняем нагрузку (если указана) в PLATE_LOAD_DETAILS и raw в PLATE_LENGTH_DM_RAW
        width_rounded = round(width_m, 3)
        ldr_norm = (length_dm_raw or "").strip()
        if load_code is not None and load_code > 0:
            key_new = (length_rounded, width_rounded, load_code, ldr_norm)
            PLATE_LOAD_DETAILS[key_new] = PLATE_LOAD_DETAILS.get(key_new, 0) + qty
            PLATE_LENGTH_DM_RAW[key_new] = ldr_norm
            if line_idx is not None and 0 <= line_idx < len(line_plate_load_details):
                _ld = line_plate_load_details[line_idx]
                _ld[key_new] = _ld.get(key_new, 0) + qty
            if line_idx is not None:
                _record_contribution(line_idx, length_rounded, width_rounded, float(load_code), ldr_norm)
        elif line_idx is not None:
            _record_contribution(line_idx, length_rounded, width_rounded, None, ldr_norm)

    # Список нераспознанных строк для отчёта пользователю
    unparsed_lines = []

    for line_idx, raw in enumerate(lines):
        parsed = False
        parsed_line = parse_line(raw)
        diag: dict[str, Any] = {
            "raw_input": raw,
            "parse_stage": parsed_line.stage,
            "recognized_by": "parser",
        }

        if not parsed_line.parsed:
            diag["validation_status"] = "failed"
            diag["reason_code"] = parsed_line.reason_code or "pattern_not_matched"
            diag["rejection_reason"] = parsed_line.reason_text or "строка не распознана"
            LAST_PARSE_DIAGNOSTICS.append(diag)
            unparsed_lines.append(f"{raw} (пропущено: {diag['rejection_reason']})")
            continue

        validation = validate_plate_values(parsed_line.width_m, parsed_line.length_m, parsed_line.qty)
        if not validation.ok:
            diag["validation_status"] = "failed"
            diag["reason_code"] = validation.reason_code
            diag["rejection_reason"] = validation.reason_text
            LAST_PARSE_DIAGNOSTICS.append(diag)
            unparsed_lines.append(f"{raw} (пропущено: {validation.reason_text})")
            continue

        add_items(
            parsed_line.width_m,
            parsed_line.length_m,
            parsed_line.qty,
            parsed_line.load_code,
            length_dm_raw=parsed_line.length_dm_raw,
            line_idx=line_idx,
        )
        parsed = True
        diag["validation_status"] = "ok"
        diag["normalized_input"] = raw
        LAST_PARSE_DIAGNOSTICS.append(diag)
        if parsed:
            continue

    _recompute_totals_from_lists()

    # Заполняем кэш номенклатуры один раз по всем ключам PLATE_LOAD_DETAILS + PLATE_LENGTH_DM_RAW
    try:
        from core.kp_db import fill_plate_nomenclature_cache
        fill_plate_nomenclature_cache()
    except Exception as _e:
        logger.warning(f"Не удалось заполнить PLATE_NOMENCLATURE_CACHE: {_e}")

    # Логируем нераспознанные строки для отладки
    if unparsed_lines:
        logger.warning(
            f"Парсинг завершён с {len(unparsed_lines)} нераспознанными строками. "
            f"Всего строк обработано: {len(lines)}"
        )
        for i, line in enumerate(unparsed_lines[:5], 1):  # Показываем первые 5
            logger.debug(f"  Нераспознанная строка {i}: {line}")
        if len(unparsed_lines) > 5:
            logger.debug(f"  ... и ещё {len(unparsed_lines) - 5} строк")
    else:
        logger.info(f"Парсинг завершён успешно. Обработано строк: {len(lines)}")
    
    return unparsed_lines, line_contributions, line_plate_load_details


def format_reinforcement_from_load_code(load_code: float | int) -> str:
    """Преобразует код нагрузки (8/10/12/12.5/11/6...) в суффикс вида '8п', '10п', '12п', '12,5п'.
    
    ВАЖНО: 12.5 отображается как '12,5п' (с запятой), но считается по цене как 12п.
    """
    try:
        code = float(load_code)
    except Exception:
        code = 8.0
    if code <= 0:
        code = 8.0
    
    # Проверяем, дробное ли число (например, 12.5)
    if abs(code - int(code)) < 1e-6:
        # Целое число: 8.0 → "8п"
        return f"{int(code)}п"
    else:
        # Дробное число: 12.5 → "12,5п" (с запятой, как в России)
        return f"{code:.1f}п".replace('.', ',')


def make_plate_name(
    length_m: float,
    width_m: float,
    reinforcement: str = '8п',
    load_code: int | None = None,
    length_dm_raw: str | None = None,
) -> str:
    """Формирует строку наименования в стиле прайса: 'Плиты ПБ 63-12-8п'.
    Ширина всегда в дециметрах: 1.2м→'12', 0.3м→'3', 0.2м→'2'.

    Если передан load_code (6/8/10/12/11...), он переопределяет reinforcement.
    Если передан length_dm_raw (например "59,81"), он используется для части длины в марке.
    """
    if load_code is not None:
        reinforcement = format_reinforcement_from_load_code(load_code)

    # Длина в марке: используем length_dm_raw если передан, иначе вычисляем из length_m
    if length_dm_raw and str(length_dm_raw).strip():
        length_str = str(length_dm_raw).strip().replace('.', ',')
    else:
        length_dm_val = length_m * 10
        branch_001 = abs(length_dm_val - round(length_dm_val)) < 0.01
        if branch_001:
            length_str = str(int(round(length_dm_val)))
        else:
            length_str = f'{length_dm_val:.1f}'.rstrip('0').rstrip('.').replace('.', ',')
        # #region agent log (57/57,1: где теряется сотка по длине)
        if 5.69 <= length_m <= 5.73:
            try:
                _log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'debug-8e9428.log')
                with open(_log_path, 'a', encoding='utf-8') as _f:
                    _f.write(__import__('json').dumps({"sessionId": "8e9428", "hypothesisId": "H_length_001", "location": "config_and_data:make_plate_name", "message": "57/57,1: length_m -> length_str", "data": {"length_m": length_m, "length_dm_raw": length_dm_raw, "length_dm_val": length_dm_val, "length_str": length_str, "branch_001_used": branch_001}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
            except Exception:
                pass
        # #endregion
        # #region agent log
        try:
            _log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'debug-b59370.log')
            with open(_log_path, 'a', encoding='utf-8') as _f:
                _f.write(__import__('json').dumps({"sessionId": "b59370", "hypothesisId": "H2", "location": "config_and_data:make_plate_name", "message": "length_str from length_m (no length_dm_raw)", "data": {"length_m": length_m, "length_dm_val": length_dm_val, "length_str": length_str}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
    
    # Единая логика формирования ширины: всё в дециметрах.
    # 1.2м → '12', 0.53м → '5,3', 0.3м → '3', 0.2м → '2'
    # (parse_pb_width_to_m по-прежнему понимает старый формат "0.3"/"0.2" для обратной совместимости)
    width_dm = round(width_m * 10, 2)
    if abs(width_dm - round(width_dm)) < 1e-6:
        width_str = str(int(round(width_dm)))
    else:
        width_str = f'{width_dm:.2f}'.rstrip('0').rstrip('.').replace('.', ',')
    return f'Плиты ПБ {length_str}-{width_str}-{reinforcement}'


def parse_name_to_sizes(name: str) -> tuple:
    """Достаёт (length_m, width_m) из строки прайса.
    Длина — по правилу length_dm_to_m (целое → номинал в дм, длина = номинал/10 м; с запятой/точкой → дм/10).
    Ширина: дециметры делим на 10. Обратная совместимость: '0.3'/'0.2' распознаются как метры.
    Примеры: '39-12' → (3.9, 1.2); '38,9-12' → (3.89, 1.2); '74-3-8п' → (7.4, 0.3)."""
    s = name.replace(',', '.')
    m = re.search(r'(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)', s)
    if not m:
        return None, None
    length_m = length_dm_to_m(m.group(1))
    width_m = parse_pb_width_to_m(m.group(2))
    return length_m, width_m


def plate_name_to_prays_variant(name: str) -> Optional[str]:
    """Возвращает вариант имени плиты для поиска в prays_plity.

    Бот формирует ленты 0.3 м и 0.2 м с шириной в метрах: '-0.3-' / '-0.2-'.
    В справочнике prays_plity те же плиты записаны в дециметрах: '-3,0-' / '-2,0-'.
    Функция подставляет вариант справочника, чтобы lookup мог найти запись.

    Также: бот пишет целую ширину 5/7/9 дм как '-5-' / '-7-' / '-9-', а в prays_plity
    они хранятся как '-5,0-' / '-7,0-' / '-9,0-'.

    Возвращает строку с заменённой шириной, или None если замена неприменима
    (т.е. ширина в марке уже в дм или это не лента 0.3/0.2 м).

    Примеры:
      'Плиты ПБ 42-0.3-8п'  → 'Плиты ПБ 42-3,0-8п'
      'Плиты ПБ 25-0.2-8п'  → 'Плиты ПБ 25-2,0-8п'
      'Плиты ПБ 61,8-5-8п'  → 'Плиты ПБ 61,8-5,0-8п'
      'Плиты ПБ 45-7-6п'    → 'Плиты ПБ 45-7,0-6п'
      'Плиты ПБ 37,9-9-8п'  → 'Плиты ПБ 37,9-9,0-8п'
      'Плиты ПБ 63-12-8п'   → None
      'Плиты ПБ 42-3,0-8п'  → None  (уже вариант справочника — не трогаем)
    """
    # Ленты 0.3 / 0.2 м: бот пишет -0.3-, в prays -3,0-
    variant = re.sub(r'(?<=-)0\.3(?=-)', '3,0', name)
    if variant != name:
        return variant
    variant = re.sub(r'(?<=-)0\.2(?=-)', '2,0', name)
    if variant != name:
        return variant

    # Целая ширина 5/7/9 дм: бот пишет -5-, в prays -5,0-
    for w in ('5', '7', '9'):
        pattern = rf'(?<=-){re.escape(w)}(?=-)'
        replacement = f'{w},0'
        variant = re.sub(pattern, replacement, name)
        if variant != name:
            return variant

    return None


def _apply_width_prays_variant(name: str) -> str:
    """Подставляет варианты ширины для prays_plity: -0.3-→-3,0-, -0.2-→-2,0-, -N-→-N,0- (N=2..9, одна цифра между дефисами)."""
    v = re.sub(r'(?<=-)0\.3(?=-)', '3,0', name)
    v = re.sub(r'(?<=-)0\.2(?=-)', '2,0', v)
    # Одна цифра ширины между дефисами: -7- → -7,0-
    v = re.sub(r'-([2-9])-', r'-\1,0-', v)
    return v


def _apply_length_prays_variant(name: str) -> str:
    """Целая длина в марке: ПБ 45- → ПБ 45,0- (только первое число после ПБ)."""
    return re.sub(r'((?:Плиты\s+)?П[БК]\s+)(\d+)(?=-)', r'\1\2,0', name, count=1)


def plate_name_to_prays_variants(name: str) -> List[str]:
    """Возвращает список вариантов имени плиты для поиска в prays_plity.

    При отсутствии точного совпадения lookup пробует каждый вариант по очереди.
    Порядок: только ширина, только длина, оба (ширина + длина).
    """
    result: List[str] = []
    seen: set = set()

    def add(v: str) -> None:
        if v and v != name and v not in seen:
            seen.add(v)
            result.append(v)

    width_fixed = _apply_width_prays_variant(name)
    add(width_fixed)

    length_fixed = _apply_length_prays_variant(name)
    add(length_fixed)

    both = _apply_length_prays_variant(width_fixed)
    add(both)

    return result


def parse_load_code_from_name(name: str, default: int = 8) -> int:
    """
    Извлекает код нагрузки (6/8/10/12/...) из строки вида 'Плиты ПБ 71-12-10п'.

    Возвращает целое число (например, 8, 10, 12).
    Если не удалось распознать нагрузку — возвращает default.

    Примеры:
      'Плиты ПБ 71-12-8п'   -> 8
      'ПБ 69-12-12,5п'      -> 12
      'ПБ 141-12-11п'       -> 11
    """
    s = str(name).lower().replace(',', '.')

    # Ищем последнюю часть перед буквой "п": ...-8п, ...-10п, ...-12.5п
    m = re.search(r'-\s*([\d\.]+)\s*п\b', s)
    if not m:
        return default
    try:
        val = float(m.group(1))
    except ValueError:
        return default

    # Округляем до целого кода нагрузки:
    # 8.0 -> 8, 10.0 -> 10, 12.5 -> 13, 11.0 -> 11 (всегда вверх для .5)
    load_code = int(val + 0.5)
    if load_code <= 0:
        return default
    return load_code


def extract_length_dm_raw_from_plate_name(plate_name: str) -> str | None:
    """
    Извлекает подстроку длины из марки плиты для сохранения в length_dm_raw.

    Примеры:
      'Плиты ПБ 59,8-12-8п'   -> '59,8'
      'Плиты ПБ 61,2-12-8п'   -> '61,2'
      'ПБ 78-12-8п'           -> '78'
    """
    if not plate_name or not str(plate_name).strip():
        return None
    m = re.search(r'(?:Плиты\s+)?П[БК]\s*([\d,\.]+)\s*-', str(plate_name), re.IGNORECASE)
    if not m:
        return None
    raw = m.group(1).strip().replace('.', ',')
    return raw if raw else None


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
    
    # Если почти целое — возвращаем int
    if abs(val - round(val)) < 1e-6:
        return int(round(val))
    
    # Оставляем дробный код (например, 12.5)
    return round(val, 1)


def canonical_plate_key(length, width, load_code) -> tuple:
    """
    Единственный способ создать ключ плиты во всём проекте.

    Нормализует длину (round 2 знака), ширину (int мм) и код нагрузки
    через normalize_load_code, чтобы ключи были сравнимы независимо от источника.

    Примеры:
    - canonical_plate_key(5.700001, 1200.0, 800) == (5.7, 1200, 8)
    - canonical_plate_key(5.71, 530, '8')       == (5.71, 530, 8)
    """
    return (
        round(float(length), 2),
        int(round(float(width))),
        normalize_load_code(load_code, default=8),
    )


def load_code_for_price_match(value, default: int = 8) -> int:
    """
    Код нагрузки для сопоставления при подборе цены в КП.
    Для 12.5 и 13 возвращает 12 (в БД и прайсе только 12), для остальных — floor(normalize_load_code).
    """
    n = normalize_load_code(value, default=default)
    if n is None:
        return default
    try:
        v = float(n)
    except (TypeError, ValueError):
        return default
    if abs(v - 12.5) < 1e-6 or abs(v - 13) < 1e-6:
        return 12
    return int(math.floor(v))


def get_load_code_for_plate(length_m: float, width_m: float, default: int = 8) -> int:
    """
    Возвращает код нагрузки для плиты по (длина, ширина).

    Логика:
      1) Ищем в PLATE_LOAD_DETAILS — самая частая нагрузка для этих размеров;
      2) если ничего не нашли — fallback: 6 для узких плит (<1.0 м) или default для широких.
    """
    try:
        key_base = (round(float(length_m), 3), round(float(width_m), 3))
    except Exception:
        return 6 if (isinstance(width_m, (int, float)) and float(width_m) < 1.0) else default

    # Ищем в PLATE_LOAD_DETAILS (самая частая нагрузка для этих размеров)
    matching_loads = []
    for key, qty in PLATE_LOAD_DETAILS.items():
        L, W, load = key[0], key[1], key[2]
        if abs(L - key_base[0]) <= 0.005 and abs(W - key_base[1]) <= 0.005:
            matching_loads.append((load, qty))
    if matching_loads:
        most_common_load = max(matching_loads, key=lambda x: x[1])[0]
        return most_common_load

    # Fallback по ширине
    try:
        w_val = float(width_m)
    except Exception:
        w_val = 1.2
    if w_val < 1.0:
        return 6
    return default


def get_exact_width(length_m: float, target_list_name: str, default_width: float) -> float:
    """
    Возвращает точную ширину плиты, если она была сохранена при парсинге.
    Иначе возвращает дефолтное значение (среднее/минимальное для диапазона).
    
    Args:
        length_m: Длина плиты в метрах
        target_list_name: Имя списка ('PLATES_0_46', 'PLATES_0_32', ...)
        default_width: Дефолтная ширина в метрах (0.46 для PLATES_0_46)
    
    Returns:
        Точная ширина в метрах (например, 0.53 вместо 0.46)
    
    Example:
        >>> # Плита "ПБ 28-5,3-8п" была добавлена в PLATES_0_46
        >>> get_exact_width(2.8, 'PLATES_0_46', 0.46)
        0.53  # Точная ширина 530мм, а не 460мм!
    """
    key = (round(float(length_m), 3), target_list_name)
    return PLATE_EXACT_WIDTHS.get(key, default_width)


def approximate_weight_kg(length_m: float, width_m: float, thickness_m: float = 0.22) -> float:
    """
    Расчет веса плиты в килограммах по формуле в дециметрах.

    Формула: WEIGHT_KG_PER_DM2 * length_dm * width_dm.
    Аргумент thickness_m оставлен для обратной совместимости сигнатуры.
    """
    _ = thickness_m  # совместимость с существующими вызовами
    length_dm = float(length_m) * 10.0
    width_dm = float(width_m) * 10.0
    return round(WEIGHT_KG_PER_DM2 * length_dm * width_dm, 1)


def register_plate_metadata(plates: List[Dict[str, Any]]) -> None:
    """Регистрирует метаданные плит перед визуализацией."""
    PLATE_METADATA.clear()
    for plate in plates:
        try:
            length = round(float(plate.get('length_m', 0)), 2)
            width_mm = int(plate.get('width_mm', 0))
        except (TypeError, ValueError):
            continue
        entry = {
            'forming_week': plate.get('forming_week'),
            'contractor': plate.get('contractor'),
            'name': plate.get('name'),
        }
        PLATE_METADATA.setdefault((length, width_mm), []).append(entry)


def consume_plate_metadata(length_m: float, width_mm: int, qty: int) -> List[Dict[str, Any]]:
    """Возвращает и удаляет из буфера метаданные, соответствующие плитам."""
    key = (round(float(length_m), 2), int(width_mm))
    bucket = PLATE_METADATA.get(key, [])
    taken = bucket[:qty]
    PLATE_METADATA[key] = bucket[qty:]
    return taken


def clear_plate_metadata() -> None:
    """Полностью очищает буфер метаданных плит."""
    PLATE_METADATA.clear()





