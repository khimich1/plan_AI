#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конфигурация и данные проекта:
- Константы (размеры дорожки, цены резов)
- Глобальные списки плит
- Парсинг текста пользователя
"""
import os
import re
from typing import Any, Dict, List, Tuple

# ==================== КОНСТАНТЫ ====================

TRACK_LENGTH_M = 101.0
TRACK_WIDTH_M = 1.2

# Пути к прайсам (делаем абсолютными относительно файла скрипта)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRICE_XLSX_PATH = os.path.join(BASE_DIR, 'банк знаний', 'Новые цены для прайса с 19.08.24.xlsx')
CUTS_DOCX_PATH = os.path.join(BASE_DIR, 'банк знаний', 'Письмо Цены с 29.05.2024 цены на резы.docx')
PRICE_DB_PATH = os.path.join(BASE_DIR, 'pb.db')

# Стоимость резов
LONG_CUT_PRICE_PER_M = 460.0  # Продольный рез, руб/пог.м
TRANSVERSE_CUT_PRICE = 1200.0  # Поперечный (или скошенный) рез, руб/шт

# ==================== ГЛОБАЛЬНЫЕ СПИСКИ ПЛИТ ====================

# Данные из согласованного КЗ-плана
# 1) Плиты 1.2 м — без резов (новый заказ)
PLATES_1_2 = [3.39]*2

# Дополнительные целевые ширины, которые получаем продольным резом из 1.2 м
PLATES_1_08 = []         # нет 1.08 в этом заказе
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


# ==================== ФУНКЦИИ ПАРСИНГА ====================

def _clear_all_plate_lists():
    """Очищает все глобальные списки плит"""
    global PLATES_1_2, PLATES_1_5_TO_1_2, PLATES_1_0, PLATES_1_08
    global PLATES_0_46, PLATES_0_32, PLATES_0_72, PLATES_0_70, PLATES_0_86
    global PLATES_0_74, PLATES_0_88, PLATES_0_48, PLATES_0_50, PLATES_0_34
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


def set_plate_lists_from_text(user_text: str) -> None:
    """Парсит свободный текст пользователя и заполняет списки PLATES_*.

    Поддерживаем форматы:
      - "1.2×3.39 — 2 шт" / "0,32x6,63 - 4"
      - "Плиты ПБ 78-12-8п 3" (длина в дм, ширина 12 => 1.2м, количество 3)
    Неизвестные ширины игнорируем.
    """
    _clear_all_plate_lists()

    text = (user_text or '').replace('\u00d7', 'x').replace('×', 'x')
    lines = [l.strip() for l in re.split(r'[\n;]+', text) if l.strip()]

    def add_items(width_m: float, length_m: float, qty: int):
        # Специальная обработка плит 1.5 м → заменяем на 1.2 м + 0.3 м
        if 1.45 <= width_m <= 1.55:  # 1.5 м (диапазон ±50 мм)
            # Добавляем плиту 1.2 м
            for _ in range(max(0, qty)):
                PLATES_1_2.append(round(float(length_m), 2))
            # Добавляем плиту 0.3 м (записываем в PLATES_0_32)
            for _ in range(max(0, qty)):
                PLATES_0_32.append(round(float(length_m), 2))
            return
        
        target = None
        if 1.15 <= width_m <= 1.25:
            target = PLATES_1_2
        elif 0.98 <= width_m <= 1.02:
            target = PLATES_1_0
        elif 1.06 <= width_m <= 1.12:
            target = PLATES_1_08
        # Основные части (по таблице допустимых резов: 260-320, 460-530, 660-720, 860-920):
        elif 0.26 <= width_m <= 0.32:    # 260-320 мм
            target = PLATES_0_32
        elif 0.46 <= width_m <= 0.53:    # 460-530 мм
            target = PLATES_0_46
        elif 0.66 <= width_m <= 0.71:    # 660-710 мм → PLATES_0_70
            target = PLATES_0_70
        elif 0.71 < width_m <= 0.72:     # 710-720 мм → PLATES_0_72
            target = PLATES_0_72
        elif 0.86 <= width_m <= 0.92:    # 860-920 мм
            target = PLATES_0_86
        # Остатки (если пользователь явно указал остаточные ширины):
        # Примечание: остатки обычно создаются автоматически оптимизатором
        elif 0.33 < width_m <= 0.35:     # ~340 мм (остаток от 860)
            target = PLATES_0_34
        elif 0.47 < width_m <= 0.49:     # ~480 мм (остаток от 720)
            target = PLATES_0_48
        elif 0.49 < width_m <= 0.51:     # ~500 мм (остаток от 700)
            target = PLATES_0_50
        elif 0.73 < width_m <= 0.75:     # ~740 мм (остаток от 460)
            target = PLATES_0_74
        elif 0.87 < width_m <= 0.89:     # ~880 мм (остаток от 320)
            target = PLATES_0_88
        else:
            return
        for _ in range(max(0, qty)):
            target.append(round(float(length_m), 2))

    for raw in lines:
        s = raw.lower()
        # 1) формат WxL x qty (поддерживает запятую и точку)
        s_norm = s.replace(',', '.')
        m = re.search(r'(\d+(?:\.\d+)?)\s*[xх]\s*(\d+(?:\.\d+)?)\D*(\d+)?', s)
        if m:
            w = float(m.group(1).replace(',', '.'))
            L = float(m.group(2).replace(',', '.'))
            q = int((m.group(3) or '1').replace(',', '.'))
            add_items(w, L, q)
            continue
        # 2) формат "Плиты ПБ 78,3-3,2-8п 3" или "ПБ 78-12-8п 10"
        m2 = re.search(r'плиты?\s*пб\s*([\d\.,]+)\s*-\s*([\d\.,]+)', s)
        if not m2:
            m2 = re.search(r'\bпб\s*([\d\.,]+)\s*-\s*([\d\.,]+)', s)
        if m2:
            Ldm_str = m2.group(1).replace(' ', '').replace(',', '.')
            Wdm_str = m2.group(2).replace(' ', '').replace(',', '.')
            try:
                # В нотации ПБ длина указывается в дм. Всегда делим на 10 → метры
                L = float(Ldm_str) / 10.0
            except Exception:
                continue
            try:
                # Ширина также в дм (12 → 1.2; 3.2 → 0.32; 8.6 → 0.86)
                W = float(Wdm_str) / 10.0
            except Exception:
                continue
            q = 1
            # Количество — последнее число в строке
            mq = re.search(r'(\d+)\s*(шт)?\s*$', s)
            if mq:
                try:
                    q = int(mq.group(1))
                except Exception:
                    q = 1
            add_items(W, L, q)
            continue

    _recompute_totals_from_lists()


def make_plate_name(length_m: float, width_m: float, reinforcement: str = '8п') -> str:
    """Формирует строку наименования в стиле прайса: 'Плиты ПБ 63-12-8п'.
    Для лент 0.3/0.2 записывает ширину как '0.3'/'0.2'."""
    length_dm = int(round(length_m * 10))
    if width_m < 0.5:
        width_str = '0.3' if abs(width_m - 0.3) < 1e-6 else ('0.2' if abs(width_m - 0.2) < 1e-6 else f'{width_m:.1f}'.replace('.', ','))
    else:
        width_dm = int(round(width_m * 10))
        width_str = str(width_dm)
    return f'Плиты ПБ {length_dm}-{width_str}-{reinforcement}'


def parse_name_to_sizes(name: str) -> tuple:
    """Достаёт (length_m, width_m) из строки прайса."""
    m = re.search(r'(\d+)-(\d+)', name.replace(',', '.'))
    if not m:
        return None, None
    return float(m.group(1)) / 10.0, float(m.group(2)) / 10.0


def approximate_weight_kg(length_m: float, width_m: float, thickness_m: float = 0.22) -> float:
    """Примерный расчёт веса плиты в килограммах"""
    volume = length_m * width_m * thickness_m
    return round(volume * 2400, 1)


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





