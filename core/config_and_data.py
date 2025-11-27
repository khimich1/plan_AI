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

# Пути к прайсам (BASE_DIR указывает на корень проекта, на уровень выше core/)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PRICE_XLSX_PATH = os.path.join(BASE_DIR, 'банк знаний', 'Новые цены для прайса с 19.08.24.xlsx')
CUTS_DOCX_PATH = os.path.join(BASE_DIR, 'банк знаний', 'Письмо Цены с 29.05.2024 цены на резы.docx')
PRICE_DB_PATH = os.path.join(BASE_DIR, 'pb.db')

# Стоимость резов
LONG_CUT_PRICE_PER_M = 460.0  # Продольный рез, руб/пог.м
TRANSVERSE_CUT_PRICE = 1200.0  # Поперечный (или скошенный) рез, руб/шт

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

# Карта нагрузок по (длина, ширина) → load_code (6/8/10/12/11...)
# СТАРАЯ ВЕРСИЯ: хранит только последнее значение нагрузки для (длина, ширина)
# Заполняется на этапе парсинга текста пользователя.
PLATE_LOAD_MAP: Dict[Tuple[float, float], int] = {}

# НОВАЯ ВЕРСИЯ: Детальная карта плит с нагрузкой
# Формат: (длина, ширина, нагрузка) → количество
# Пример: {(7.3, 1.2, 8): 93, (7.3, 1.2, 10): 4} - 93 плиты 8п + 4 плиты 10п
PLATE_LOAD_DETAILS: Dict[Tuple[float, float, int], int] = {}

# Карта точных размеров плит: (длина, имя_списка) → точная_ширина_в_метрах
# Сохраняет ТОЧНУЮ ширину каждой плиты при парсинге заказа.
# Пример: {(2.8, 'PLATES_0_46'): 0.53} - плита ПБ 28-5,3-8п имеет ширину 530мм, а не 460мм
PLATE_EXACT_WIDTHS: Dict[Tuple[float, str], float] = {}


# ==================== ФУНКЦИИ ПАРСИНГА ====================

def _clear_all_plate_lists():
    """Очищает все глобальные списки плит"""
    global PLATES_1_2, PLATES_1_5_TO_1_2, PLATES_1_0, PLATES_1_08
    global PLATES_0_46, PLATES_0_32, PLATES_0_72, PLATES_0_70, PLATES_0_86
    global PLATES_0_74, PLATES_0_88, PLATES_0_48, PLATES_0_50, PLATES_0_34
    global PLATE_LOAD_MAP, PLATE_LOAD_DETAILS, PLATE_EXACT_WIDTHS
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
    PLATE_LOAD_MAP.clear()
    PLATE_LOAD_DETAILS.clear()
    PLATE_EXACT_WIDTHS.clear()


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


def set_plate_lists_from_text(user_text: str) -> list[str]:
    """Парсит свободный текст пользователя и заполняет списки PLATES_*.

    Поддерживаем форматы:
      - "1.2×3.39 — 2 шт" / "0,32x6,63 - 4"
      - "Плиты ПБ 78-12-8п 3" (длина в дм, ширина 12 => 1.2м, количество 3)
    Неизвестные ширины игнорируем.
    
    Возвращает список нераспознанных строк.
    """
    _clear_all_plate_lists()

    text = (user_text or '').replace('\u00d7', 'x').replace('×', 'x')
    lines = [l.strip() for l in re.split(r'[\n;]+', text) if l.strip()]

    def normalize_dimension(value_str: str) -> float:
        """
        Нормализует размер плиты (длину или ширину) из строки.
        
        Проблема: пользователи пишут по-разному:
        - "5,30" (правильно: 5.30 дм)
        - "530" (неправильно, но имели в виду 5.30 дм)
        - "665" (имели в виду 6.65 дм)
        - "6,65" (правильно: 6.65 дм)
        
        Логика:
        1) Заменяем запятую на точку
        2) Если получилось число > 20 дм, значит забыли запятую:
           530 → 5.30, 665 → 6.65, 395 → 3.95
        3) Возвращаем значение в дециметрах
        """
        # Убираем пробелы и заменяем запятую на точку
        clean = value_str.strip().replace(' ', '').replace(',', '.')
        try:
            val = float(clean)
        except ValueError:
            return 0.0
        
        # Если число больше 20 дм (2 метра), скорее всего забыли запятую
        # Например: 530 должно быть 5.30, 665 должно быть 6.65
        if val > 20.0:
            # Делим на 100, чтобы получить дециметры
            # 530 → 5.30, 665 → 6.65, 395 → 3.95
            val = val / 100.0
        
        return val

    def add_items(width_m: float, length_m: float, qty: int, load_code: int = None):
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
        """
        # Специальная обработка плит 1.5 м → заменяем на 1.2 м + 0.3 м
        if 1.45 <= width_m <= 1.55:  # 1.5 м (диапазон ±50 мм)
            length_rounded = round(float(length_m), 2)
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
        elif 1.06 <= width_m <= 1.12:
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
        # Остатки (если пользователь явно указал остаточные ширины):
        # Примечание: остатки обычно создаются автоматически оптимизатором
        elif 0.33 < width_m <= 0.35:     # ~340 мм (остаток от 860)
            target = PLATES_0_34
            target_name = 'PLATES_0_34'
        elif 0.47 < width_m <= 0.49:     # ~480 мм (остаток от 720)
            target = PLATES_0_48
            target_name = 'PLATES_0_48'
        elif 0.49 < width_m <= 0.51:     # ~500 мм (остаток от 700)
            target = PLATES_0_50
            target_name = 'PLATES_0_50'
        elif 0.73 < width_m <= 0.75:     # ~740 мм (остаток от 460)
            target = PLATES_0_74
            target_name = 'PLATES_0_74'
        elif 0.87 < width_m <= 0.89:     # ~880 мм (остаток от 320)
            target = PLATES_0_88
            target_name = 'PLATES_0_88'
        else:
            # Здесь ширина не попала ни в один диапазон.
            # Применяем правило: «берём меньший рез».
            # Допустимые стандартные и остаточные ширины (в метрах).
            STANDARD_WIDTHS = [
                0.20, 0.30,            # специальные ленты
                0.32, 0.34,            # рез и остаток ~320 / 340
                0.46, 0.48,            # рез и остаток ~460 / 480
                0.50, 0.53,            # остаток ~500 и рез до 530
                0.70, 0.72, 0.74,      # рез и остаток ~700 / 720 / 740
                0.86, 0.88, 0.92,      # рез и остаток ~860 / 880 / 920
                1.00, 1.08, 1.20,      # стандартные ширины плит
            ]
            # Берём максимальную стандартную ширину, не превышающую фактическую
            candidates = [w for w in STANDARD_WIDTHS if w <= width_m + 1e-6]
            if not candidates:
                # Слишком узкая или совсем нестандартная плита — игнорируем
                return
            snapped_width = max(candidates)

            # Рекурсивный вызов с притянутой шириной, чтобы сработали диапазоны выше.
            add_items(snapped_width, length_m, qty, load_code)  # Передаём нагрузку дальше
            return

        # Добавляем плиты в список и сохраняем точную ширину
        for _ in range(max(0, qty)):
            length_rounded = round(float(length_m), 2)
            target.append(length_rounded)
            
            # Сохраняем точную ширину для этой плиты
            if target_name:
                key = (length_rounded, target_name)
                PLATE_EXACT_WIDTHS[key] = round(width_m, 3)
        
        # Сохраняем нагрузку (если указана) в PLATE_LOAD_DETAILS
        if load_code is not None and load_code > 0:
            length_rounded = round(float(length_m), 2)
            width_rounded = round(width_m, 3)
            
            # Сохраняем в старом формате для обратной совместимости
            key_old = (length_rounded, width_rounded)
            PLATE_LOAD_MAP[key_old] = load_code
            
            # Сохраняем в новом формате с количеством
            key_new = (length_rounded, width_rounded, load_code)
            PLATE_LOAD_DETAILS[key_new] = PLATE_LOAD_DETAILS.get(key_new, 0) + qty

    # Список нераспознанных строк для отчёта пользователю
    unparsed_lines = []

    for raw in lines:
        s = raw.lower()
        parsed = False
        
        # 1) формат WxL x qty (поддерживает запятую и точку)
        s_norm = s.replace(',', '.')
        m = re.search(r'(\d+(?:\.\d+)?)\s*[xх]\s*(\d+(?:\.\d+)?)\D*(\d+)?', s)
        if m:
            first = float(m.group(1).replace(',', '.'))
            second = float(m.group(2).replace(',', '.'))
            q = int((m.group(3) or '1').replace(',', '.'))

            # Пользователи часто пишут "длина × ширина". Если первая цифра
            # намного больше (в метрах), а вторая похожа на ширину — меняем местами.
            if first > 2.0 and second <= 1.5:
                width_m, length_m = second, first
            else:
                width_m, length_m = first, second

            add_items(width_m, length_m, q)
            parsed = True
            continue
        
        # 2) формат "Плиты ПБ 78,3-3,2-8п 3" или "ПБ 78-12-8п 10"
        #    а также "Плита ПК 80-12-8 шт 7", "ПК 80-12-8 7"
        #    (плитА/плитЫ + ПБ или ПК)
        m2 = re.search(r'плит[аы]?\s*п[бк]\s*([\d\.,]+)\s*-\s*([\d\.,]+)', s)
        if not m2:
            # Вариант без слова "плита/плиты": "ПБ 78-12-8п 3" / "ПК 80-12-8 7"
            m2 = re.search(r'\bп[бк]\s*([\d\.,]+)\s*-\s*([\d\.,]+)', s)
        if m2:
            Ldm_str = m2.group(1)
            Wdm_str = m2.group(2)
            
            # ДЛИНУ просто переводим в число (80 -> 8.0 м, 60 -> 6.0 м, 54,3 -> 5.43 м)
            # Для длины НЕ применяем нормализацию, так как значения типа 80, 60 - это нормальные дециметры
            try:
                Ldm = float(Ldm_str.replace(' ', '').replace(',', '.'))
            except ValueError:
                Ldm = 0.0
            
            # ШИРИНУ нормализуем (530 -> 5.30, 665 -> 6.65, 12 -> 12, 6,0 -> 6.0)
            # Только для ширины применяем автоисправление забытых запятых
            Wdm = normalize_dimension(Wdm_str)
            
            if Ldm <= 0 or Wdm <= 0:
                # Если не удалось распознать размеры, пропускаем строку
                continue
            
            # Переводим из дециметров в метры
            L = Ldm / 10.0
            W = round(Wdm / 10.0, 3)
            q = 1
            # Количество — число ПОСЛЕ маркировки "8п" (ищем после дефиса с "8п")
            # Сначала ищем маркировку "-8п" и берём число ПОСЛЕ неё
            mq = re.search(r'-\d+п\s*[-—–]\s*(\d+)', s)  # "ПБ 66,2-12-8п — 6"
            if not mq:
                # Если не нашли с дефисом, ищем просто после "-8п" до конца строки
                mq = re.search(r'-\d+п\s+(\d+)\s*(шт)?\s*$', s)  # "ПБ 66,2-12-8п 6"
            if not mq:
                # Последний вариант: число в конце строки (но НЕ в составе "-8п")
                # Используем negative lookbehind, чтобы не захватить "8" из "-8п"
                mq = re.search(r'(?<!-\d)(\d+)\s*(шт)?\s*$', s)
            if mq:
                try:
                    q = int(mq.group(1))
                except Exception:
                    q = 1

            # Пытаемся извлечь нагрузку из марки: ...-8п / -10п / -12,5п / -8 (без "п")
            load_code = None
            # Сначала пробуем с "п" (приоритет): "-8п", "-10п", "-12,5п"
            load_match = re.search(r'-\s*([\d\.,]+)\s*п\b', s)
            if not load_match:
                # Если не нашли с "п", ищем формат "-8 шт" или "-8" в конце строки
                # Это третье число после двух дефисов: ПБ 80-12-8 или ПБ 80-6,00-8
                load_match = re.search(r'п[бк]\s*[\d\.,]+\s*-\s*[\d\.,]+\s*-\s*([\d\.,]+)', s)
            
            if load_match:
                try:
                    load_val = float(load_match.group(1).replace(',', '.'))
                    # ВАЖНО: Сохраняем как float, чтобы 12.5 осталось 12.5 (не округляем до 13!)
                    # При группировке и ценах будем использовать math.floor()
                    load_code = load_val
                    if load_code <= 0:
                        load_code = None
                except Exception:
                    load_code = None

            # Передаём нагрузку в add_items
            add_items(W, L, q, load_code)
            parsed = True
            continue
        
        # Если строка не распознана ни одним паттерном, добавляем в список
        if not parsed:
            unparsed_lines.append(raw)

    _recompute_totals_from_lists()
    
    # Выводим в консоль нераспознанные строки для отладки
    if unparsed_lines:
        print(f"[PARSER WARNING] Нераспознанные строки ({len(unparsed_lines)}):")
        for line in unparsed_lines:
            print(f"  - {line}")
    
    return unparsed_lines


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
) -> str:
    """Формирует строку наименования в стиле прайса: 'Плиты ПБ 63-12-8п'.
    Для лент 0.3/0.2 записывает ширину как '0.3'/'0.2'.

    Если передан load_code (6/8/10/12/11...), он переопределяет reinforcement,
    чтобы в имени плиты отражалась фактическая нагрузка, как в заказе.
    """
    if load_code is not None:
        reinforcement = format_reinforcement_from_load_code(load_code)

    length_dm = int(round(length_m * 10))
    # Единая логика формирования ширины, как в боте (bot_handlers.py):
    # - плиты 1.2м / 1.08м / 1.0м → '12' / '10,8' / '10'
    # - узкие плиты 0.46 / 0.32 / 0.86 и т.п. → '4,6' / '3,2' / '8,6'
    # - специальные ленты 0.3 / 0.2 → '0.3' / '0.2'
    if abs(width_m - 0.3) < 1e-6:
        width_str = '0.3'
    elif abs(width_m - 0.2) < 1e-6:
        width_str = '0.2'
    else:
        # Переводим в дм (например, 0.46м → 4.6; 1.2м → 12.0)
        width_dm = round(width_m * 10, 2)  # Увеличили точность для 6.65
        # Умное форматирование: убираем лишние нули (6.65→"6,65", 5.3→"5,3", 12.0→"12")
        if abs(width_dm - round(width_dm)) < 1e-6:
            # Целое число (12.0 → "12")
            width_str = str(int(round(width_dm)))
        else:
            # Дробное число: убираем лишние нули (6.65→"6,65", 5.30→"5,3")
            width_str = f'{width_dm:.2f}'.rstrip('0').rstrip('.').replace('.', ',')
    return f'Плиты ПБ {length_dm}-{width_str}-{reinforcement}'


def parse_name_to_sizes(name: str) -> tuple:
    """Достаёт (length_m, width_m) из строки прайса."""
    m = re.search(r'(\d+)-(\d+)', name.replace(',', '.'))
    if not m:
        return None, None
    return float(m.group(1)) / 10.0, float(m.group(2)) / 10.0


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


def get_load_code_for_plate(length_m: float, width_m: float, default: int = 8) -> int:
    """
    Возвращает код нагрузки для плиты по (длина, ширина).

    Логика:
      1) если во время парсинга заказа мы видели эту плиту с явной нагрузкой
         (ПБ 71-12-10п, ПБ 69-12-12,5п и т.п.), берём код из PLATE_LOAD_MAP;
      2) если точного совпадения нет — ищем по близким значениям длины/ширины;
      3) если ничего не нашли — возвращаем запасной вариант:
         6 для узких плит (<1.0 м по ширине) или default (обычно 8) для широких.
    """
    try:
        key_base = (round(float(length_m), 2), round(float(width_m), 3))
    except Exception:
        # Если длина/ширина странные, просто используем fallback
        return 6 if (isinstance(width_m, (int, float)) and float(width_m) < 1.0) else default

    # 1) НОВОЕ: Ищем в PLATE_LOAD_DETAILS (приоритет - самая частая нагрузка для этих размеров)
    matching_loads = []
    for (L, W, load), qty in PLATE_LOAD_DETAILS.items():
        if abs(L - key_base[0]) <= 0.05 and abs(W - key_base[1]) <= 0.05:
            matching_loads.append((load, qty))
    
    if matching_loads:
        # Возвращаем нагрузку с максимальным количеством плит
        most_common_load = max(matching_loads, key=lambda x: x[1])[0]
        return most_common_load

    # 2) Точное совпадение в PLATE_LOAD_MAP (обратная совместимость)
    code = PLATE_LOAD_MAP.get(key_base)
    if isinstance(code, int) and code > 0:
        return code

    # 3) Поиск по близким значениям (на случай округлений и нормализации ширины)
    LENGTH_TOL = 0.05      # 5 см по длине
    WIDTH_TOL_M = 0.05     # 0.05 м = 50 мм по ширине

    for (L, W), c in PLATE_LOAD_MAP.items():
        if not isinstance(c, int) or c <= 0:
            continue
        if abs(L - key_base[0]) <= LENGTH_TOL and abs(W - key_base[1]) <= WIDTH_TOL_M:
            return c

    # 4) Fallback: старая логика по ширине
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
    key = (round(float(length_m), 2), target_list_name)
    return PLATE_EXACT_WIDTHS.get(key, default_width)


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





