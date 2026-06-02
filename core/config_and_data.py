#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конфигурация и данные проекта:
- Константы (размеры дорожки, цены резов) и ``get_config()`` → :class:`core.config.app_settings.AppConfig`
- Глобальные списки плит (legacy-имена PLATES_* / PLATE_* через PEP 562 ``__getattr__`` → рантайм)
- Парсинг текста пользователя

Мутабельные имена заказа не назначаются через ``cfg.NAME = ...``; только чтение и in-place мутация объектов.
"""
import math
import re
import logging
from typing import Any, Dict, List, Optional, Tuple

from .debug_paths import get_debug_log_path
from .project_paths import (
    BASE_DIR,
    CUTS_DOCX_PATH,
    PRICE_DB_PATH,
    PRICE_XLSX_PATH,
)
from .project_settings import WEIGHT_SOURCE
from .plate_runtime_state import (
    MUTABLE_ATTR_MAP,
    MUTABLE_LEGACY_NAMES,
    bind_plate_mutable_runtime,
    get_plate_mutable_runtime,
    new_plate_mutable_runtime_empty,
    plate_mutable_runtime_scope,
    reset_plate_mutable_runtime,
)
from .config.app_settings import AppConfig, get_config
from .config.constants import (
    LONG_CUT_PRICE_PER_M,
    MIN_BILLABLE_TRIM_MM,
    TRACK_LENGTH_M,
    TRACK_WIDTH_M,
    TRANSVERSE_CUT_PRICE,
    WEIGHT_KG_PER_DM2,
    length_dm_to_m,
    normalize_dimension,
    parse_pb_width_to_m,
)
from .domain.plate_order import (
    PlateOrder,
    get_current_plate_order,
    normalize_load_code,
)
from .parsing.plate_lists import (
    LineContributionKey,
    _clear_all_plate_lists,
    _recompute_totals_from_lists,
    get_last_parse_diagnostics,
    set_plate_lists_from_text,
)

# Настройка логирования
logger = logging.getLogger(__name__)
_DEBUG_LOG_8E9428 = get_debug_log_path("debug-8e9428.log")
_DEBUG_LOG_B59370 = get_debug_log_path("debug-b59370.log")

# Пути и env: см. core.project_paths / core.project_settings (реэкспорт для обратной совместимости).

# ==================== МУТАБЕЛЬНОЕ СОСТОЯНИЕ ЗАКАЗА (PLATE-CTX-001) ====================
# Списки PLATES_* / PLATE_LOAD_DETAILS и связанные структуры живут в
# core.plate_runtime_state (thread-local + опциональный ContextVar для asyncio).
# Доступ через атрибуты этого модуля — см. ``__getattr__`` в конце файла (PEP 562).


# ==================== КЛАСС ЗАКАЗА (PlateOrder) ====================
# PlateOrder, get_current_plate_order, normalize_load_code реализованы в core.domain.plate_order
# и реэкспортируются выше для from core.config_and_data import ...

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
                _log_path = _DEBUG_LOG_8E9428
                with open(_log_path, 'a', encoding='utf-8') as _f:
                    _f.write(__import__('json').dumps({"sessionId": "8e9428", "hypothesisId": "H_length_001", "location": "config_and_data:make_plate_name", "message": "57/57,1: length_m -> length_str", "data": {"length_m": length_m, "length_dm_raw": length_dm_raw, "length_dm_val": length_dm_val, "length_str": length_str, "branch_001_used": branch_001}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
            except Exception:
                pass
        # #endregion
        # #region agent log
        try:
            _log_path = _DEBUG_LOG_B59370
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

    # Ищем в plate_load_details (самая частая нагрузка для этих размеров)
    matching_loads = []
    rt = get_plate_mutable_runtime()
    for key, qty in rt.plate_load_details.items():
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
    return get_plate_mutable_runtime().plate_exact_widths.get(key, default_width)


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
    meta = get_plate_mutable_runtime().plate_metadata
    meta.clear()
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
        meta.setdefault((length, width_mm), []).append(entry)


def consume_plate_metadata(length_m: float, width_mm: int, qty: int) -> List[Dict[str, Any]]:
    """Возвращает и удаляет из буфера метаданные, соответствующие плитам."""
    meta = get_plate_mutable_runtime().plate_metadata
    key = (round(float(length_m), 2), int(width_mm))
    bucket = meta.get(key, [])
    taken = bucket[:qty]
    meta[key] = bucket[qty:]
    return taken


def clear_plate_metadata() -> None:
    """Полностью очищает буфер метаданных плит."""
    get_plate_mutable_runtime().plate_metadata.clear()


def __getattr__(name: str) -> Any:
    if name in MUTABLE_LEGACY_NAMES:
        return getattr(get_plate_mutable_runtime(), MUTABLE_ATTR_MAP[name])
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def __dir__() -> list[str]:
    names = set(globals().keys())
    names.update(MUTABLE_LEGACY_NAMES)
    return sorted(names)
