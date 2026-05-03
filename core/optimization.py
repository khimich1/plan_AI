#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль оптимизации раскроя плит:
- Оптимизация резов по ширине (PuLP)
- Оптимизация с учётом длин
- Полная оптимизация с narrowing
- Каскадные продольные резы (вторичное использование остатков)

Терминология:
- ПРОДОЛЬНЫЙ РЕЗ: режет вдоль длины, уменьшает ШИРИНУ (1.2м → 0.32м + 0.88м)
- ПОПЕРЕЧНЫЙ РЕЗ: режет поперёк, уменьшает ДЛИНУ (6.0м → 3.0м + 3.0м)
"""
# Относительные импорты внутри core/
import os as _os
from pathlib import Path as _Path
from . import config_and_data as cfg
from .config_and_data import canonical_plate_key
from .price_db import get_price
from dataclasses import dataclass

_DEBUG_LOG_5b5324 = _Path(__file__).resolve().parent.parent / "debug-5b5324.log"
_DEBUG_RUNTIME_LOG_648532 = _Path(__file__).resolve().parent.parent / "debug-648532.log"
_DEBUG_RUNTIME_SESSION_ID_648532 = "648532"


# ==================== ХЕЛПЕРЫ ОПТИМИЗАЦИИ ====================

def _canonical_length(length) -> float:
    """
    Каноническая длина (метры, round до 2 знаков).
    Используется как единый формат сравнения длин в ILP-модели,
    чтобы убрать float-tolerance с риском дрейфа при round-trip через JSON.
    """
    try:
        return round(float(length), 2)
    except (TypeError, ValueError):
        return 0.0


def _opt_debug_enabled() -> bool:
    """
    Включены ли подробные debug-логи оптимизатора.
    По умолчанию выключены: дебаг-регионы пишут в файлы только при OPT_DEBUG_LOG=1.
    Это нужно для честных замеров и чтобы prod не засорял диск.
    """
    return _os.environ.get("OPT_DEBUG_LOG", "").strip() in ("1", "true", "True", "yes", "on")


class _DbgNullFile:
    """
    No-op file handle: используется как заглушка для debug-логов,
    когда OPT_DEBUG_LOG выключен. Поддерживает и контекст-менеджер,
    и прямой `.write(...)` без `with`.
    """

    def write(self, *_args, **_kwargs):
        return 0

    def __enter__(self):
        return self

    def __exit__(self, *_args, **_kwargs):
        return False


_DBG_NULL_FILE = _DbgNullFile()


def _dbg_open_append(path):
    """
    Append-handle для debug-логов оптимизатора.
    При OPT_DEBUG_LOG=0 возвращает no-op handle, поэтому никакая запись не идёт.
    Никогда не бросает исключений: при ошибке открытия — тоже no-op.
    """
    if not _opt_debug_enabled():
        return _DBG_NULL_FILE
    try:
        return open(path, 'a', encoding='utf-8')
    except Exception:
        return _DBG_NULL_FILE


def verify_coverage(
    demand_2d: dict,
    primary_cuts: list,
    secondary_cuts: list,
) -> dict:
    """
    Сверяет фактическое покрытие спроса по primary_cuts + secondary_cuts.

    Returns:
        {
            "demand_total": int,
            "covered_total": int,
            "missing": {(L, W, lc): int},     # дефицит по ключам
            "surplus": {(L, W, lc): int},     # перепроизводство по ключам
            "ok": bool                          # True, если нет дефицита
        }
    Все ключи приведены через canonical_plate_key, так что 800/8 и
    дробные длины не дают ложного несоответствия.
    """
    from collections import Counter as _Counter

    demand_norm: _Counter = _Counter()
    for key, qty in (demand_2d or {}).items():
        if isinstance(key, tuple) and len(key) == 3:
            demand_norm[canonical_plate_key(*key)] += int(qty)

    coverage: _Counter = _Counter()
    for cut in primary_cuts or []:
        ak = cut.get("assignment_key")
        if ak and isinstance(ak, tuple) and len(ak) == 3:
            coverage[canonical_plate_key(*ak)] += 1
    for cut in secondary_cuts or []:
        tk = cut.get("target_order_key")
        if tk and isinstance(tk, tuple) and len(tk) == 3:
            coverage[canonical_plate_key(*tk)] += 1

    missing: dict = {}
    surplus: dict = {}
    for key, need in demand_norm.items():
        have = coverage.get(key, 0)
        if have < need:
            missing[key] = need - have
        elif have > need:
            surplus[key] = have - need

    return {
        "demand_total": int(sum(demand_norm.values())),
        "covered_total": int(sum(coverage.values())),
        "missing": missing,
        "surplus": surplus,
        "ok": not missing,
    }


def _debug_runtime_write_648532(
    run_id: str,
    hypothesis_id: str,
    location: str,
    message: str,
    data: dict,
) -> None:
    if not _opt_debug_enabled():
        return
    try:
        import json as _json
        import time as _time

        with open(_DEBUG_RUNTIME_LOG_648532, "a", encoding="utf-8") as _f:
            _f.write(
                _json.dumps(
                    {
                        "sessionId": _DEBUG_RUNTIME_SESSION_ID_648532,
                        "runId": run_id,
                        "hypothesisId": hypothesis_id,
                        "location": location,
                        "message": message,
                        "data": data,
                        "timestamp": int(_time.time() * 1000),
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
    except Exception:
        pass


# ==================== КОНФИГУРАЦИЯ ОПТИМИЗАЦИИ ====================

# Ширина пропила (в мм) - НЕ используется в расчётах
# Пропил косвенно учитывается в таблице NARROWING_TABLE через значения narrowing
# Формально для совместимости оставляем константу, но не применяем в расчётах
KERF_WIDTH_MM = 0

@dataclass
class OptimizationConfig:
    """
    Конфигурация параметров оптимизации.
    Позволяет экспериментировать с разными коэффициентами штрафов и бонусов.
    """
    # Коэффициент штрафа за неиспользованные остатки
    # OLD: 0.5 (50% стоимости остатка)
    # NEW: 0.15 (15% стоимости остатка)
    unused_rest_penalty_coeff: float = 0.15
    
    # Бонус за использование вторичных резов (отрицательное значение = бонус)
    # OLD: -500 (экономический стимул использовать остатки)
    # NEW: 0 (нет бонуса, остатки используются только если это выгодно)
    secondary_reuse_bonus: float = 0.0


# Дефолтная конфигурация (NEW поведение)
DEFAULT_CONFIG = OptimizationConfig(
    unused_rest_penalty_coeff=0.15,
    secondary_reuse_bonus=0.0
)

# Старая конфигурация (OLD поведение, для экспериментов)
OLD_CONFIG = OptimizationConfig(
    unused_rest_penalty_coeff=0.5,
    secondary_reuse_bonus=-500.0
)

# ==================== ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ ОПТИМИЗАЦИИ ====================

OPT_WIDTH_PRIORITY = []  # приоритет ширин: ['0_32','0_46','0_70','0_72','0_86']
OPT_PLAN = {}  # результат полной оптимизации: как закрывать спрос
OPT_CASCADING_PLAN = {}  # результат каскадной оптимизации с вторичными резами
OPT_CASCADING_PLAN_BY_LOAD = {}  # результат каскадной оптимизации, СГРУППИРОВАННЫЙ ПО НАГРУЗКЕ
LOAD_TO_REINFORCEMENT_MAP = {}  # маппинг: load_code → [reinforcement_keys] для поиска плана по нагрузке


# ==================== ЛЕГАСИ-АДАПТЕРЫ ====================
from collections import Counter, defaultdict


def _group_plate_lengths(plates: list[float]) -> dict[float, int]:
    """Группируем длины (в метрах) -> количество."""
    return Counter(round(float(length), 2) for length in (plates or []))


def _append_actions(actions: list, width_mm: int, lengths: dict, long_cuts: int, src_type: str):
    """Добавляем aggregated записи в OPT_PLAN['actions']."""
    rest_mm = max(0, 1200 - width_mm)
    for length_m, qty in lengths.items():
        actions.append((
            src_type,
            width_mm,
            rest_mm if src_type != 'solid' else 0,
            length_m,
            qty,
            long_cuts,
            0  # поперечные резы в legacy-плане не моделируем
        ))


def apply_width_optimization() -> dict:
    """
    Упрощённый наследуемый оптимизатор ширин.
    Наполняет OPT_WIDTH_PRIORITY и OPT_PLAN['actions'] данными из cfg.PLATES_*.
    """
    priority = []
    actions = []

    priority_groups = [
        ('0_32', 320, cfg.PLATES_0_32),
        ('0_46', 460, cfg.PLATES_0_46),
        ('0_70', 700, cfg.PLATES_0_70),
        ('0_72', 720, cfg.PLATES_0_72),
        ('0_86', 860, cfg.PLATES_0_86),
        ('0_74', 740, cfg.PLATES_0_74),
        ('0_88', 880, cfg.PLATES_0_88),
        ('0_48', 480, cfg.PLATES_0_48),
        ('0_50', 500, cfg.PLATES_0_50),
        ('0_34', 340, cfg.PLATES_0_34),
    ]

    for code, width_mm, plate_list in priority_groups:
        if not plate_list:
            continue
        priority.append(code)
        lengths = _group_plate_lengths(plate_list)
        _append_actions(actions, width_mm, lengths, long_cuts=1, src_type='split')

    solid_groups = [
        (1200, cfg.PLATES_1_2),
    ]
    split_groups = [
        (1080, cfg.PLATES_1_08),
        (1000, cfg.PLATES_1_0),
    ]

    for width_mm, plate_list in solid_groups:
        if not plate_list:
            continue
        lengths = _group_plate_lengths(plate_list)
        _append_actions(actions, width_mm, lengths, long_cuts=0, src_type='solid')

    for width_mm, plate_list in split_groups:
        if not plate_list:
            continue
        lengths = _group_plate_lengths(plate_list)
        _append_actions(actions, width_mm, lengths, long_cuts=1, src_type='split')

    OPT_WIDTH_PRIORITY.clear()
    OPT_WIDTH_PRIORITY.extend(priority)

    OPT_PLAN.clear()
    OPT_PLAN.update({
        'orders': {width: len(lst) for (_, width, lst) in priority_groups if lst},
        'actions': actions
    })

    return OPT_PLAN


def optimize_cuts_pulp(orders: dict | None = None) -> dict:
    """
    Legacy-обёртка над новой каскадной оптимизацией (для совместимости с визуализатором).
    """
    if orders is None:
        orders = {}
        for width_mm, plates in [
            (320, cfg.PLATES_0_32), (460, cfg.PLATES_0_46), (700, cfg.PLATES_0_70),
            (720, cfg.PLATES_0_72), (860, cfg.PLATES_0_86), (880, cfg.PLATES_0_88),
            (740, cfg.PLATES_0_74), (480, cfg.PLATES_0_48), (500, cfg.PLATES_0_50),
            (340, cfg.PLATES_0_34)
        ]:
            if plates:
                orders[width_mm] = len(plates)

    result = optimize_with_cascading_longitudinal_cuts(orders=orders) if orders else {}
    if result:
        OPT_PLAN.clear()
        OPT_PLAN.update(result)
    return result


# ==================== СОВРЕМЕННЫЕ ФУНКЦИИ ОПТИМИЗАЦИИ ====================

def _get_next_order_info(order_info_list: dict, key: tuple) -> dict:
    """
    Возвращает информацию о следующем КП с qty_remaining > 0 и уменьшает счётчик.
    
    Простыми словами:
    - Ищет в списке записей для данного (length, width, load_code) первую запись, 
      у которой ещё есть неназначенные плиты (qty_remaining > 0)
    - Уменьшает счётчик qty_remaining на 1
    - Возвращает копию информации о КП
    
    Args:
        order_info_list: словарь {(length, width, load_code): [список записей КП]}
        key: кортеж (length, width, load_code)
    
    Returns:
        dict: информация о КП (kp_id, customer, kp_date, plate_name, load_code) или пустой словарь
    """
    # #region agent log (session 5b5324) _get_next_order_info entry
    try:
        with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
            _f.write(__import__('json').dumps({"sessionId": "5b5324", "hypothesisId": "H_get_next", "location": "optimization:_get_next_order_info:entry", "message": "key requested", "data": {"key": list(key) if isinstance(key, tuple) else key}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    entries = order_info_list.get(key, [])
    for entry in entries:
        if entry.get('qty_remaining', 0) > 0:
            entry['qty_remaining'] -= 1
            out = {
                'kp_id': entry.get('kp_id'),
                'customer': entry.get('customer'),
                'kp_date': entry.get('kp_date'),
                'plate_name': entry.get('plate_name'),
                'load_code': entry.get('load_code'),
                'reinforcement': entry.get('reinforcement'),
                'identity_match_type': 'exact'
            }
            # #region agent log (session 5b5324) _get_next_order_info return exact
            try:
                with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
                    _f.write(__import__('json').dumps({"sessionId": "5b5324", "hypothesisId": "H_get_next", "location": "optimization:_get_next_order_info:return", "message": "exact match", "data": {"match_type": "exact", "kp_id": out.get("kp_id"), "plate_name": (out.get("plate_name") or "")[:60]}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
            except Exception:
                pass
            # #endregion
            return out
    # Fallback по (length, width) без load_code — ищем любой ключ с теми же длиной и шириной
    if len(key) == 3:
        length, width, load_code = key
        for candidate_key, candidate_entries in order_info_list.items():
            if len(candidate_key) >= 2 and candidate_key[0] == length and candidate_key[1] == width:
                for entry in candidate_entries:
                    if entry.get('qty_remaining', 0) > 0:
                        entry['qty_remaining'] -= 1
                        try:
                            _req_lc = key[2] if len(key) >= 3 else None
                            _found_lc = candidate_key[2] if len(candidate_key) >= 3 else None
                            with _dbg_open_append(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log") as _f:
                                _f.write(__import__('json').dumps({
                                    "hypothesisId": "H2_fallback",
                                    "location": "optimization.py:_get_next_order_info",
                                    "message": "fallback used (length, width)",
                                    "data": {
                                        "requested_key": list(key),
                                        "found_key": list(candidate_key),
                                        "fallback_reason": "load_code_mismatch",
                                        "requested_load_code": _req_lc,
                                        "found_load_code": _found_lc,
                                        "kp_id": entry.get('kp_id'),
                                        "plate_name": (entry.get('plate_name') or '')[:50]
                                    },
                                    "timestamp": __import__('time').time()
                                }, ensure_ascii=False) + '\n')
                        except Exception:
                            pass
                        out_fb = {
                            'kp_id': entry.get('kp_id'),
                            'customer': entry.get('customer'),
                            'kp_date': entry.get('kp_date'),
                            'plate_name': entry.get('plate_name'),
                            'load_code': entry.get('load_code'),
                            'reinforcement': entry.get('reinforcement'),
                            'identity_match_type': 'fallback_same_length_width'
                        }
                        # #region agent log (session 5b5324) fallback_same_length_width
                        try:
                            with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
                                _f.write(__import__('json').dumps({"sessionId": "5b5324", "hypothesisId": "H_get_next", "location": "optimization:_get_next_order_info:return", "message": "fallback_same_length_width", "data": {"requested_key": list(key), "found_key": list(candidate_key), "kp_id": out_fb.get("kp_id"), "plate_name": (out_fb.get("plate_name") or "")[:60]}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
                        except Exception:
                            pass
                        # #endregion
                        return out_fb
        # Fallback по «соседней» длине (±0.02 м), та же ширина и load_code (61,2↔61,1; 59,8↔59,9)
        # Иначе при конкурирующих длинах решатель даёт общий объём, список по точной длине кончается —
        # плиты получают kp_id из opt (первый КП), а в БД они в другом КП и не списываются.
        LEN_TOL = 0.02
        for candidate_key, candidate_entries in order_info_list.items():
            if len(candidate_key) < 3:
                continue
            c_len, c_width, c_lc = candidate_key[0], candidate_key[1], candidate_key[2]
            if abs(c_len - length) <= LEN_TOL and c_width == width and c_lc == load_code:
                for entry in candidate_entries:
                    if entry.get('qty_remaining', 0) > 0:
                        entry['qty_remaining'] -= 1
                        out_n = {
                            'kp_id': entry.get('kp_id'),
                            'customer': entry.get('customer'),
                            'kp_date': entry.get('kp_date'),
                            'plate_name': entry.get('plate_name'),
                            'load_code': entry.get('load_code'),
                            'reinforcement': entry.get('reinforcement'),
                            'identity_match_type': 'fallback_neighbor_length'
                        }
                        # #region agent log (session 5b5324) fallback_neighbor_length
                        try:
                            with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
                                _f.write(__import__('json').dumps({"sessionId": "5b5324", "hypothesisId": "H_get_next", "location": "optimization:_get_next_order_info:return", "message": "fallback_neighbor_length", "data": {"requested_key": list(key), "found_key": list(candidate_key), "kp_id": out_n.get("kp_id"), "plate_name": (out_n.get("plate_name") or "")[:60]}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
                        except Exception:
                            pass
                        # #endregion
                        return out_n
    # #region agent log (session 5b5324) _get_next_order_info return empty
    try:
        with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
            _f.write(__import__('json').dumps({"sessionId": "5b5324", "hypothesisId": "H_get_next", "location": "optimization:_get_next_order_info:return", "message": "empty", "data": {"key": list(key) if isinstance(key, tuple) else key}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    return {}


def _build_proportional_slot_lists(
    orders_2d: list,
    demand_2d: dict,
) -> tuple:
    """
    Строит пропорциональные слоты атрибуции по ключу (length, width, load_code).
    Возвращает (slot_lists, slot_cursors).
    slot_lists[key] — список из demand_2d[key] атрибуций, пропорционально qty заказов
    (floor + остаток по убыванию qty). Курсоры инициализированы в 0.
    """
    groups: dict = {}
    for order in orders_2d:
        key = (order['length'], order['width'], order.get('load_code', 800))
        groups.setdefault(key, []).append(order)

    slot_lists: dict = {}
    for key, need in demand_2d.items():
        entries = groups.get(key, [])
        total_qty = sum(o.get('qty', 1) for o in entries)
        if total_qty == 0:
            slot_lists[key] = []
            continue
        shares = [int(need * o.get('qty', 1) / total_qty) for o in entries]
        remainder = need - sum(shares)
        for idx in sorted(range(len(entries)), key=lambda i: -entries[i].get('qty', 1)):
            if remainder <= 0:
                break
            shares[idx] += 1
            remainder -= 1
        slots = []
        for entry, share in zip(entries, shares):
            info = {
                k: entry.get(k)
                for k in ('kp_id', 'customer', 'kp_date', 'plate_name', 'load_code', 'reinforcement')
            }
            slots.extend([info] * share)
        slot_lists[key] = slots

    cursors = {key: 0 for key in slot_lists}
    return slot_lists, cursors


def _next_slot_info(
    slot_lists: dict,
    slot_cursors: dict,
    key: tuple,
) -> dict:
    """
    Возвращает следующую атрибуцию по ключу из предрасчитанных слотов и сдвигает курсор.
    При исчерпании слотов возвращает пустой dict, чтобы не дублировать identity.
    """
    slots = slot_lists.get(key, [])
    idx = slot_cursors.get(key, 0)
    if not slots or idx >= len(slots):
        return {}
    entry = dict(slots[idx])
    entry['identity_match_type'] = 'slot_proportional'
    slot_cursors[key] = idx + 1
    return entry


def _peek_order_info(order_info_list: dict, key: tuple) -> dict:
    """
    Возвращает информацию о первом КП с qty_remaining > 0 БЕЗ уменьшения счётчика.
    
    Используется для получения информации при создании primary_options,
    когда ещё не известно, будет ли опция использована.
    
    Args:
        order_info_list: словарь {(length, width, load_code): [список записей КП]}
        key: кортеж (length, width, load_code)
    
    Returns:
        dict: информация о КП (включая load_code) или пустой словарь
    """
    entries = order_info_list.get(key, [])
    for entry in entries:
        if entry.get('qty_remaining', 0) > 0:
            return {
                'kp_id': entry.get('kp_id'),
                'customer': entry.get('customer'),
                'kp_date': entry.get('kp_date'),
                'plate_name': entry.get('plate_name'),
                'load_code': entry.get('load_code'),
                'reinforcement': entry.get('reinforcement')
            }
    # Fallback по (length, width) без load_code
    if len(key) == 3:
        length, width, _ = key
        for candidate_key, candidate_entries in order_info_list.items():
            if len(candidate_key) >= 2 and candidate_key[0] == length and candidate_key[1] == width:
                for entry in candidate_entries:
                    if entry.get('qty_remaining', 0) > 0:
                        return {
                            'kp_id': entry.get('kp_id'),
                            'customer': entry.get('customer'),
                            'kp_date': entry.get('kp_date'),
                            'plate_name': entry.get('plate_name'),
                            'load_code': entry.get('load_code'),
                            'reinforcement': entry.get('reinforcement')
                        }
    return {}


def _optimize_2d_with_lengths(orders_2d: list, plate_width: int = 1200,
                               min_useful_width: int = 200,
                               opt_config: OptimizationConfig = None) -> dict:
    """
    ПРИВАТНАЯ функция: Полная 2D оптимизация с длинами в ILP модели.
    Минимизирует СТОИМОСТЬ (не просто количество плит!) с учётом ОБЕИХ размерностей (длина + ширина).
    
    УЛУЧШЕНИЯ (версия 2.0):
    ✅ Учёт реальных цен плит из базы данных (get_price)
    ✅ Учёт стоимости продольных и поперечных резов
    ✅ Фильтрация бесполезных вариантов (скорость ↑ в 2-3 раза)
    ✅ Настраиваемые штрафы и бонусы через OptimizationConfig
    ✅ Сохранение kp_id и customer для каждой плиты (для диаграммы Ганта)
    
    Args:
        orders_2d: [{'length': 5.6, 'width': 320, 'qty': 11, 'kp_id': 1, 'customer': 'Роман'}, ...] — спрос по (длина, ширина)
        plate_width: ширина исходной плиты в мм (1200)
        min_useful_width: минимальная полезная ширина остатка
        opt_config: конфигурация параметров оптимизации (штрафы, бонусы)
    
    Returns:
        {
            'primary_cuts': [{'width', 'rest', 'qty', 'lengths': [5.6, ...]}, ...],
            'secondary_cuts': [{'source', 'cuts', 'qty', 'pieces', 'lengths': [...], 'type'}, ...],
            'total_plates': int,
            'plate_assignments': [{'length', 'width', 'source', 'kp_id', 'customer', ...}, ...]
        }
    """
    # Используем дефолтную конфигурацию, если не передана
    if opt_config is None:
        opt_config = DEFAULT_CONFIG
    try:
        from pulp import LpProblem, LpMinimize, LpVariable, LpInteger, lpSum, value, PULP_CBC_CMD, LpStatus
    except ImportError:
        print('[OPT_2D] PuLP не установлен.')
        return {}
    
    if not orders_2d:
        return {}
    
    print(f"\n[OPT_2D] === ПОЛНАЯ 2D ОПТИМИЗАЦИЯ ===")
    print(f"[OPT_2D] Заказ:")
    for order in orders_2d:
        print(f"  {order['qty']}x {order['length']}м x {order['width']}мм")
    
    # 1. ПОДГОТОВКА: Группируем спрос по (length, width, load_code)
    # Используем canonical_plate_key — единый формат ключей во всём проекте.
    demand_2d = {}  # {(length_round2, width_int_mm, load_code_norm): qty}
    for order in orders_2d:
        key = canonical_plate_key(order['length'], order['width'], order.get('load_code', 800))
        demand_2d[key] = demand_2d.get(key, 0) + order['qty']
    # #region agent log: demand keys 59/10 (H3,H4)
    _demand_59_10 = [(list(k), v) for k, v in demand_2d.items() if abs(k[0] - 5.99) < 0.02 and (k[2] == 10 or abs(float(k[2]) - 10) < 0.01)]
    if _demand_59_10:
        try:
            _dbg_open_append(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log").write(__import__("json").dumps({"hypothesisId": "H_59_10_demand", "location": "optimization.py:demand_2d_built", "message": "demand_2d: ключи 5.99м 10п (length, width, load_code)", "data": {"keys": _demand_59_10}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
    # #endregion
    # 1.5 НОВОЕ: Создаём маппинг (length, width, load_code) -> СПИСОК информации о КП
    # Используем canonical_plate_key — тот же формат, что в demand_2d.
    order_info_list = {}  # {(length_round2, width_int_mm, load_code_norm): [список записей]}
    for order in orders_2d:
        load_code = cfg.normalize_load_code(order.get('load_code', 800))
        key = canonical_plate_key(order['length'], order['width'], load_code)
        if key not in order_info_list:
            order_info_list[key] = []
        # Добавляем запись для КАЖДОГО заказа с qty_remaining
        order_info_list[key].append({
            'kp_id': order.get('kp_id'),
            'customer': order.get('customer', 'неизвестно'),
            'kp_date': order.get('kp_date', 'неизвестно'),
            'plate_name': order.get('plate_name', ''),
            'load_code': load_code,
            'reinforcement': order.get('reinforcement', 0),
            'qty_remaining': order.get('qty', 1)  # Сколько плит этого КП осталось назначить
        })

    slot_lists, slot_cursors = _build_proportional_slot_lists(orders_2d, demand_2d)
    # #region agent log (session 5b5324) после построения slot_lists
    try:
        _slot_summary = [(list(k), len(slots)) for k, slots in list(slot_lists.items())[:20]]
        _sample_slot = []
        for k, slots in list(slot_lists.items())[:3]:
            if slots:
                _sample_slot.append({"key": list(k), "first_identity": [slots[0].get("kp_id"), (slots[0].get("plate_name") or "")[:50]]})
        with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
            _f.write(__import__('json').dumps({"sessionId": "5b5324", "hypothesisId": "H_slots", "location": "optimization:after_build_slot_lists", "message": "slot_lists summary", "data": {"slot_summary": _slot_summary, "sample_slot_identity": _sample_slot}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # #region agent log
    try:
        _slot_mismatch = []
        _slot_empty = []
        for _k, _need in demand_2d.items():
            _slots_len = len(slot_lists.get(_k, []))
            if _slots_len == 0 and _need > 0:
                _slot_empty.append({"key": list(_k), "need": int(_need)})
            if _slots_len < _need:
                _slot_mismatch.append({"key": list(_k), "need": int(_need), "slots": int(_slots_len)})
        _debug_runtime_write_648532(
            "run1",
            "H1_slot_key_alignment",
            "optimization:after_build_slot_lists",
            "Demand keys versus proportional slot capacity",
            {
                "demand_total": int(sum(demand_2d.values())),
                "slot_total": int(sum(len(v) for v in slot_lists.values())),
                "demand_keys": int(len(demand_2d)),
                "slot_keys": int(len(slot_lists)),
                "empty_slot_keys": _slot_empty[:50],
                "short_slot_keys": _slot_mismatch[:50],
            },
        )
    except Exception:
        pass
    # #endregion

    tolerance_width = 20     # ±20мм по ширине (генерация опций, напр. поперечный рез остаток→цель)
    demand_tolerance_width = 10  # ±10мм при сопоставлении спроса с источником (по письму: допуск реза при работе)
    # Длины сравниваем через канонический ключ (round до 2 знаков), а не через float-tolerance.
    # tolerance_length оставлен для совместимости со старым блоком sources_to_demands,
    # который будет удалён в этапе 2.7 ассайнмент-рефакторинга.
    tolerance_length = 0

    # 2. ГЕНЕРАЦИЯ ОПЦИЙ ПЕРВИЧНЫХ РЕЗОВ (с длинами!)
    primary_options = []
    option_id = 0
    
    # Таблица сужений (из таблицы допустимых резов)
    # Формат: (исходная_ширина_остатка, целевая_ширина, отход)
    # ВАЖНО: Значения остатков рассчитаны БЕЗ явного учёта пропила в коде
    # Пропил косвенно учтён через значения narrowing (разница между source_rest и target_w)
    NARROWING_TABLE = [
        (480, 460, 20),   # Остаток 480мм → 460мм (из реза 720+480)
        (500, 460, 40),   # Остаток 500мм → 460мм (из реза 700+500)
        (495, 460, 35),   # Остаток 495мм → 460мм (из реза 720+495 или 700+495)
        (740, 720, 20),   # Остаток 740мм → 720мм (из реза 460+740)
        (690, 660, 30),   # Остаток 690мм → 660мм (из реза 460+690)
        (890, 860, 30),   # Остаток 890мм → 860мм (из реза 320+890)
        (495, 480, 15),   # Остаток 495мм → 480мм
    ]
    
    # Создаём обратный индекс: для каждой целевой ширины -> список (основная_ширина, ширина_остатка, отход)
    target_to_sources = {}  # {460: [(720, 480, 20), (700, 500, 40), ...]}
    for source_rest, target_w, waste in NARROWING_TABLE:
        if target_w not in target_to_sources:
            target_to_sources[target_w] = []
        # Ищем, из какого первичного реза получается source_rest
        main_w = plate_width - source_rest
        if 200 <= main_w <= 1000:  # Разумный диапазон для основной части
            target_to_sources[target_w].append((main_w, source_rest, waste))
    
    print(f"[OPT_2D] Таблица narrowing создана: {len(NARROWING_TABLE)} правил")
    for target_w, sources in target_to_sources.items():
        print(f"  {target_w}мм можно получить через: {sources}")
    
    solid_widths = sorted(set([plate_width, 1080]))

    for (length, width, load_code), qty in demand_2d.items():
        # Получаем информацию о КП для этой плиты (без уменьшения счётчика)
        # ИСПРАВЛЕНИЕ: Ключ теперь включает load_code
        order_info = _peek_order_info(order_info_list, (length, width, load_code))
        
        # Вариант 1: Плита БЕЗ реза (ширины из списка solid_widths)
        # Эти ширины НЕ РЕЖУТСЯ и используются как есть
        if width in solid_widths:
            primary_options.append({
                'id': option_id,
                'length': length,
                'main': width,
                'rest': 0,
                'type': 'solid',  # Без резов
                'load_code': order_info.get('load_code', 800),  # ИСПРАВЛЕНИЕ: добавляем load_code
                'kp_id': order_info.get('kp_id'),
                'customer': order_info.get('customer'),
                'kp_date': order_info.get('kp_date'),
                'plate_name': order_info.get('plate_name')
            })
            option_id += 1

        # Вариант 2: Плита С ПРЯМЫМ резом (ширина < исходной плиты)
        elif width < plate_width:
            # Пропил косвенно учтён в таблице NARROWING_TABLE
            rest = plate_width - width
            # Создаём вариант для ЛЮБОЙ ширины
            # Если rest < min_useful_width, остаток просто пойдёт в отход
            primary_options.append({
                'id': option_id,
                'length': length,
                'main': width,
                'rest': rest,
                'type': 'direct',  # Прямой рез
                'load_code': order_info.get('load_code', 800),  # ИСПРАВЛЕНИЕ: добавляем load_code
                'kp_id': order_info.get('kp_id'),
                'customer': order_info.get('customer'),
                'kp_date': order_info.get('kp_date'),
                'plate_name': order_info.get('plate_name')
            })
            option_id += 1
            
            # НОВОЕ! Вариант 3: Плита через НЕПРЯМОЙ рез (с narrowing остатка)
            # Ищем, можно ли получить эту ширину через сужение остатка от ДРУГОГО реза
            if width in target_to_sources:
                for main_w, rest_w, waste in target_to_sources[width]:
                    # Создаём первичный рез main_w + rest_w
                    # Остаток rest_w потом автоматически сузится до width
                    if main_w != width and rest_w >= min_useful_width:  # Не дублируем прямой рез
                        primary_options.append({
                            'id': option_id,
                            'length': length,
                            'main': main_w,           # Например, 720мм (основная часть)
                            'rest': rest_w,           # Например, 480мм (остаток)
                            'type': 'indirect',       # Непрямой рез через narrowing
                            'target_width': width,    # Целевая ширина: 460мм (что нужно)
                            'narrowing_waste': waste, # Отход при сужении: 20мм
                            'load_code': order_info.get('load_code', 800),  # ИСПРАВЛЕНИЕ: добавляем load_code
                            'kp_id': order_info.get('kp_id'),
                            'customer': order_info.get('customer'),
                            'kp_date': order_info.get('kp_date'),
                            'plate_name': order_info.get('plate_name')
                        })
                        option_id += 1
    
    print(f"[OPT_2D] Опций первичных резов (до фильтрации): {len(primary_options)}")
    # #region agent log (2d5c43) Plan B: опции для 5.1/320 и 6/530 до фильтра
    try:
        _log = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-2d5c43.log"
        _opts_320 = [{"id": o['id'], "length": o['length'], "main": o['main'], "type": o.get('type'), "load_code": o.get('load_code')} for o in primary_options if o.get('main') == 320 or o.get('target_width') == 320]
        _opts_530 = [{"id": o['id'], "length": o['length'], "main": o['main'], "type": o.get('type'), "load_code": o.get('load_code')} for o in primary_options if o.get('main') == 530 or o.get('target_width') == 530]
        with _dbg_open_append(_log) as _f:
            _f.write(__import__('json').dumps({"sessionId": "2d5c43", "hypothesisId": "H_opt_gen", "location": "optimization:primary_options_after_build", "message": "options for 320 and 530 before filter", "data": {"opts_320": _opts_320, "opts_530": _opts_530, "solid_widths": solid_widths}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # 2.5 ФИЛЬТРАЦИЯ ПЕРВИЧНЫХ ОПЦИЙ (Улучшение 4: убираем заведомо невыгодные)
    filtered_primary = []
    for opt in primary_options:
        # Правило 1 УДАЛЕНО: теперь плиты с маленьким остатком (< min_useful_width) 
        # тоже создаются, а остаток просто идёт в отход
        
        # Правило 2: Пропускаем indirect, если есть direct с тем же результатом
        if opt.get('type') == 'indirect':
            target_w = opt.get('target_width')
            has_direct = any(
                o['type'] == 'direct' and 
                o['main'] == target_w and
                _canonical_length(o['length']) == _canonical_length(opt['length'])
                for o in primary_options
            )
            if has_direct:
                continue
        
        filtered_primary.append(opt)
    
    primary_options = filtered_primary
    print(f"[OPT_2D] После фильтрации осталось: {len(primary_options)} первичных опций")
    # #region agent log (2d5c43) Plan B: опции для 320/530 после фильтра
    try:
        _log = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-2d5c43.log"
        _opts_320 = [{"id": o['id'], "length": o['length'], "main": o['main'], "type": o.get('type')} for o in primary_options if o.get('main') == 320 or o.get('target_width') == 320]
        _opts_530 = [{"id": o['id'], "length": o['length'], "main": o['main'], "type": o.get('type')} for o in primary_options if o.get('main') == 530 or o.get('target_width') == 530]
        with _dbg_open_append(_log) as _f:
            _f.write(__import__('json').dumps({"sessionId": "2d5c43", "hypothesisId": "H_opt_filter", "location": "optimization:primary_options_after_filter", "message": "options for 320 and 530 after filter", "data": {"opts_320": _opts_320, "opts_530": _opts_530}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # 3. ГЕНЕРАЦИЯ ОПЦИЙ ВТОРИЧНЫХ РЕЗОВ (2D: длина + ширина!)
    secondary_options = []
    
    # Собираем все возможные остатки (length, rest_width)
    possible_rests = {}  # {(length, rest_width): [source_option_ids]}
    for opt in primary_options:
        key = (opt['length'], opt['rest'])
        if key not in possible_rests:
            possible_rests[key] = []
        possible_rests[key].append(opt['id'])
    
    sec_id = 0
    for (source_length, source_width), source_ids in possible_rests.items():
        # Пропускаем остатки нулевой ширины (плиты 1200мм без резов)
        if source_width < min_useful_width:
            continue
        
        # Для каждого остатка проверяем все целевые (length, width, load_code)
        # ИСПРАВЛЕНИЕ: Теперь ключ включает load_code
        for (target_length, target_width, target_load_code), qty in demand_2d.items():
            
            # Вариант A: Множественная резка по ширине (одинаковая длина)
            # ВАЖНО: Нельзя получить больше, чем есть! target_width <= source_width
            if _canonical_length(target_length) == _canonical_length(source_length) and target_width <= source_width:
                # РАНЬШЕ: брали только максимум кусков (pieces = source_width // target_width)
                # ТЕПЕРЬ: перебираем все варианты от 1 до max_pieces, чтобы можно было
                # получать 1, 2, ... плит из одного остатка (например, 0.88 → 1×0.32 с хвостом 0.56)
                max_pieces = source_width // target_width
                for pieces in range(1, max_pieces + 1):
                    waste = source_width - (pieces * target_width)
                    # Отбрасываем варианты с слишком большим отходом.
                    # Для случаев с ОДНОЙ плитой позволяем больше отхода (до 80%),
                    # чтобы не выкидывать схему 0.88 → 0.32 + 0.56.
                    max_waste_fraction = 0.8 if pieces == 1 else 0.5
                    if waste <= source_width * max_waste_fraction:
                        secondary_options.append({
                            'id': sec_id,
                            'source_length': source_length,
                            'source_rest': source_width,
                            'output_length': target_length,
                            'output_width': target_width,
                            'pieces': pieces,
                            'waste': waste,
                            'type': 'multiple',
                            'source_ids': source_ids,
                            'target_order_key': (target_length, target_width, target_load_code)  # ИСПРАВЛЕНИЕ: Добавляем load_code
                        })
                        sec_id += 1

            
            # Вариант A2: Комбинированная резка (множественная по ширине + поперечная по длине)
            # Это позволяет резать остаток 5.6м × 880мм → 2× 3.31м × 320мм
            # ВАЖНО: Нельзя получить больше, чем есть! target_width <= source_width
            if target_length < source_length - 0.1 and target_width <= source_width:  # Целевая длина КОРОЧЕ остатка
                pieces = source_width // target_width
                if pieces >= 1:
                    # Проверяем, что целевая длина влезает хотя бы раз
                    waste_width = source_width - (pieces * target_width)
                    waste_length = (source_length - target_length) * 1000  # в мм
                    
                    if waste_width < source_width * 0.5:
                        secondary_options.append({
                            'id': sec_id,
                            'source_length': source_length,
                            'source_rest': source_width,
                            'output_length': target_length,  # КОРОЧЕ остатка!
                            'output_width': target_width,
                            'pieces': pieces,  # Кусков по ширине
                            'waste': waste_width,
                            'length_waste': waste_length,
                            'type': 'multiple_transverse',  # Комбинированный тип
                            'source_ids': source_ids,
                            'target_order_key': (target_length, target_width, target_load_code)  # ИСПРАВЛЕНИЕ: Добавляем load_code
                        })
                        sec_id += 1
            
            # Вариант B: Сужение (narrowing)
            if (_canonical_length(target_length) == _canonical_length(source_length) and
                target_width < source_width <= target_width + 100):
                waste = source_width - target_width
                if waste <= 100:
                    secondary_options.append({
                        'id': sec_id,
                        'source_length': source_length,
                        'source_rest': source_width,
                        'output_length': target_length,
                        'output_width': target_width,
                        'pieces': 1,
                        'waste': waste,
                        'type': 'narrowing',
                        'source_ids': source_ids,
                        'target_order_key': (target_length, target_width, target_load_code)  # ИСПРАВЛЕНИЕ: Добавляем load_code
                    })
                    sec_id += 1
            
            # Вариант C: Поперечный рез (transverse cut)
            # ВАЖНО: Нельзя получить больше, чем есть! target_width <= source_width
            # (раньше проверяли abs(target_width - source_width) <= tolerance_width, 
            #  что позволяло target_width > source_width на 20мм — это баг!)
            if (target_length < source_length - 0.1 and
                target_width <= source_width and
                source_width - target_width <= tolerance_width):
                length_waste = (source_length - target_length) * 1000  # в мм
                secondary_options.append({
                    'id': sec_id,
                    'source_length': source_length,
                    'source_rest': source_width,
                    'output_length': target_length,
                    'output_width': target_width,
                    'pieces': 1,
                    'waste': 0,
                    'length_waste': length_waste,
                    'type': 'transverse',
                    'source_ids': source_ids,
                    'target_order_key': (target_length, target_width, target_load_code)  # ИСПРАВЛЕНИЕ: Добавляем load_code
                })
                sec_id += 1
    
    print(f"[OPT_2D] Опций вторичных резов (до фильтрации): {len(secondary_options)}")
    
    # 3.5 ФИЛЬТРАЦИЯ ВТОРИЧНЫХ ОПЦИЙ (Улучшение 4: убираем дубликаты и невыгодные)
    filtered_secondary = []
    seen_combinations = set()
    
    for opt in secondary_options:
        # Правило 3: Убираем дубликаты (одинаковые варианты).
        # ВАЖНО: учитываем также количество кусков (pieces), чтобы варианты
        # 0.88 → 1×0.32 и 0.88 → 2×0.32 не считались одинаковыми.
        key = (
            opt['source_length'], 
            opt['source_rest'], 
            opt['output_length'], 
            opt['output_width'], 
            opt['type'],
            opt.get('pieces', 1)
        )
        
        if key in seen_combinations:
            continue
        seen_combinations.add(key)
        
        # Правило 4: Пропускаем варианты с огромными отходами.
        # Для случаев с ОДНОЙ плитой (pieces == 1) позволяем до 80% площади отхода,
        # иначе — как раньше, 30%.
        waste_width = opt.get('waste', 0)
        waste_length = opt.get('length_waste', 0)
        
        source_area = opt['source_length'] * opt['source_rest']
        waste_area = (waste_width * opt['source_length']) + (waste_length * opt['source_rest'] / 1000.0)
        
        max_waste_fraction_area = 0.8 if opt.get('pieces', 1) == 1 else 0.3
        if opt['type'] != 'multiple_transverse' and waste_area > source_area * max_waste_fraction_area:
            continue
        
        # Правило 5: Пропускаем transverse с отходами > 50% длины
        if opt['type'] == 'transverse':
            waste_fraction = waste_length / (opt['source_length'] * 1000) if opt['source_length'] > 0 else 0
            if waste_fraction > 0.5:
                continue
        
        filtered_secondary.append(opt)
    
    secondary_options = filtered_secondary
    print(f"[OPT_2D] После фильтрации осталось: {len(secondary_options)} вторичных опций")
    
    # 4. СОЗДАНИЕ ILP МОДЕЛИ
    prob = LpProblem("2D_Optimization", LpMinimize)
    
    # Переменные
    x_prim = {opt['id']: LpVariable(f"prim_{opt['id']}", lowBound=0, cat=LpInteger) 
              for opt in primary_options}
    x_sec = {opt['id']: LpVariable(f"sec_{opt['id']}", lowBound=0, cat=LpInteger) 
             for opt in secondary_options}
    
    # 5. ASSIGNMENT-MODEL: пары совместимости (opt × demand_key) и z-переменные
    #
    # Идея: вместо "сумма всех источников >= qty" с группировкой по frozenset,
    # вводим явные переменные распределения z_prim[(p, d)] / z_sec[(s, d)] —
    # сколько штук производства от опции p (или s) идёт на закрытие конкретного
    # спроса d. Это:
    #   * убирает двойной счёт переменных между разными группами,
    #   * полностью решает tolerance-edge cases на этапе построения пар,
    #   * делает покрытие точным (`demand_d == q_d`) — нет ни потерь, ни лишнего,
    #   * избавляет от костылей вроде demand_598665_min и пост-коррекции,
    #     которые лечили симптомы старой "flow + group-surplus" модели.
    primary_options_by_id = {o['id']: o for o in primary_options}
    secondary_options_by_id = {o['id']: o for o in secondary_options}

    dk_list = list(demand_2d.keys())
    dk_to_idx = {dk: i for i, dk in enumerate(dk_list)}
    primary_pairs_per_dk: dict = {dk: [] for dk in dk_list}    # dk -> [opt_id]
    secondary_pairs_per_dk: dict = {dk: [] for dk in dk_list}  # dk -> [opt_id]
    solid_pairs_per_dk: dict = {dk: [] for dk in dk_list}      # dk -> [opt_id] (solid only)
    no_sources_keys: list = []  # [(dk, qty), ...] — для совместимости с force-add safety net

    for dk in dk_list:
        target_length, target_width, target_load_code = dk
        for opt in primary_options:
            if (_canonical_length(opt['length']) != _canonical_length(target_length)
                    or opt.get('load_code', 800) != target_load_code):
                continue
            opt_type = opt.get('type')
            if opt_type in ('direct', 'solid'):
                if abs(opt['main'] - target_width) <= demand_tolerance_width:
                    primary_pairs_per_dk[dk].append(opt['id'])
                    if opt_type == 'solid' and target_width in solid_widths:
                        solid_pairs_per_dk[dk].append(opt['id'])
            elif opt_type == 'indirect':
                if abs(opt.get('target_width', 0) - target_width) <= demand_tolerance_width:
                    primary_pairs_per_dk[dk].append(opt['id'])

        for opt in secondary_options:
            opt_target_key = opt.get('target_order_key', (0, 0, 800))
            opt_target_load = opt_target_key[2] if len(opt_target_key) == 3 else 800
            if (_canonical_length(opt['output_length']) == _canonical_length(target_length)
                    and abs(opt['output_width'] - target_width) <= demand_tolerance_width
                    and opt_target_load == target_load_code):
                secondary_pairs_per_dk[dk].append(opt['id'])

        if not primary_pairs_per_dk[dk] and not secondary_pairs_per_dk[dk]:
            no_sources_keys.append((dk, demand_2d[dk]))
            import logging as _no_src_log
            _no_src_log.getLogger(__name__).error(
                "[OPT_2D] ❌ НЕТ ИСТОЧНИКОВ для плиты: %sм x %sмм (load=%s) x%dшт — закроется через unmet/post-correction",
                target_length, target_width, target_load_code, demand_2d[dk],
            )

    # 5.1 Z-переменные распределения (assignment) и slack
    z_prim: dict = {}      # (opt_id, dk) -> LpVariable, штук primary p, идущих на спрос d
    z_sec: dict = {}       # (opt_id, dk) -> LpVariable, штук secondary s, идущих на спрос d
    slack_solid: dict = {} # dk -> LpVariable: недопокрытие solid-priority (мягкий приоритет)
    unmet: dict = {}       # dk -> LpVariable: глобальный дефицит (последний рубеж)

    for dk in dk_list:
        di = dk_to_idx[dk]
        for opt_id in primary_pairs_per_dk[dk]:
            z_prim[(opt_id, dk)] = LpVariable(
                f"z_prim_{opt_id}_d{di}", lowBound=0, cat=LpInteger,
            )
        for opt_id in secondary_pairs_per_dk[dk]:
            z_sec[(opt_id, dk)] = LpVariable(
                f"z_sec_{opt_id}_d{di}", lowBound=0, cat=LpInteger,
            )
        if solid_pairs_per_dk[dk]:
            slack_solid[dk] = LpVariable(
                f"slack_solid_d{di}", lowBound=0, cat=LpInteger,
            )
        unmet[dk] = LpVariable(f"unmet_d{di}", lowBound=0, cat=LpInteger)

    # Сохранения для post-solve диагностики (вместо россыпи _dbg_*)
    _opt_to_demands = {}  # для logger ниже: opt_id -> [(dk, qty), ...]
    for dk, opt_ids in primary_pairs_per_dk.items():
        for oid in opt_ids:
            _opt_to_demands.setdefault(oid, []).append((dk, demand_2d[dk]))

    def _norm_key(k):
        if not k or len(k) < 2:
            return (0, 0, 800)
        lc = int(k[2]) if len(k) > 2 else 800
        if lc in (8, 800):
            lc = 8
        return (round(float(k[0]), 2), int(k[1]), lc)

    # 5.2 demand_d == q_d: точное закрытие спроса (assignment-вид).
    # Equality важно: одновременно "не теряем" (>=) и "не плодим" (<=).
    # При полном отсутствии источников спрос закроется через unmet[d] (см. 5.1)
    # с очень большим штрафом, чтобы модель никогда не была infeasible.
    for dk in dk_list:
        qty = demand_2d[dk]
        parts = [z_prim[(oid, dk)] for oid in primary_pairs_per_dk[dk]]
        parts += [z_sec[(oid, dk)] for oid in secondary_pairs_per_dk[dk]]
        parts.append(unmet[dk])
        prob += lpSum(parts) == qty, f"demand_d{dk_to_idx[dk]}"

    # 5.3 cap_prim: sum_d z_prim[p,d] <= x_p — производство ограничивает назначения
    prim_to_dks: dict = {}
    for _dk, _opts in primary_pairs_per_dk.items():
        for _oid in _opts:
            prim_to_dks.setdefault(_oid, []).append(_dk)
    for _oid, _dks in prim_to_dks.items():
        prob += (
            lpSum(z_prim[(_oid, _dk)] for _dk in _dks) <= x_prim[_oid],
            f"cap_prim_{_oid}",
        )

    # 5.4 cap_sec: sum_d z_sec[s,d] <= x_sec[s] * pieces_s
    sec_to_dks: dict = {}
    for _dk, _opts in secondary_pairs_per_dk.items():
        for _oid in _opts:
            sec_to_dks.setdefault(_oid, []).append(_dk)
    for _oid, _dks in sec_to_dks.items():
        _pieces = secondary_options_by_id[_oid].get('pieces', 1)
        prob += (
            lpSum(z_sec[(_oid, _dk)] for _dk in _dks) <= x_sec[_oid] * _pieces,
            f"cap_sec_{_oid}",
        )

    # 5.5 SOFT solid-priority: solid-плиты — приоритет для полных ширин (1200/1080),
    # но через slack_solid + штраф в objective, не через hard >=. Это убирает риск
    # infeasibility при нехватке solid-опций и оставляет видимый сигнал в логах.
    for _dk, _solid_ids in solid_pairs_per_dk.items():
        if not _solid_ids:
            continue
        _qty = demand_2d[_dk]
        prob += (
            lpSum(z_prim[(oid, _dk)] for oid in _solid_ids) + slack_solid[_dk] >= _qty,
            f"solid_priority_d{dk_to_idx[_dk]}",
        )
    
    if no_sources_keys:
        import logging as _no_src_summary_log
        _no_src_summary_log.getLogger(__name__).warning(
            "[OPT_2D] no_sources: %d ключей, %d плит — закроются через unmet/post-correction",
            len(no_sources_keys), sum(q for _, q in no_sources_keys),
        )
    # #region agent log
    try:
        import json as _agent_json
        import time as _agent_time
        _supply_diag = []
        for _dk in dk_list:
            _supply_diag.append({
                "dk": list(_dk),
                "need_qty": int(demand_2d.get(_dk, 0)),
                "primary_opts": len(primary_pairs_per_dk.get(_dk) or []),
                "secondary_opts": len(secondary_pairs_per_dk.get(_dk) or []),
                "solid_opts": len(solid_pairs_per_dk.get(_dk) or []),
            })
        with open(r"c:\Users\Роман\Desktop\Шишов\debug-ebb546.log", "a", encoding="utf-8") as _agent_f:
            _agent_f.write(_agent_json.dumps({
                "sessionId": "ebb546",
                "runId": "solver-localization",
                "hypothesisId": "O1,O2",
                "location": "core/optimization.py:before_solve:demand_option_supply",
                "message": "Demand keys и количество доступных опций до solve",
                "data": {
                    "demand_total": int(sum(demand_2d.values())),
                    "demand_keys": int(len(dk_list)),
                    "no_sources_keys": [{"dk": list(k), "qty": int(q)} for k, q in (no_sources_keys or [])],
                    "supply_diag": _supply_diag[:200],
                },
                "timestamp": int(_agent_time.time() * 1000),
            }, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # 6. ОГРАНИЧЕНИЯ: Баланс остатков с load_code в ключе.
    # Ключ остатка теперь (length, rest_width, load_code) — остатки разных load_code
    # не пулятся, что закрывает старый баг "общий пул остатков".
    rests_by_lkey: dict = {}  # rkey -> {'produced': [opt_id], 'consumed': [opt_id]}
    for opt in primary_options:
        rest_w = opt.get('rest', 0)
        if rest_w > 0 and opt.get('type') != 'solid':
            rkey = (opt['length'], rest_w, opt.get('load_code', 800))
            rests_by_lkey.setdefault(rkey, {'produced': [], 'consumed': []})['produced'].append(opt['id'])
    for opt in secondary_options:
        target_key = opt.get('target_order_key', (0, 0, 800))
        sec_lc = target_key[2] if len(target_key) == 3 else 800
        rkey = (opt['source_length'], opt['source_rest'], sec_lc)
        if rkey in rests_by_lkey:
            rests_by_lkey[rkey]['consumed'].append(opt['id'])
    for rkey, rec in rests_by_lkey.items():
        if rec['produced'] and rec['consumed']:
            prob += (
                lpSum(x_sec[i] for i in rec['consumed']) <= lpSum(x_prim[i] for i in rec['produced']),
                f"balance_L{rkey[0]}_R{rkey[1]}_LC{rkey[2]}",
            )
    
    # 7. ЦЕЛЕВАЯ ФУНКЦИЯ
    # Структура: cost_prim + cost_sec + waste + unused_rests + plates_priority
    #            + M_SOLID*slack_solid + M_UNMET*unmet
    # Иерархия штрафов: M_UNMET (1e7) >> M_SOLID (1e5) >> цены (1e3..1e4).
    # M_UNMET — последний рубеж: оплачиваем дефицит, чтобы модель никогда не была
    # infeasible. M_SOLID > цены, чтобы solid-priority доминировал в выборе типа реза,
    # но при отсутствии solid-опций solver не падал.
    M_UNMET = 1e7
    M_SOLID = 1e5

    obj_terms: list = []

    # 7.1 Стоимость первичных резов (цена плиты + продольный рез)
    print(f"[OPT_2D] Расчёт стоимости первичных резов...")
    for opt in primary_options:
        plate_price = get_price(opt['length'], 8, cfg.PRICE_DB_PATH) or 10000
        cut_cost = (cfg.LONG_CUT_PRICE_PER_M * opt['length']
                    if opt['type'] in ('direct', 'indirect') else 0)
        obj_terms.append(x_prim[opt['id']] * (plate_price + cut_cost))

    # 7.2 Стоимость вторичных резов (продольный + поперечный)
    for opt in secondary_options:
        if opt['type'] in ('narrowing', 'multiple', 'multiple_transverse'):
            obj_terms.append(x_sec[opt['id']] * cfg.LONG_CUT_PRICE_PER_M * opt['source_length'])
        if opt['type'] in ('transverse', 'multiple_transverse'):
            obj_terms.append(x_sec[opt['id']] * cfg.TRANSVERSE_CUT_PRICE)

    # 7.3 Штраф за неиспользованные остатки (rest produced, но не consumed)
    for rkey, rec in rests_by_lkey.items():
        if not (rec['produced'] and rec['consumed']):
            continue
        unused_expr = (lpSum(x_prim[i] for i in rec['produced'])
                       - lpSum(x_sec[i] for i in rec['consumed']))
        base_price = get_price(rkey[0], 6, cfg.PRICE_DB_PATH) or 5000
        rest_price = base_price * (rkey[1] / float(plate_width))
        obj_terms.append(unused_expr * rest_price * opt_config.unused_rest_penalty_coeff)

    # 7.4 Штраф за отходы (геометрический): waste по ширине и по длине, ~1000 руб/м²
    for opt in secondary_options:
        waste_w = opt.get('waste', 0)
        if waste_w > 0:
            waste_area_m2 = (waste_w / 1000.0) * opt['source_length']
            obj_terms.append(x_sec[opt['id']] * waste_area_m2 * 1000)
        waste_l = opt.get('length_waste', 0)
        if waste_l > 0:
            waste_area_m2 = (waste_l / 1000.0) * (opt['source_rest'] / 1000.0)
            obj_terms.append(x_sec[opt['id']] * waste_area_m2 * 1000)

    # 7.5 Бонус за переиспользование остатков (опционально через config)
    if opt_config.secondary_reuse_bonus:
        for opt in secondary_options:
            obj_terms.append(x_sec[opt['id']] * opt_config.secondary_reuse_bonus)

    # 7.6 Мягкий приоритет: меньше исходных плит = лучше (избегаем "раздутого" плана).
    obj_terms.append(lpSum(x_prim.values()) * 5000.0)

    # 7.7 Slack-штрафы — последняя линия защиты модели от infeasibility.
    if slack_solid:
        obj_terms.append(M_SOLID * lpSum(slack_solid.values()))
    if unmet:
        obj_terms.append(M_UNMET * lpSum(unmet.values()))

    print(f"[OPT_2D] Конфиг: unused_penalty={opt_config.unused_rest_penalty_coeff}, "
          f"reuse_bonus={opt_config.secondary_reuse_bonus}")
    prob += lpSum(t for t in obj_terms if t != 0)

    # 8. РЕШЕНИЕ
    print(f"[OPT_2D] Запуск решателя: {len(primary_options)} primary, "
          f"{len(secondary_options)} secondary, {len(z_prim) + len(z_sec)} z-vars, "
          f"{len(dk_list)} demand keys")
    prob.solve(PULP_CBC_CMD(msg=0, timeLimit=60, gapRel=0.005))
    
    # 8.1 Статус решателя и graceful fallback.
    # При Not Optimal не выходим пустотой — у нас есть unmet[d] как safety net,
    # post-correction добёрет недостающее. Это критично, чтобы пользователь
    # получил хоть какой-то план + видимый сигнал в логах.
    _solver_status = LpStatus[prob.status]
    print(f"[OPT_2D] Статус решателя: {_solver_status}")
    if _solver_status not in ('Optimal',):
        import logging as _solver_status_log
        _solver_status_log.getLogger(__name__).warning(
            "[OPT_2D] Решатель завершился со статусом %s — "
            "будет извлечён частичный результат, остатки добёрет post-correction",
            _solver_status,
        )
        # Если решатель не нашёл вообще никакого решения — возвращаем пустой результат,
        # чтобы upstream мог корректно отреагировать.
        if _solver_status in ('Infeasible', 'Undefined'):
            print(f"[OPT_2D] ⚠️ Решение не найдено! Статус: {_solver_status}")
            return {}

    # 8.2 Диагностика slack/unmet — ВИДИМЫЙ сигнал, что safety net сработал.
    _slack_total = int(round(sum((value(s) or 0) for s in slack_solid.values())))
    _unmet_total = int(round(sum((value(u) or 0) for u in unmet.values())))
    if _slack_total > 0 or _unmet_total > 0:
        import logging as _slack_diag_log
        _slack_logger = _slack_diag_log.getLogger(__name__)
        _slack_logger.warning(
            "[OPT_2D] [SLACK] solid_slack=%d, unmet=%d (см. unmet_d* / slack_solid_d*)",
            _slack_total, _unmet_total,
        )
        for dk, sv in unmet.items():
            v = value(sv) or 0
            if v > 0.5:
                _slack_logger.warning("[OPT_2D] [UNMET] %s = %d", dk, int(round(v)))
        for dk, sv in slack_solid.items():
            v = value(sv) or 0
            if v > 0.5:
                _slack_logger.warning("[OPT_2D] [SOLID_SLACK] %s = %d", dk, int(round(v)))
    # #region agent log
    try:
        import json as _agent_json
        import time as _agent_time

        _coverage_diag = []
        _missing_after_solver = []
        for _dk in dk_list:
            _need = int(demand_2d.get(_dk, 0))
            _z_prim = int(round(sum((value(z_prim.get((_oid, _dk))) or 0) for _oid in (primary_pairs_per_dk.get(_dk) or []))))
            _z_sec = int(round(sum((value(z_sec.get((_oid, _dk))) or 0) for _oid in (secondary_pairs_per_dk.get(_dk) or []))))
            _unmet_v = int(round(value(unmet.get(_dk)) or 0))
            _covered = _z_prim + _z_sec
            _coverage_diag.append({
                "dk": list(_dk),
                "need_qty": _need,
                "z_prim": _z_prim,
                "z_sec": _z_sec,
                "covered": _covered,
                "unmet": _unmet_v,
                "supply_primary_opts": len(primary_pairs_per_dk.get(_dk) or []),
                "supply_secondary_opts": len(secondary_pairs_per_dk.get(_dk) or []),
            })
            if _covered < _need:
                _missing_after_solver.append({
                    "dk": list(_dk),
                    "need_qty": _need,
                    "covered": _covered,
                    "missing": _need - _covered,
                    "unmet": _unmet_v,
                })
        with open(r"c:\Users\Роман\Desktop\Шишов\debug-ebb546.log", "a", encoding="utf-8") as _agent_f:
            _agent_f.write(_agent_json.dumps({
                "sessionId": "ebb546",
                "runId": "solver-localization",
                "hypothesisId": "O3,O4,O5",
                "location": "core/optimization.py:after_solve:z_coverage",
                "message": "Покрытие спроса по dk после solve через z_prim/z_sec/unmet",
                "data": {
                    "solver_status": _solver_status,
                    "demand_total": int(sum(demand_2d.values())),
                    "z_prim_total": int(round(sum((value(v) or 0) for v in z_prim.values()))),
                    "z_sec_total": int(round(sum((value(v) or 0) for v in z_sec.values()))),
                    "unmet_total": _unmet_total,
                    "slack_total": _slack_total,
                    "missing_after_solver_total": int(sum(x["missing"] for x in _missing_after_solver)),
                    "missing_after_solver": _missing_after_solver[:120],
                    "coverage_diag": _coverage_diag[:250],
                },
                "timestamp": int(_agent_time.time() * 1000),
            }, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # 9. ИЗВЛЕЧЕНИЕ РЕЗУЛЬТАТОВ
    from .plate_audit import PlateAudit as _PlateAudit
    import logging as _log_mod
    _audit = _PlateAudit(orders_2d)  # checkpoint "input" создаётся автоматически
    _audit.checkpoint("demand_2d", demand_2d)

    result = {
        'primary_cuts': [],
        'secondary_cuts': [],
        'total_plates': 0,
        'plate_assignments': [],
        # Сохраняем исходный заказ, чтобы визуализация и отчёты точно знали, что просил пользователь
        'orders_requested': [order.copy() for order in orders_2d],
        # Метаданные для отслеживания остатков
        'rests_created': [],  # Остатки, созданные при первичных резах
        'rests_used': []      # Остатки, использованные во вторичных резах
    }
    _next_primary_instance_id = 1
    _next_secondary_instance_id = 1
    _primary_instances_by_opt_id: dict[int, list[str]] = defaultdict(list)

    planned_primary_cuts = []
    planned_secondary_cuts = []

    # Первичные резы: извлекаем напрямую из z_prim[(p, d)] — каждая единица z = одна
    # плита, идущая на конкретный спрос d. assignment_key = dk: больше нет
    # неоднозначности "какой dk закрыла эта плита", которую старая модель решала
    # эвристикой по opt['target_width']/'main'.
    #
    # x_prim[p] >= sum_d z_prim[p,d] (cap_prim_p), поэтому если solver установил
    # x_prim больше суммы z, лишние плиты — это "пустые" слоты в физическом
    # смысле, и они не выходят в план (что и нужно: точное покрытие).
    import math as _math
    for (opt_id, dk), zv in z_prim.items():
        raw_val = value(zv) or 0
        qty = _math.ceil(raw_val - 1e-6) if raw_val > 1e-6 else 0
        if qty <= 0:
            continue
        opt = primary_options_by_id[opt_id]
        target_length, target_width, target_load_code = dk
        for _ in range(qty):
            primary_instance_id = f"prim-{_next_primary_instance_id}"
            _next_primary_instance_id += 1
            planned_primary_cuts.append({
                'width': opt['main'],
                'demand_width': target_width,
                'rest': opt['rest'],
                'qty': 1,
                'lengths': [opt['length']],
                'load_code': target_load_code,
                'assignment_key': dk,
                'source_opt_id': opt_id,
                'primary_instance_id': primary_instance_id,
            })
            _primary_instances_by_opt_id[opt_id].append(primary_instance_id)
            result['total_plates'] += 1

    # ========== НОВАЯ ЛОГИКА: СОРТИРОВКА ДЛЯ ПРОИЗВОДСТВА ==========
    # Требования завода:
    # 1. Первая плита ДОЛЖНА быть целой (без реза, rest=0)
    # 2. Плиты с одинаковым резом должны идти подряд
    print("[OPT_2D] 🔧 Применяем правила завода для порядка плит...")
    
    # Разделяем primary_cuts на группы
    solid_plates = []    # Целые плиты (rest=0)
    cut_plates = []      # Плиты с резом (rest>0)
    
    for cut in planned_primary_cuts:
        if cut['rest'] == 0:
            solid_plates.append(cut)
        else:
            cut_plates.append(cut)
    
    # ВАЖНО: Сортируем ЦЕЛЫЕ плиты тоже! (чтобы они не перемешивались с резанными)
    # Сортируем по ширине (целые 1200мм все вместе), потом по первой длине
    solid_plates.sort(key=lambda x: (-x['width'], -x['lengths'][0] if x.get('lengths') else 0))
    
    # Сортируем плиты с резом по (rest, width) — одинаковые резы идут подряд
    # Сортируем по убыванию ширины остатка, чтобы крупные резы шли в начале
    cut_plates.sort(key=lambda x: (-x['rest'], -x['width']))
    
    # Новый порядок: СНАЧАЛА целые (ОТСОРТИРОВАННЫЕ!), ПОТОМ плиты с резом (сгруппированные)
    result['primary_cuts'] = solid_plates + cut_plates

    # Вторичные резы:
    # 1) rests_used агрегируем по операциям — x_sec[s] — это сколько раз применена
    #    опция s. Каждое применение потребляет один остаток source_length × source_rest.
    # 2) secondary_cuts — по выходным плитам через z_sec[(s, d)]: каждая единица z
    #    это одна выходная плита, назначенная на конкретный спрос d.
    # cap_sec гарантирует sum_d z_sec[s,d] <= x_sec[s] * pieces_s; нераспределённые
    # выходы (если есть) — это "implicit waste", в план они не выходят.
    for opt in secondary_options:
        apps = int(round(value(x_sec[opt['id']]) or 0))
        for _ in range(apps):
            result['rests_used'].append({
                'source_length': opt['source_length'],
                'source_rest_mm': opt['source_rest'],
            })

    for (opt_id, dk), zv in z_sec.items():
        raw_val = value(zv) or 0
        qty = int(round(raw_val))
        if qty <= 0:
            continue
        opt = secondary_options_by_id[opt_id]
        target_length, target_width, target_load_code = dk
        for _ in range(qty):
            parent_instance_id = None
            for source_opt_id in opt.get('source_ids') or []:
                queue = _primary_instances_by_opt_id.get(source_opt_id) or []
                if queue:
                    parent_instance_id = queue.pop(0)
                    break
            secondary_instance_id = f"sec-{_next_secondary_instance_id}"
            _next_secondary_instance_id += 1
            planned_secondary_cuts.append({
                'source': opt['source_rest'],
                'cuts': [opt['output_width']],
                'qty': 1,
                'pieces': 1,
                'waste': opt.get('waste', 0),
                'type': opt['type'],
                'source_lengths': [opt['source_length']],
                'lengths': [opt['output_length']],
                'target_order_key': dk,
                'load_code': target_load_code,
                'parent_instance_id': parent_instance_id,
                'secondary_instance_id': secondary_instance_id,
                'source_opt_ids': list(opt.get('source_ids') or []),
            })

    result['secondary_cuts'] = planned_secondary_cuts

    # #region agent log
    try:
        import json as _agent_json
        import time as _agent_time
        with open(r"c:\Users\Роман\Desktop\Шишов\debug-ebb546.log", "a", encoding="utf-8") as _agent_f:
            _agent_f.write(_agent_json.dumps({
                "sessionId": "ebb546",
                "runId": "solver-localization",
                "hypothesisId": "O6",
                "location": "core/optimization.py:before_solver_output_checkpoint",
                "message": "Фактические primary/secondary перед PlateAudit solver_output",
                "data": {
                    "primary_cuts_len": len(result.get("primary_cuts") or []),
                    "secondary_cuts_len": len(result.get("secondary_cuts") or []),
                    "total_for_audit": len(result.get("primary_cuts") or []) + len(result.get("secondary_cuts") or []),
                    "secondary_sample": [
                        {
                            "source": c.get("source"),
                            "cuts": c.get("cuts"),
                            "lengths": c.get("lengths"),
                            "target_order_key": c.get("target_order_key"),
                            "load_code": c.get("load_code"),
                        }
                        for c in (result.get("secondary_cuts") or [])[:20]
                    ],
                },
                "timestamp": int(_agent_time.time() * 1000),
            }, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion

    # PlateAudit: checkpoint после сбора данных из solver (до пост-коррекции)
    _audit.checkpoint("solver_output", result['primary_cuts'] + result['secondary_cuts'])

    # Универсальная пост-коррекция: считаем уже покрытый спрос по ОБОИМ источникам,
    # а добираем только реально unmet demand.
    # BUG-3 FIX: нормализуем ключи через _norm_key, чтобы load_code 8/800 не давал
    # ложный «have=0» и не порождал дубликаты или пропуски.
    planned_coverage_by_key: Counter[tuple] = Counter()
    for cut in result['primary_cuts']:
        assignment_key = cut.get('assignment_key')
        if assignment_key:
            planned_coverage_by_key[_norm_key(assignment_key)] += 1
    for cut in result['secondary_cuts']:
        target_key = cut.get('target_order_key')
        if target_key:
            planned_coverage_by_key[_norm_key(target_key)] += 1

    _post_correction_added = 0
    for (L, W, lc), need in demand_2d.items():
        have = planned_coverage_by_key.get(_norm_key((L, W, lc)), 0)
        if have >= need:
            continue
        if W > plate_width:
            continue
        rest = (plate_width - W) if W < plate_width else 0
        for _ in range(need - have):
            primary_instance_id = f"prim-{_next_primary_instance_id}"
            _next_primary_instance_id += 1
            result['primary_cuts'].append({
                'width': W,
                'demand_width': W,
                'rest': rest,
                'qty': 1,
                'lengths': [L],
                'load_code': lc,
                'assignment_key': (L, W, lc),
                'identity_match_type': 'post_correction_pending',
                'source_opt_id': None,
                'primary_instance_id': primary_instance_id,
            })
            result['total_plates'] += 1
            _post_correction_added += 1
    if _post_correction_added:
        import logging as _logging
        _logging.getLogger(__name__).warning(
            '[OPT_2D] [POST_CORRECTION] Добавлено %d плит восстановлением покрытия спроса.',
            _post_correction_added,
        )

    # BUG-1 FIX: явная страховочная проверка для no_sources_keys.
    # Плиты без источников в модели должны быть добавлены как «прямой рез»,
    # если пост-коррекция их не покрыла (например, из-за W > plate_width скипа).
    if no_sources_keys:
        import logging as _logging
        _logger_ns = _logging.getLogger(__name__)
        # Пересчитываем покрытие с учётом только что добавленных пост-коррекцией
        _final_coverage: Counter[tuple] = Counter()
        for _cut in result['primary_cuts']:
            _ak = _cut.get('assignment_key')
            if _ak:
                _final_coverage[_norm_key(_ak)] += 1
        for _cut in result['secondary_cuts']:
            _tk = _cut.get('target_order_key')
            if _tk:
                _final_coverage[_norm_key(_tk)] += 1

        _force_added = 0
        for (_key, _qty) in no_sources_keys:
            _L, _W, _lc = _key
            _have = _final_coverage.get(_norm_key(_key), 0)
            if _have >= _qty:
                continue
            _shortfall = _qty - _have
            _rest = max(0, plate_width - _W) if _W <= plate_width else 0
            for _ in range(_shortfall):
                primary_instance_id = f"prim-{_next_primary_instance_id}"
                _next_primary_instance_id += 1
                result['primary_cuts'].append({
                    'width': _W,
                    'demand_width': _W,
                    'rest': _rest,
                    'qty': 1,
                    'lengths': [_L],
                    'load_code': _lc,
                    'assignment_key': (_L, _W, _lc),
                    'identity_match_type': 'force_added_no_sources',
                    'source_opt_id': None,
                    'primary_instance_id': primary_instance_id,
                })
                result['total_plates'] += 1
                _force_added += 1
                _final_coverage[_norm_key(_key)] += 1
        if _force_added:
            _logger_ns.warning(
                '[OPT_2D] [BUG-1 FIX] Force-added %d плит из no_sources_keys '
                '(не было источника в оптимизаторе, добавлены как прямой рез).',
                _force_added,
            )

    # PlateAudit: checkpoint после пост-коррекции и force-add (финальный срез оптимизатора)
    _audit.checkpoint("post_correction", result['primary_cuts'] + result['secondary_cuts'])
    if _audit.has_losses("demand_2d", "post_correction"):
        import logging as _log_mod2
        _log_mod2.getLogger(__name__).error(
            "[AUDIT] Потери плит в оптимизаторе!\n%s", _audit.summary()
        )
    else:
        import logging as _log_mod2
        _log_mod2.getLogger(__name__).info("[AUDIT] Оптимизатор: потерь нет.\n%s", _audit.summary())
    # #region agent log
    try:
        _missing_after_post = {}
        for (_L, _W, _lc), _need in demand_2d.items():
            _have = planned_coverage_by_key.get(_norm_key((_L, _W, _lc)), 0)
            if _have < _need:
                _missing_after_post[str([_L, _W, _lc])] = int(_need - _have)
        _debug_runtime_write_648532(
            "run1",
            "H3_post_correction_balance",
            "optimization:after_post_correction",
            "Demand minus planned coverage after post-correction",
            {
                "demand_total": int(sum(demand_2d.values())),
                "planned_coverage_total": int(sum(planned_coverage_by_key.values())),
                "missing_keys_after_post": _missing_after_post,
                "missing_total_after_post": int(sum(_missing_after_post.values())),
            },
        )
    except Exception:
        pass
    # #endregion
    result['_plate_audit'] = _audit

    # ОБЯЗАТЕЛЬНАЯ ПРОВЕРКА ПОКРЫТИЯ: логируем итог по demand_2d vs primary+secondary.
    # Сейчас работает как наблюдатель; в этап 2 модель должна выдавать ok=True всегда.
    _coverage_summary = verify_coverage(demand_2d, result['primary_cuts'], result['secondary_cuts'])
    result['_coverage_summary'] = _coverage_summary
    import logging as _log_cov
    _cov_logger = _log_cov.getLogger(__name__)
    if not _coverage_summary['ok']:
        _cov_logger.error(
            "[OPT_2D] [COVERAGE] missing=%d плит по %d ключам, surplus=%d плит по %d ключам",
            sum(_coverage_summary['missing'].values()),
            len(_coverage_summary['missing']),
            sum(_coverage_summary['surplus'].values()),
            len(_coverage_summary['surplus']),
        )
    else:
        _cov_logger.info(
            "[OPT_2D] [COVERAGE] OK: demand=%d, covered=%d, surplus=%d по %d ключам",
            _coverage_summary['demand_total'],
            _coverage_summary['covered_total'],
            sum(_coverage_summary['surplus'].values()),
            len(_coverage_summary['surplus']),
        )

    # #region agent log (2d5c43) H1,H2,H5: demand vs primary_cuts, 6m 530/1200
    try:
        from pathlib import Path as _Path
        _log_2d5c43 = _Path(__file__).resolve().parent.parent / "debug-2d5c43.log"
        _demand_total = sum(demand_2d.values())
        _target_keys = [(6.0, 1200, 8), (6.0, 530, 8), (5.1, 320, 8)]
        _demand_by_key = {}
        for k, q in demand_2d.items():
            for tk in _target_keys:
                if abs(round(k[0], 2) - round(tk[0], 2)) <= 0.02 and k[1] == tk[1] and (k[2] == tk[2] or k[2] == '8'):
                    _demand_by_key[tuple(tk)] = _demand_by_key.get(tuple(tk), 0) + q
                    break
        _prim_total = len(result['primary_cuts'])
        _prim_by_key = {tk: 0 for tk in _target_keys}
        _prim_6_530_1200 = []
        for c in result['primary_cuts']:
            L = round((c.get('lengths') or [0])[0], 2)
            W = c.get('demand_width', c.get('width', 0))
            lc = c.get('load_code', 8)
            for tk in _target_keys:
                if abs(L - tk[0]) <= 0.02 and W == tk[1] and (lc == tk[2] or lc == '8'):
                    _prim_by_key[tk] = _prim_by_key.get(tk, 0) + 1
                    break
            if 5.98 <= L <= 6.02 and W in (530, 1200) and len(_prim_6_530_1200) < 25:
                _prim_6_530_1200.append({"length": L, "width": W, "rest": c.get('rest', 0), "plate_name": (c.get('plate_name') or "")[:60]})
        _demand_by_key_ser = [list(k) + [v] for k, v in _demand_by_key.items()]
        _prim_by_key_ser = [list(k) + [v] for k, v in _prim_by_key.items()]
        with _dbg_open_append(_log_2d5c43) as _f:
            _f.write(__import__('json').dumps({"sessionId": "2d5c43", "hypothesisId": "H1_H2_H5", "location": "optimization:_optimize_2d:primary_cuts_done", "message": "demand vs primary_cuts", "data": {"demand_total": _demand_total, "demand_by_key": _demand_by_key_ser, "primary_total": _prim_total, "primary_by_key": _prim_by_key_ser, "primary_6m_530_1200_sample": _prim_6_530_1200}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    print(f"[OPT_2D] ✓ Целых плит в начале: {len(solid_plates)}")
    print(f"[OPT_2D] ✓ Плит с резом (сгруппировано): {len(cut_plates)}")

    # Финальная атрибуция primary: теперь расходуем общий slot-ledger строго по demand key.
    _empty_primary_keys = []

    # Пересоздаём plate_assignments в правильном порядке
    # ИСПРАВЛЕНИЕ: добавляем load_code, иначе в треках подставляется 8п и списание ищет 8п вместо 10п
    result['plate_assignments'] = []
    for cut in result['primary_cuts']:
        assignment_key = cut.get('assignment_key') or (
            (cut.get('lengths') or [0])[0],
            cut.get('demand_width', cut.get('width')),
            cut.get('load_code', 800),
        )
        plate_info = _next_slot_info(slot_lists, slot_cursors, assignment_key)
        if not plate_info:
            _empty_primary_keys.append(assignment_key)

        cut['load_code'] = plate_info.get('load_code', cut.get('load_code', 800)) if plate_info else cut.get('load_code', 800)
        cut['kp_id'] = plate_info.get('kp_id') if plate_info else None
        cut['customer'] = plate_info.get('customer') if plate_info else None
        cut['kp_date'] = plate_info.get('kp_date') if plate_info else None
        cut['plate_name'] = plate_info.get('plate_name') if plate_info else None
        cut['identity_match_type'] = plate_info.get('identity_match_type') if plate_info else 'slot_exhausted'

        for length in cut['lengths']:
            result['plate_assignments'].append({
                'length': length,
                'width': cut['width'],
                'source': 'primary',
                'rest_width': cut['rest'],
                'kp_id': cut.get('kp_id'),
                'customer': cut.get('customer'),
                'kp_date': cut.get('kp_date'),
                'plate_name': cut.get('plate_name'),
                'load_code': cut.get('load_code', 800),
                'identity_match_type': cut.get('identity_match_type'),
                'unit_id': cut.get('primary_instance_id'),
                'parent_unit_id': None,
                'source_opt_id': cut.get('source_opt_id'),
            })
            
            # Сохраняем информацию об остатке для отслеживания
            if cut['rest'] > 0:
                result['rests_created'].append({
                    'length': length,
                    'rest_width_mm': cut['rest'],
                    'source_width_mm': cut['width']
                })
    # #region agent log: summary empty plate_info (H2)
    if _empty_primary_keys:
        try:
            _c = Counter(_empty_primary_keys)
            _summary = [{"key": list(k), "count": _c[k]} for k in sorted(_c.keys())]
            with _dbg_open_append(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log") as _f:
                _f.write(__import__('json').dumps({"hypothesisId": "H2", "location": "optimization.py:primary_emit_summary", "message": "plan keys with empty plate_info (summary)", "data": {"by_key": _summary, "total_plates_empty": len(_empty_primary_keys), "unique_keys": len(_c)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
        except Exception:
            pass
    # #endregion

    # #region agent log (session 5b5324) после пересборки primary plate_assignments
    try:
        _prim_sample = [{"kp_id": p.get("kp_id"), "plate_name": (p.get("plate_name") or "")[:50], "identity_match_type": p.get("identity_match_type")} for p in result['plate_assignments'][:5]]
        with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
            _f.write(__import__('json').dumps({"sessionId": "5b5324", "hypothesisId": "H_primary_pa", "location": "optimization:plate_assignments_after_primary", "message": "primary plate_assignments count and sample", "data": {"primary_plate_assignments_count": len(result['plate_assignments']), "sample": _prim_sample}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    
    print(f"[OPT_2D] ✓ Создано остатков: {len(result['rests_created'])}")
    # ========== КОНЕЦ НОВОЙ ЛОГИКИ ==========

    # Вторичные резы: получаем identity из того же общего slot-ledger.
    _secondary_attribution_log = []  # (target_key, kp_id, plate_name, match_type) для лога 5b5324
    _empty_secondary_keys = []
    for cut in result['secondary_cuts']:
        target_key = cut.get('target_order_key')
        plate_info = _next_slot_info(slot_lists, slot_cursors, target_key) if target_key else {}
        if target_key and not plate_info:
            _empty_secondary_keys.append(target_key)
        if len(_secondary_attribution_log) < 50:
            _secondary_attribution_log.append({
                "target_key": list(target_key) if isinstance(target_key, tuple) else target_key,
                "kp_id": plate_info.get('kp_id') if plate_info else None,
                "plate_name": (plate_info.get('plate_name') or '')[:50] if plate_info else None,
                "match_type": plate_info.get('identity_match_type') if plate_info else 'empty',
            })
        # #region agent log: secondary plate kp_id (H2, H5)
        if not plate_info and target_key:
            try:
                with _dbg_open_append(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log") as _f:
                    _f.write(__import__('json').dumps({"hypothesisId": "H2", "location": "optimization.py:secondary_emit", "message": "secondary plate_info empty", "data": {"target_key": list(target_key) if isinstance(target_key, tuple) else target_key, "output_length": cut.get('lengths', [None])[0], "output_width": cut.get('cuts', [None])[0]}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
            except Exception:
                pass
        # #endregion
        cut['kp_id'] = plate_info.get('kp_id') if plate_info else None
        cut['customer'] = plate_info.get('customer') if plate_info else None
        cut['kp_date'] = plate_info.get('kp_date') if plate_info else None
        cut['plate_name'] = plate_info.get('plate_name') if plate_info else None
        cut['identity_match_type'] = plate_info.get('identity_match_type') if plate_info else 'secondary_unmapped'

        result['plate_assignments'].append({
            'length': cut['lengths'][0],
            'width': cut['cuts'][0],
            'source': 'secondary',
            'source_rest': cut['source'],
            'kp_id': cut.get('kp_id'),
            'customer': cut.get('customer'),
            'kp_date': cut.get('kp_date'),
            'plate_name': cut.get('plate_name'),
            'load_code': cut.get('load_code', 800),
            'identity_match_type': cut.get('identity_match_type'),
            'unit_id': cut.get('secondary_instance_id'),
            'parent_unit_id': cut.get('parent_instance_id'),
        })
    # #region agent log (session 5b5324) после вторичных резов
    try:
        _sec_count = sum(1 for p in result['plate_assignments'] if p.get('source') == 'secondary')
        with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
            _f.write(__import__('json').dumps({"sessionId": "5b5324", "hypothesisId": "H_secondary", "location": "optimization:after_secondary", "message": "secondary attributions and plate_assignments", "data": {"secondary_attribution_sample": _secondary_attribution_log[:40], "plate_assignments_total": len(result['plate_assignments']), "secondary_count": _sec_count}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # #region agent log
    try:
        _slot_exhausted_by_key = {}
        if _empty_primary_keys:
            _slot_exhausted_by_key["primary"] = {str(list(k)): int(v) for k, v in Counter(_empty_primary_keys).items()}
        if _empty_secondary_keys:
            _slot_exhausted_by_key["secondary"] = {str(list(k)): int(v) for k, v in Counter(_empty_secondary_keys).items()}
        _slot_cursor_overview = []
        for _k in list(slot_lists.keys())[:80]:
            _slot_cursor_overview.append(
                {
                    "key": list(_k),
                    "cursor": int(slot_cursors.get(_k, 0)),
                    "slots": int(len(slot_lists.get(_k, []))),
                }
            )
        _debug_runtime_write_648532(
            "run1",
            "H2_slot_exhaustion",
            "optimization:after_slot_attribution",
            "Slot consumption and exhausted attribution keys",
            {
                "plate_assignments_total": int(len(result.get("plate_assignments", []))),
                "empty_primary_count": int(len(_empty_primary_keys)),
                "empty_secondary_count": int(len(_empty_secondary_keys)),
                "slot_exhausted_by_key": _slot_exhausted_by_key,
                "slot_cursor_overview": _slot_cursor_overview,
            },
        )
    except Exception:
        pass
    # #endregion
    
    print(f"[OPT_2D] OK! Готово! Использовано {result['total_plates']} плит")
    print(f"[OPT_2D] Создано {len(result['plate_assignments'])} готовых плит")
    print(f"[OPT_2D] Остатков использовано вторично: {len(result['rests_used'])}")
    # #region agent log: result counts (H5) + plates by key (H_rescue_trace)
    try:
        _dbg_open_append(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log").write(__import__('json').dumps({"sessionId": "debug-session", "runId": "run1", "hypothesisId": "H5", "location": "optimization.py:result_built", "message": "plate_assignments count", "data": {"len_plate_assignments": len(result['plate_assignments']), "total_plates": result.get('total_plates', 0), "demand_sum": sum(demand_2d.values())}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # Лог по ключам (length, width, load_code) — для сравнения с дорожками и РЕСКЬЮ (любые размеры).
    try:
        _norm_lc = getattr(cfg, 'normalize_load_code', lambda x: int(x) if x is not None else 8)
        _by_key = {}
        for cut in result.get('primary_cuts', []):
            lc = _norm_lc(cut.get('load_code', 8))
            for L in cut.get('lengths', []):
                k = (round(float(L), 2), int(cut.get('width', 0)), lc)
                _by_key[k] = _by_key.get(k, 0) + 1
        for cut in result.get('secondary_cuts', []):
            lengths = cut.get('lengths', [])
            widths = cut.get('cuts', [])
            tk = cut.get('target_order_key')
            lc = _norm_lc(tk[2] if isinstance(tk, (tuple, list)) and len(tk) > 2 else 8)
            L = float(lengths[0]) if lengths else 0
            W = int(widths[0]) if widths else (int(tk[1]) if isinstance(tk, (tuple, list)) and len(tk) > 1 else 0)
            if L and W:
                k = (round(L, 2), W, lc)
                _by_key[k] = _by_key.get(k, 0) + 1
        _dbg_open_append(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log").write(__import__('json').dumps({"hypothesisId": "H_opt_plates_by_key", "location": "optimization.py:result_built", "message": "optimizer output plates by (length, width, load_code)", "data": {"plates_by_key": {str(list(k)): v for k, v in _by_key.items()}, "total": sum(_by_key.values())}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    return result


def _optimize_1d_widths_only(orders: dict, plate_width: int = 1200, 
                              min_useful_width: int = 200) -> dict:
    """
    ПРИВАТНАЯ функция: Оптимизация только по ширинам (1D).
    Длины НЕ учитываются в оптимизации, присваиваются позже.
    
    Args:
        orders: {300: 4, 500: 3} — спрос по ширинам в мм (без учёта длин!)
        plate_width: ширина исходной плиты в мм (1200)
        min_useful_width: минимальная полезная ширина остатка
    
    Returns:
        {
            'primary_cuts': [{'width': 320, 'rest': 880, 'qty': 2}, ...],
            'secondary_cuts': [{'source': 880, 'cuts': [320, 560], 'qty': 1}, ...],
            'total_plates': 5,
            'total_cost': 75000,
            'waste_width': 120
        }
    """
    try:
        from pulp import (
            LpProblem, LpMinimize, LpVariable, LpInteger, lpSum, value,
            PULP_CBC_CMD, LpStatus,
        )
    except ImportError:
        print('[OPT_CASCADING] PuLP не установлен, пропускаем.')
        return {}
    
    if not orders:
        return {}
    
    # Преобразуем заказы в список ширин с количеством
    target_widths = sorted(orders.keys())
    
    # Допустимый диапазон для каждой ширины (±20 мм)
    tolerance = 20
    
    # Генерируем все возможные варианты первичных резов (из плиты заданной ширины)
    # Для каждой целевой ширины создаём варианты: target_width + остаток
    primary_cut_options = []
    solid_widths = sorted(set([plate_width, 1080]))
    for target_w in target_widths:
        if target_w in solid_widths:
            primary_cut_options.append({
                'id': f'prim_{target_w}',
                'main': target_w,
                'rest': 0,
            })
            continue

        # Пропил косвенно учтён в таблице NARROWING_TABLE
        rest_w = plate_width - target_w
        if rest_w >= min_useful_width:  # Остаток достаточно большой
            primary_cut_options.append({
                'id': f'prim_{target_w}',
                'main': target_w,
                'rest': rest_w,
            })
    
    # Генерируем варианты вторичных резов (из остатков)
    # Для каждого возможного остатка смотрим, на какие ширины его можно разрезать
    secondary_cut_options = []
    possible_rests = set(opt['rest'] for opt in primary_cut_options if opt['rest'] > 0)
    
    for rest_w in possible_rests:
        # Пробуем разрезать остаток на 2 части
        for target_w1 in target_widths:
            target_w2 = rest_w - target_w1
            # Проверяем, подходит ли вторая часть для какой-то из целевых ширин
            for target_w2_candidate in target_widths:
                if abs(target_w2 - target_w2_candidate) <= tolerance:
                    secondary_cut_options.append({
                        'id': f'sec_{rest_w}_to_{target_w1}_{target_w2_candidate}',
                        'source_rest': rest_w,
                        'output1': target_w1,
                        'output2': target_w2_candidate,
                        'waste': abs(rest_w - target_w1 - target_w2_candidate),
                    })
                    break
            
            # Также проверяем вариант: остаток целиком режем на несколько одинаковых частей
            # Например, 880 мм → 2 части по 320 мм (остаток 240 мм)
            for target_w_candidate in target_widths:
                # Считаем, сколько частей нужной ширины влезет в остаток
                num_pieces = rest_w // target_w_candidate
                if num_pieces >= 2:  # Минимум 2 куска, иначе не выгодно
                    waste = rest_w - (target_w_candidate * num_pieces)
                    if waste < rest_w * 0.5:  # Отход < 50% остатка (разумное ограничение)
                        secondary_cut_options.append({
                            'id': f'sec_{rest_w}_to_{num_pieces}x{target_w_candidate}',
                            'source_rest': rest_w,
                            'output1': target_w_candidate,
                            'output2': 0,  # вторая "часть" не используется
                            'pieces': num_pieces,
                            'waste': waste,
                        })
                        # Не ставим break - проверяем все варианты
            
            # НОВАЯ ЛОГИКА: Сужение (narrowing) - из остатка делаем ОДНУ плиту с небольшим отходом
            # Например, 340 мм → 320 мм (отход 20 мм)
            for target_w_candidate in target_widths:
                if target_w_candidate < rest_w <= target_w_candidate + 100:  # Остаток чуть больше целевой ширины
                    waste = rest_w - target_w_candidate
                    # Разрешаем сужение до 100 мм (но лучше меньше)
                    if waste <= 100:
                        secondary_cut_options.append({
                            'id': f'sec_{rest_w}_narrow_to_{target_w_candidate}',
                            'source_rest': rest_w,
                            'output1': target_w_candidate,
                            'output2': 0,
                            'pieces': 1,  # Только ОДНА плита
                            'waste': waste,
                        })
    
    print(f"[DEBUG] Найдено вариантов вторичных резов: {len(secondary_cut_options)}")
    for opt in secondary_cut_options[:3]:  # Показываем первые 3
        print(f"  {opt['source_rest']}мм -> {opt.get('pieces', 1)}x{opt['output1']}мм (отход {opt['waste']}мм)")
    
    # Создаём задачу оптимизации
    prob = LpProblem('cascading_longitudinal_cuts', LpMinimize)
    
    # Переменные: количество первичных резов каждого типа
    x_prim = {}
    for i, opt in enumerate(primary_cut_options):
        x_prim[i] = LpVariable(f"prim_{i}_{opt['id']}", lowBound=0, cat=LpInteger)
    
    # Переменные: количество вторичных резов каждого типа
    x_sec = {}
    for i, opt in enumerate(secondary_cut_options):
        x_sec[i] = LpVariable(f"sec_{i}_{opt['id']}", lowBound=0, cat=LpInteger)
    
    # ASSIGNMENT-MODEL: пары совместимости (option × target_width).
    # Один проход даёт три коллекции:
    #   - primary_pairs_per_w[w]: список (i, opt) совместимых первичных опций;
    #   - secondary_pairs_per_w[w]: (i, opt, contribution) — pieces для output1,
    #     либо 1 для output2.
    #   - на основе пар вводим z_prim_w[(i, w)] и z_sec_w[(i, w)].
    target_widths_keys = list(orders.keys())
    primary_pairs_per_w: dict = {w: [] for w in target_widths_keys}
    secondary_pairs_per_w: dict = {w: [] for w in target_widths_keys}

    for w in target_widths_keys:
        for i, opt in enumerate(primary_cut_options):
            if abs(opt['main'] - w) <= tolerance:
                primary_pairs_per_w[w].append(i)
        for i, opt in enumerate(secondary_cut_options):
            if abs(opt['output1'] - w) <= tolerance:
                secondary_pairs_per_w[w].append((i, opt.get('pieces', 1)))
            if opt['output2'] > 0 and abs(opt['output2'] - w) <= tolerance:
                secondary_pairs_per_w[w].append((i, 1))

    z_prim_w: dict = {}  # (i, w) -> LpVariable, штук primary i, идущих на спрос w
    z_sec_w: dict = {}   # (i, w) -> LpVariable, штук secondary i (по общему вкладу),
                         #          идущих на спрос w
    unmet_w: dict = {}   # w -> LpVariable: глобальный дефицит (последний рубеж)
    for w in target_widths_keys:
        for i in primary_pairs_per_w[w]:
            z_prim_w[(i, w)] = LpVariable(f"z1d_prim_{i}_w{w}", lowBound=0, cat=LpInteger)
        for (i, _) in secondary_pairs_per_w[w]:
            z_sec_w[(i, w)] = LpVariable(f"z1d_sec_{i}_w{w}", lowBound=0, cat=LpInteger)
        unmet_w[w] = LpVariable(f"unmet_w{w}", lowBound=0, cat=LpInteger)

    # 1) demand_w == qty[w]: точное закрытие; unmet[w] закрывает экстремальные кейсы.
    for w, qty in orders.items():
        parts = [z_prim_w[(i, w)] for i in primary_pairs_per_w[w]]
        parts += [z_sec_w[(i, w)] for (i, _) in secondary_pairs_per_w[w]]
        parts.append(unmet_w[w])
        prob += lpSum(parts) == qty, f"demand_w{w}"

    # 2) cap_prim: sum_w z_prim_w[i,w] <= x_prim[i]
    prim_to_ws_1d: dict = {}
    for w, ids in primary_pairs_per_w.items():
        for i in ids:
            prim_to_ws_1d.setdefault(i, []).append(w)
    for i, ws in prim_to_ws_1d.items():
        prob += (
            lpSum(z_prim_w[(i, w)] for w in ws) <= x_prim[i],
            f"cap_prim_w_{i}",
        )

    # 3) cap_sec: sum_w z_sec_w[i,w]*share_unit <= x_sec[i] * pieces_total.
    # Здесь z_sec_w[i,w] — это штуки выходных плит (с учётом и output1*pieces,
    # и output2*1). Капа: сумма этих штук для опции i не должна превышать
    # количество произведённых выходов = x_sec[i] * total_outputs_i.
    sec_to_ws_1d: dict = {}
    for w, pairs in secondary_pairs_per_w.items():
        for (i, contrib) in pairs:
            sec_to_ws_1d.setdefault(i, []).append((w, contrib))
    for i, ws_contribs in sec_to_ws_1d.items():
        opt = secondary_cut_options[i]
        # Сколько выходных плит даёт ОДНО применение опции i:
        # output1 даёт pieces штук, output2 (если есть и >0) даёт ещё 1 штуку.
        outputs_per_app = opt.get('pieces', 1)
        if opt.get('output2', 0) > 0:
            outputs_per_app += 1
        prob += (
            lpSum(z_sec_w[(i, w)] for (w, _) in ws_contribs) <= x_sec[i] * outputs_per_app,
            f"cap_sec_w_{i}",
        )

    # 4) Баланс остатков: sum_consumed <= sum_produced (как и было).
    for rest_w in possible_rests:
        produced = [x_prim[i] for i, opt in enumerate(primary_cut_options) if opt['rest'] == rest_w]
        consumed = [x_sec[i] for i, opt in enumerate(secondary_cut_options) if opt['source_rest'] == rest_w]
        if produced and consumed:
            prob += lpSum(consumed) <= lpSum(produced), f"balance_rest_{rest_w}"

    # 5) Целевая функция: total_plates + unused_rests_penalty + waste_penalty
    #    + M_unmet_w * unmet (последний рубеж — обычно 0, но избавляет от infeasible).
    M_UNMET_W = 1e7

    total_plates = lpSum(x_prim.values())

    unused_rests_penalty = 0
    for rest_w in possible_rests:
        produced = [x_prim[i] for i, opt in enumerate(primary_cut_options) if opt['rest'] == rest_w]
        consumed = [x_sec[i] for i, opt in enumerate(secondary_cut_options) if opt['source_rest'] == rest_w]
        if produced and consumed:
            unused = lpSum(produced) - lpSum(consumed)
            unused_rests_penalty += unused * (rest_w / 1000.0) * 0.05

    waste_penalty = 0
    for i, opt in enumerate(secondary_cut_options):
        waste_penalty += x_sec[i] * opt.get('waste', 0) * 0.0001

    obj = total_plates + unused_rests_penalty + waste_penalty
    if unmet_w:
        obj = obj + M_UNMET_W * lpSum(unmet_w.values())
    prob += obj
    prob.solve(PULP_CBC_CMD(msg=0, timeLimit=60, gapRel=0.005))

    _solver_status_1d = LpStatus[prob.status]
    if _solver_status_1d not in ('Optimal',):
        import logging as _solver_status_log_1d
        _solver_status_log_1d.getLogger(__name__).warning(
            "[OPT_1D] Решатель завершился со статусом %s — извлекаем частичный результат",
            _solver_status_1d,
        )
        if _solver_status_1d in ('Infeasible', 'Undefined'):
            return {}

    _unmet_total_1d = int(round(sum((value(u) or 0) for u in unmet_w.values())))
    if _unmet_total_1d > 0:
        import logging as _unmet_log_1d
        _unmet_log_1d.getLogger(__name__).warning(
            "[OPT_1D] [UNMET] Дефицит: %d плит — модель не нашла источников",
            _unmet_total_1d,
        )
    
    # Извлекаем результаты
    result = {
        'primary_cuts': [],
        'secondary_cuts': [],
        'total_plates': 0,
        'total_cost': 0,
        'waste_width': 0,
    }
    
    for i, opt in enumerate(primary_cut_options):
        try:
            qty = int(round(value(x_prim[i])))
            if qty > 0:
                result['primary_cuts'].append({
                    'width': opt['main'],
                    'rest': opt['rest'],
                    'qty': qty,
                })
                result['total_plates'] += qty
        except:
            pass
    
    for i, opt in enumerate(secondary_cut_options):
        try:
            qty = int(round(value(x_sec[i])))
            if qty > 0:
                cuts = [opt['output1']]
                if opt['output2'] > 0:
                    cuts.append(opt['output2'])
                result['secondary_cuts'].append({
                    'source': opt['source_rest'],
                    'cuts': cuts,
                    'pieces': opt.get('pieces', 1),
                    'qty': qty,
                    'waste': opt.get('waste', 0),
                })
                result['waste_width'] += opt.get('waste', 0) * qty
        except:
            pass
    
    # ========== НОВАЯ ЛОГИКА: СОРТИРОВКА ДЛЯ ПРОИЗВОДСТВА ==========
    # Требования завода:
    # 1. Первая плита ДОЛЖНА быть целой (без реза, rest=0)
    # 2. Плиты с одинаковым резом должны идти подряд
    print("[OPT_1D] 🔧 Применяем правила завода для порядка плит...")
    
    # Разделяем primary_cuts на группы
    solid_plates = []    # Целые плиты (rest=0)
    cut_plates = []      # Плиты с резом (rest>0)
    
    for cut in result['primary_cuts']:
        if cut['rest'] == 0:
            solid_plates.append(cut)
        else:
            cut_plates.append(cut)
    
    # ВАЖНО: Сортируем ЦЕЛЫЕ плиты тоже! (чтобы они не перемешивались с резанными)
    # Сортируем по ширине (целые 1200мм все вместе), потом по first элементу
    if solid_plates:
        solid_plates.sort(key=lambda x: (-x['width']))
    
    # Сортируем плиты с резом по (rest, width) — одинаковые резы идут подряд
    # Сортируем по убыванию ширины остатка, чтобы крупные резы шли в начале
    cut_plates.sort(key=lambda x: (-x['rest'], -x['width']))
    
    # Новый порядок: СНАЧАЛА целые (ОТСОРТИРОВАННЫЕ!), ПОТОМ плиты с резом (сгруппированные)
    result['primary_cuts'] = solid_plates + cut_plates
    
    print(f"[OPT_1D] ✓ Целых плит в начале: {len(solid_plates)}")
    print(f"[OPT_1D] ✓ Плит с резом (сгруппировано): {len(cut_plates)}")
    # ========== КОНЕЦ НОВОЙ ЛОГИКИ ==========
    
    print(f"[DEBUG] Оптимизатор выбрал вторичных резов: {len(result['secondary_cuts'])}")
    
    # Рассчитываем стоимость
    plate_price = 12000  # примерная цена плиты
    long_cut_cost = 460 * 6  # продольный рез (460 руб/м × 6 м средняя длина)
    
    result['total_cost'] = (
        result['total_plates'] * plate_price +
        result['total_plates'] * long_cut_cost +  # первичные резы
        len(result['secondary_cuts']) * long_cut_cost  # вторичные резы
    )
    
    return result


def optimize_with_cascading_longitudinal_cuts(orders: dict = None, 
                                               orders_2d: list = None,
                                               plate_width: int = 1200, 
                                               min_useful_width: int = 200,
                                               opt_config: OptimizationConfig = None) -> dict:
    """
    Универсальная оптимизация с каскадными резами (PUBLIC API).
    
    АВТОМАТИЧЕСКИ ВЫБИРАЕТ РЕЖИМ на основе входных данных:
    
    Режим 1D (старый, обратная совместимость):
        >>> result = optimize_with_cascading_longitudinal_cuts(
        ...     orders={320: 14, 860: 9}
        ... )
        # Оптимизирует по ШИРИНАМ, длины присваиваются позже
    
    Режим 2D (новый, полная оптимизация):
        >>> result = optimize_with_cascading_longitudinal_cuts(
        ...     orders_2d=[
        ...         {'length': 5.6, 'width': 320, 'qty': 11},
        ...         {'length': 6.63, 'width': 860, 'qty': 4}
        ...     ]
        ... )
        # Полная 2D оптимизация (длина + ширина) в ILP модели
    
    Args:
        orders: {width: qty} — спрос по ширинам (для режима 1D)
        orders_2d: [{'length', 'width', 'qty'}] — спрос 2D (для режима 2D)
        plate_width: ширина исходной плиты (1200 мм)
        min_useful_width: минимальная полезная ширина остатка
        opt_config: конфигурация параметров оптимизации (штрафы, бонусы)
    
    Returns:
        dict с результатами оптимизации:
            {
                'primary_cuts': [...],
                'secondary_cuts': [...],
                'total_plates': int,
                'plate_assignments': [...],  # только для режима 2D
                ...
            }
    """
    
    # АВТООПРЕДЕЛЕНИЕ РЕЖИМА
    if orders_2d is not None and len(orders_2d) > 0:
        # ===== РЕЖИМ 2D (НОВЫЙ) =====
        print("[OPT] Режим: ПОЛНАЯ 2D оптимизация (длина + ширина)")
        return _optimize_2d_with_lengths(orders_2d, plate_width, min_useful_width, opt_config)
    
    elif orders is not None and len(orders) > 0:
        # ===== РЕЖИМ 1D (СТАРЫЙ) =====
        print("[OPT] Режим: 1D оптимизация (только ширина, обратная совместимость)")
        return _optimize_1d_widths_only(orders, plate_width, min_useful_width)
    
    else:
        print("[OPT] ⚠️ Не указаны ни orders, ни orders_2d!")
        return {}


# ==================== FFD ОПТИМИЗАЦИЯ РАСКРОЯ ДОРОЖЕК ====================

from dataclasses import dataclass, field


@dataclass
class Piece:
    """Кусок плиты для укладки в дорожку"""
    length_m: float
    qty: int
    kind: str              # 'standard' | 'addon'
    load_class: float
    width_m: float = 1.196


@dataclass
class Track:
    """Дорожка (линия производства)"""
    width_m: float = 1.196
    total_m: float = 0.0
    pieces: list = field(default_factory=list)
    leftover_m: float = 0.0


def first_fit_decreasing(
    pieces: list[Piece],
    stock_len_m: float = 9.88
) -> list[Track]:
    """
    Алгоритм First Fit Decreasing для оптимизации раскроя
    Минимизирует количество дорожек (плит-заготовок)
    
    Args:
        pieces: Список Piece объектов (куски для размещения)
        stock_len_m: Длина заготовки (максимальная длина плиты)
        
    Returns:
        Список Track объектов (дорожек) с размещёнными кусками
    """
    pool = []
    
    # Сортируем по убыванию длины (FFD алгоритм)
    sorted_pieces = sorted(pieces, key=lambda x: x.length_m, reverse=True)
    
    # Развёртываем количество в отдельные элементы
    expanded = []
    for p in sorted_pieces:
        for _ in range(p.qty):
            expanded.append(Piece(p.length_m, 1, p.kind, p.load_class, p.width_m))
    
    # Размещаем каждый кусок
    for piece in expanded:
        placed = False
        
        # Пробуем поместить в существующие дорожки
        for track in pool:
            if track.total_m + piece.length_m <= stock_len_m:
                track.pieces.append(piece)
                track.total_m += piece.length_m
                placed = True
                break
        
        # Если не поместился, создаём новую дорожку
        if not placed:
            track = Track()
            track.pieces.append(piece)
            track.total_m = piece.length_m
            pool.append(track)
    
    # Вычисляем остатки
    for track in pool:
        track.leftover_m = stock_len_m - track.total_m
    
    return pool


def optimize_tracks(
    items: list,
    stock_len_m: float = 9.88
) -> dict:
    """
    Оптимизирует размещение плит в дорожки
    
    Args:
        items: Список позиций [{'length_m': float, 'qty': int, 'kind': str, 'load_class': float}, ...]
        stock_len_m: Длина заготовки (максимальная длина)
        
    Returns:
        Словарь с результатами оптимизации
    """
    pieces = []
    
    for item in items:
        pieces.append(Piece(
            length_m=item.get('length_m', 0),
            qty=item.get('qty', 1),
            kind=item.get('kind', 'standard'),
            load_class=item.get('load_class', 8.0),
            width_m=item.get('width_m', 1.196)
        ))
    
    tracks = first_fit_decreasing(pieces, stock_len_m)
    
    # Статистика
    total_tracks = len(tracks)
    total_used = sum(t.total_m for t in tracks)
    total_leftover = sum(t.leftover_m for t in tracks)
    efficiency = (total_used / (total_tracks * stock_len_m) * 100) if total_tracks > 0 else 0
    
    return {
        'tracks': tracks,
        'total_tracks': total_tracks,
        'total_used_m': round(total_used, 2),
        'total_leftover_m': round(total_leftover, 2),
        'efficiency_pct': round(efficiency, 1),
        'stock_length_m': stock_len_m
    }


