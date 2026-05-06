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
import logging
from pathlib import Path as _Path
from typing import Any, Callable

from core import config_and_data as cfg
from core.config_and_data import canonical_plate_key
from core.debug_paths import PROJECT_ROOT, get_debug_log_path
from dataclasses import dataclass
from core.optimization.geometry import (
    GeometryConfig,
    KERF_WIDTH_MM,
    NARROWING_TABLE,
    _canonical_length,
    filter_secondary_cut_options_2d,
    generate_primary_cut_options_1d,
    generate_primary_cut_options_2d,
    generate_raw_secondary_cut_options_2d,
    generate_secondary_cut_options_1d,
)
from core.optimization.ffd_packing import (
    Piece,
    Track,
    first_fit_decreasing,
    optimize_tracks,
    pack_tracks,
)
from core.optimization.debug_log import (
    _DEBUG_LOG_5b5324,
    _DEBUG_LOG_COMMON,
    _dbg_open_append,
    _opt_debug_enabled,
)
from core.optimization.order_dispatch import (
    _build_proportional_slot_lists,
    _get_next_order_info,
    _next_slot_info,
    _peek_order_info,
    build_order_info_list,
)
from core.optimization.ilp_model import (
    _build_residual_balance_constraints,
    _residual_phys_key,
    build_two_d_cutting_ilp,
)
from core.optimization.logging_utils import order_line_for_console
from core.optimization.result_contract import (
    ERROR_EMPTY_ORDERS_1D,
    ERROR_EMPTY_ORDERS_2D,
    ERROR_PULP_MISSING,
    ERROR_SOLVER_INFEASIBLE,
    ERROR_SOLVER_UNDEFINED,
    is_optimization_success,
    opt_error,
    opt_ok,
)

_DEBUG_AGENT_LOG_EBB546 = get_debug_log_path("debug-ebb546.log")
_DEBUG_RUNTIME_LOG_648532 = get_debug_log_path("debug-648532.log")
_DEBUG_LOG_2D5C43 = get_debug_log_path("debug-2d5c43.log")
_DEBUG_RUNTIME_SESSION_ID_648532 = "648532"

_OPT_1D_LOG = logging.getLogger(__name__)


def _opt_1d_pulp_nonneg_qty(
    pulp_value_fn: Callable[[Any], Any],
    var: Any,
    *,
    context: str,
) -> int:
    """
    Целое qty ≥ 0 из решённой PuLP-переменной (ветка 1D).

    ``None`` от value() → 0 и предупреждение (нестабильное/частичное решение без молчаливой порчи списка).
    Ошибки преобразования и сбои value() логируются и превращаются в ValueError.
    """
    try:
        raw = pulp_value_fn(var)
    except Exception as exc:
        _OPT_1D_LOG.exception(
            "[OPT_1D] pulp.value() выбросил исключение для %s",
            context,
        )
        raise ValueError(f"pulp.value failed for {context}") from exc
    if raw is None:
        _OPT_1D_LOG.warning(
            "[OPT_1D] %s: value() вернул None после решения — qty=0",
            context,
        )
        return 0
    try:
        qty = int(round(float(raw)))
    except (TypeError, ValueError, OverflowError) as exc:
        _OPT_1D_LOG.error(
            "[OPT_1D] %s: не удалось преобразовать value=%r в int",
            context,
            raw,
            exc_info=True,
        )
        raise ValueError(f"invalid pulp value for {context}: {raw!r}") from exc
    if qty < 0:
        _OPT_1D_LOG.error("[OPT_1D] %s: отрицательное qty=%s", context, qty)
        raise ValueError(f"negative qty for {context}: {qty}")
    return qty


# ==================== ХЕЛПЕРЫ ОПТИМИЗАЦИИ ====================


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

# ==================== ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ ОПТИМИЗАЦИИ (TLS + прокси, см. context.py) ====================

from core.optimization.context import (
    LOAD_TO_REINFORCEMENT_MAP,
    OPT_CASCADING_PLAN,
    OPT_CASCADING_PLAN_BY_LOAD,
    OPT_PLAN,
    OPT_WIDTH_PRIORITY,
)


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

    if not orders:
        return opt_error(
            ERROR_EMPTY_ORDERS_1D,
            "Нет исходных заказов по ширине для optimize_cuts_pulp.",
        )
    result = optimize_with_cascading_longitudinal_cuts(orders=orders)
    if is_optimization_success(result):
        OPT_PLAN.clear()
        OPT_PLAN.update(result)
    return result


# ==================== СОВРЕМЕННЫЕ ФУНКЦИИ ОПТИМИЗАЦИИ ====================


def _batch_sizes_for_secondary_z_sec(qty: int, pieces: int) -> list[int]:
    """
    Разбиение qty строк выхода z_sec на батчи: один родительский остаток на батч
    длиной до pieces (ограничение cap_sec в ILP: sum z_sec <= x_sec * pieces).

    Examples: qty=3, pieces=2 -> [2, 1].
    """
    p = max(1, int(pieces or 1))
    sizes: list[int] = []
    offset = 0
    while offset < qty:
        b = min(p, qty - offset)
        sizes.append(b)
        offset += b
    return sizes


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
        from pulp import PULP_CBC_CMD, LpStatus, value
    except ImportError:
        print('[OPT_2D] PuLP не установлен.')
        return opt_error(
            ERROR_PULP_MISSING,
            "PuLP не установлен — 2D ILP недоступен.",
        )

    if not orders_2d:
        return opt_error(ERROR_EMPTY_ORDERS_2D, "Пустой список заказов orders_2d.")
    
    # #region agent log
    try:
        import json as _aj
        import time as _at
        with open(_Path(__file__).resolve().parent.parent / "debug-7e420e.log", "a", encoding="utf-8") as _lf:
            _lf.write(
                _aj.dumps(
                    {
                        "sessionId": "7e420e",
                        "hypothesisId": "H_OPT_ENTER",
                        "location": "optimization.py:_optimize_2d_with_lengths:entry",
                        "message": "2D ILP optimizer entered (fresh plan build)",
                        "data": {"n_orders": len(orders_2d), "plate_width": int(plate_width)},
                        "timestamp": int(_at.time() * 1000),
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
    except Exception:
        pass
    # #endregion
    
    print(f"\n[OPT_2D] === ПОЛНАЯ 2D ОПТИМИЗАЦИЯ ===")
    print(f"[OPT_2D] Заказ:")
    for order in orders_2d:
        print(order_line_for_console(order))
    
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
            _dbg_open_append(_DEBUG_LOG_COMMON).write(__import__("json").dumps({"hypothesisId": "H_59_10_demand", "location": "optimization.py:demand_2d_built", "message": "demand_2d: ключи 5.99м 10п (length, width, load_code)", "data": {"keys": _demand_59_10}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
    # #endregion
    # 1.5 НОВОЕ: маппинг (length, width, load_code) -> список записей КП (тот же формат ключей, что demand_2d).
    order_info_list = build_order_info_list(orders_2d, cfg)

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
    geometry_config = GeometryConfig(
        plate_width=plate_width,
        min_useful_width=min_useful_width,
        tolerance_width=tolerance_width,
    )
    primary_result = generate_primary_cut_options_2d(
        demand_2d=demand_2d,
        order_info_list=order_info_list,
        order_info_getter=_peek_order_info,
        config=geometry_config,
    )
    primary_options = primary_result.raw_options
    solid_widths = primary_result.solid_widths

    print(f"[OPT_2D] Таблица narrowing создана: {len(NARROWING_TABLE)} правил")
    for target_w, sources in primary_result.target_to_sources.items():
        print(f"  {target_w}мм можно получить через: {sources}")

    print(f"[OPT_2D] Опций первичных резов (до фильтрации): {len(primary_options)}")
    # #region agent log (2d5c43) Plan B: опции для 5.1/320 и 6/530 до фильтра
    try:
        _log = _DEBUG_LOG_2D5C43
        _opts_320 = [{"id": o['id'], "length": o['length'], "main": o['main'], "type": o.get('type'), "load_code": o.get('load_code')} for o in primary_options if o.get('main') == 320 or o.get('target_width') == 320]
        _opts_530 = [{"id": o['id'], "length": o['length'], "main": o['main'], "type": o.get('type'), "load_code": o.get('load_code')} for o in primary_options if o.get('main') == 530 or o.get('target_width') == 530]
        with _dbg_open_append(_log) as _f:
            _f.write(__import__('json').dumps({"sessionId": "2d5c43", "hypothesisId": "H_opt_gen", "location": "optimization:primary_options_after_build", "message": "options for 320 and 530 before filter", "data": {"opts_320": _opts_320, "opts_530": _opts_530, "solid_widths": solid_widths}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # 2.5 ФИЛЬТРАЦИЯ ПЕРВИЧНЫХ ОПЦИЙ (Улучшение 4: убираем заведомо невыгодные)
    primary_options = primary_result.options
    print(f"[OPT_2D] После фильтрации осталось: {len(primary_options)} первичных опций")
    # #region agent log (2d5c43) Plan B: опции для 320/530 после фильтра
    try:
        _log = _DEBUG_LOG_2D5C43
        _opts_320 = [{"id": o['id'], "length": o['length'], "main": o['main'], "type": o.get('type')} for o in primary_options if o.get('main') == 320 or o.get('target_width') == 320]
        _opts_530 = [{"id": o['id'], "length": o['length'], "main": o['main'], "type": o.get('type')} for o in primary_options if o.get('main') == 530 or o.get('target_width') == 530]
        with _dbg_open_append(_log) as _f:
            _f.write(__import__('json').dumps({"sessionId": "2d5c43", "hypothesisId": "H_opt_filter", "location": "optimization:primary_options_after_filter", "message": "options for 320 and 530 after filter", "data": {"opts_320": _opts_320, "opts_530": _opts_530}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # 3. ГЕНЕРАЦИЯ ОПЦИЙ ВТОРИЧНЫХ РЕЗОВ (2D: длина + ширина!)
    secondary_options = generate_raw_secondary_cut_options_2d(
        primary_options=primary_options,
        demand_2d=demand_2d,
        config=geometry_config,
    )
    print(f"[OPT_2D] Опций вторичных резов (до фильтрации): {len(secondary_options)}")

    # 3.5 ФИЛЬТРАЦИЯ ВТОРИЧНЫХ ОПЦИЙ (Улучшение 4: убираем дубликаты и невыгодные)
    secondary_options = filter_secondary_cut_options_2d(secondary_options)
    print(f"[OPT_2D] После фильтрации осталось: {len(secondary_options)} вторичных опций")
    
    def _norm_key(k):
        if not k or len(k) < 2:
            return (0, 0, 800)
        lc = int(k[2]) if len(k) > 2 else 800
        if lc in (8, 800):
            lc = 8
        return (round(float(k[0]), 2), int(k[1]), lc)

    # 4–7. ILP: переменные, ограничения, целевая функция (см. ilp_model)
    ilp = build_two_d_cutting_ilp(
        demand_2d=demand_2d,
        primary_options=primary_options,
        secondary_options=secondary_options,
        solid_widths=solid_widths,
        plate_width=plate_width,
        demand_tolerance_width=demand_tolerance_width,
        opt_config=opt_config,
    )
    prob = ilp.prob
    x_prim = ilp.x_prim
    x_sec = ilp.x_sec
    z_prim = ilp.z_prim
    z_sec = ilp.z_sec
    slack_solid = ilp.slack_solid
    unmet = ilp.unmet
    dk_list = ilp.dk_list
    dk_to_idx = ilp.dk_to_idx
    primary_pairs_per_dk = ilp.primary_pairs_per_dk
    secondary_pairs_per_dk = ilp.secondary_pairs_per_dk
    solid_pairs_per_dk = ilp.solid_pairs_per_dk
    primary_options_by_id = ilp.primary_options_by_id
    secondary_options_by_id = ilp.secondary_options_by_id
    no_sources_keys = ilp.no_sources_keys
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
        with open(_DEBUG_AGENT_LOG_EBB546, "a", encoding="utf-8") as _agent_f:
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
            _code = (
                ERROR_SOLVER_UNDEFINED
                if _solver_status == "Undefined"
                else ERROR_SOLVER_INFEASIBLE
            )
            return opt_error(
                _code,
                f"Решатель завершился со статусом {_solver_status}.",
                solver_status=_solver_status,
            )

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
        with open(_DEBUG_AGENT_LOG_EBB546, "a", encoding="utf-8") as _agent_f:
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
    from core.plate_audit import PlateAudit as _PlateAudit
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
    _primary_instances_by_geom_lc: dict[tuple[float, int], list[tuple[float | int, int, str]]] = defaultdict(list)

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
            if opt.get('rest', 0) > 0:
                _primary_instances_by_geom_lc[_residual_phys_key(opt['length'], opt['rest'])].append((
                    cfg.normalize_load_code(opt.get('load_code', target_load_code), default=8),
                    opt_id,
                    primary_instance_id,
                ))
            result['total_plates'] += 1

    # #region agent log
    try:
        import json as _aj
        import time as _at
        _geom_prim_counts: dict[str, int] = {}
        for _pc in planned_primary_cuts:
            if _pc.get("rest", 0) <= 0:
                continue
            _L0 = _canonical_length(_pc["lengths"][0]) if _pc.get("lengths") else 0.0
            _rk = f"{_L0}_{int(round(float(_pc['rest'])))}"
            _geom_prim_counts[_rk] = _geom_prim_counts.get(_rk, 0) + 1
        _opt_queue_lens_before_sec = {str(k): len(v) for k, v in _primary_instances_by_opt_id.items() if v}
        with open(_Path(__file__).resolve().parent.parent / "debug-7e420e.log", "a", encoding="utf-8") as _lf:
            _lf.write(
                _aj.dumps(
                    {
                        "sessionId": "7e420e",
                        "hypothesisId": "H_OPT_PRIMARY_GEOM",
                        "location": "optimization.py:after_z_prim_planned_primary",
                        "message": "primary splits count by (len_m, rest_mm); opt_id queues before any secondary pop",
                        "data": {
                            "n_planned_primary": len(planned_primary_cuts),
                            "geom_split_counts": _geom_prim_counts,
                            "opt_id_queue_nonempty": _opt_queue_lens_before_sec,
                        },
                        "timestamp": int(_at.time() * 1000),
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
    except Exception:
        pass
    # #endregion

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

    def _remove_primary_instance_from_geom(instance_id: str) -> None:
        for pool in _primary_instances_by_geom_lc.values():
            for idx, (_prim_lc, _opt_id, _inst_id) in enumerate(pool):
                if _inst_id == instance_id:
                    del pool[idx]
                    return

    _orphan_recovered_geometry = 0
    _secondary_parent_missing = 0

    for (opt_id, dk), zv in z_sec.items():
        raw_val = value(zv) or 0
        qty = int(round(raw_val))
        if qty <= 0:
            continue
        opt = secondary_options_by_id[opt_id]
        target_length, target_width, target_load_code = dk
        target_load_code = cfg.normalize_load_code(target_load_code, default=8)
        pieces = max(1, int(opt.get("pieces") or 1))
        if qty % pieces != 0:
            import logging as _batch_qty_log

            _batch_qty_log.getLogger(__name__).warning(
                "[OPT_2D] z_sec qty=%d not divisible by pieces=%d for sec_opt_id=%s dk=%s — "
                "last chunk batched as one parental rest",
                qty,
                pieces,
                opt_id,
                list(dk) if isinstance(dk, (list, tuple)) else dk,
            )

        def _pop_parent_for_secondary_rest() -> str | None:
            nonlocal _orphan_recovered_geometry
            parent_id: str | None = None
            for source_opt_id in opt.get("source_ids") or []:
                queue = _primary_instances_by_opt_id.get(source_opt_id) or []
                if queue:
                    parent_id = queue.pop(0)
                    _remove_primary_instance_from_geom(parent_id)
                    break
            if not parent_id:
                pool = (
                    _primary_instances_by_geom_lc.get(
                        _residual_phys_key(opt.get("source_length"), opt.get("source_rest"))
                    )
                    or []
                )
                for idx, (prim_lc, source_opt_id, instance_id) in enumerate(pool):
                    if cfg.normalize_load_code(prim_lc, default=8) >= target_load_code:
                        parent_id = instance_id
                        del pool[idx]
                        opt_queue = _primary_instances_by_opt_id.get(source_opt_id) or []
                        if instance_id in opt_queue:
                            opt_queue.remove(instance_id)
                        _orphan_recovered_geometry += 1
                        break
            return parent_id

        batch_sizes_list = _batch_sizes_for_secondary_z_sec(qty, pieces)
        z_block_offset = 0
        for batch_index, batch_size in enumerate(batch_sizes_list):
            _q_before = {
                str(_soid): len(_primary_instances_by_opt_id.get(_soid) or [])
                for _soid in (opt.get("source_ids") or [])
            }
            parent_instance_id = _pop_parent_for_secondary_rest()
            _q_mid = {
                str(_soid): len(_primary_instances_by_opt_id.get(_soid) or [])
                for _soid in (opt.get("source_ids") or [])
            }
            if not parent_instance_id:
                _secondary_parent_missing += batch_size
            for _ in range(batch_size):
                secondary_instance_id = f"sec-{_next_secondary_instance_id}"
                _next_secondary_instance_id += 1
                # #region agent log
                try:
                    import json as _aj
                    import time as _at

                    _q_after = {
                        str(_soid): len(_primary_instances_by_opt_id.get(_soid) or [])
                        for _soid in (opt.get("source_ids") or [])
                    }
                    with open(
                        _Path(__file__).resolve().parent.parent / "debug-7e420e.log",
                        "a",
                        encoding="utf-8",
                    ) as _lf:
                        _lf.write(
                            _aj.dumps(
                                {
                                    "sessionId": "7e420e",
                                    "hypothesisId": "H_OPT_SEC_PARENT_POP",
                                    "location": "optimization.py:z_sec_parent_assignment",
                                    "message": "z_sec output row shares parent within pieces-batch",
                                    "data": {
                                        "sec_opt_id": opt_id,
                                        "pieces": pieces,
                                        "batch_index": batch_index,
                                        "batch_size": batch_size,
                                        "batch_offset_in_z_block": z_block_offset,
                                        "source_length": opt.get("source_length"),
                                        "source_rest": opt.get("source_rest"),
                                        "target_order_key": list(dk)
                                        if isinstance(dk, (list, tuple))
                                        else dk,
                                        "source_ids": list(opt.get("source_ids") or []),
                                        "queue_lens_before_pop": _q_before,
                                        "queue_remaining_after_parent_pop": _q_mid,
                                        "queue_remaining_by_source_opt_id": _q_after,
                                        "parent_instance_id": parent_instance_id,
                                        "secondary_instance_id": secondary_instance_id,
                                        "sec_type": opt.get("type"),
                                    },
                                    "timestamp": int(_at.time() * 1000),
                                },
                                ensure_ascii=False,
                                default=str,
                            )
                            + "\n"
                        )
                except Exception:
                    pass
                # #endregion
                planned_secondary_cuts.append({
                    "source": opt["source_rest"],
                    "cuts": [opt["output_width"]],
                    "qty": 1,
                    "pieces": 1,
                    "waste": opt.get("waste", 0),
                    "type": opt["type"],
                    "source_lengths": [opt["source_length"]],
                    "lengths": [opt["output_length"]],
                    "target_order_key": dk,
                    "load_code": target_load_code,
                    "parent_instance_id": parent_instance_id,
                    "secondary_instance_id": secondary_instance_id,
                    "source_opt_ids": list(opt.get("source_ids") or []),
                })
            z_block_offset += batch_size

    result['secondary_cuts'] = planned_secondary_cuts
    if _orphan_recovered_geometry or _secondary_parent_missing:
        import logging as _parent_log
        _parent_log.getLogger(__name__).warning(
            "[OPT_2D] secondary parent assignment: recovered_by_geometry=%d, missing=%d",
            _orphan_recovered_geometry,
            _secondary_parent_missing,
        )

    # #region agent log
    try:
        import json as _aj
        import time as _at
        _null_parent = sum(1 for c in planned_secondary_cuts if not c.get("parent_instance_id"))
        _by_geom = {}
        for c in planned_secondary_cuts:
            if c.get("parent_instance_id"):
                continue
            sl = c.get("source_lengths") or []
            _L = _canonical_length(sl[0]) if sl else None
            _src = c.get("source")
            _gk = f"{_L}_{int(round(float(_src)))}" if _L is not None and _src is not None else "?"
            _by_geom[_gk] = _by_geom.get(_gk, 0) + 1
        with open(_Path(__file__).resolve().parent.parent / "debug-7e420e.log", "a", encoding="utf-8") as _lf:
            _lf.write(
                _aj.dumps(
                    {
                        "sessionId": "7e420e",
                        "hypothesisId": "H_OPT_SEC_SUMMARY",
                        "location": "optimization.py:after_planned_secondary_cuts",
                        "message": "secondary cuts: null parent count and breakdown by geom key",
                        "data": {
                            "n_secondary": len(planned_secondary_cuts),
                            "null_parent_count": _null_parent,
                            "null_parent_by_source_geom_key": _by_geom,
                        },
                        "timestamp": int(_at.time() * 1000),
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
    except Exception:
        pass
    # #endregion

    # #region agent log
    try:
        import json as _agent_json
        import time as _agent_time
        with open(_DEBUG_AGENT_LOG_EBB546, "a", encoding="utf-8") as _agent_f:
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
        _log_2d5c43 = _DEBUG_LOG_2D5C43
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
            with _dbg_open_append(_DEBUG_LOG_COMMON) as _f:
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
                with _dbg_open_append(_DEBUG_LOG_COMMON) as _f:
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
    # #region agent log (ef42ae: H1 оптимизатор дал вторичку без родителя; H3 «ошибочно вторичный» по watchlist)
    try:
        import json as _ef42_json
        import time as _ef42_time

        _sec_asg = [p for p in result["plate_assignments"] if p.get("source") == "secondary"]
        _null_par = [p for p in _sec_asg if not p.get("parent_unit_id")]
        _watch_subs = ("25,4-3", "63,9-5,3", "42,6-5,3", "25,4-3,0", "63,9-5,3-10")
        _watch_sec = [
            {
                "unit_id": p.get("unit_id"),
                "parent_unit_id": p.get("parent_unit_id"),
                "length": p.get("length"),
                "width": p.get("width"),
                "plate_name": (p.get("plate_name") or "")[:160],
                "identity_match_type": p.get("identity_match_type"),
            }
            for p in _sec_asg
            if any(s in str(p.get("plate_name") or "") for s in _watch_subs)
        ]
        _prim_asg = [p for p in result["plate_assignments"] if p.get("source") != "secondary"]
        _watch_prim = [
            {
                "length": p.get("length"),
                "width": p.get("width"),
                "plate_name": (p.get("plate_name") or "")[:160],
                "source": p.get("source"),
            }
            for p in _prim_asg
            if any(s in str(p.get("plate_name") or "") for s in _watch_subs)
        ]
        _raw_sec_no_parent = [
            {
                "secondary_instance_id": c.get("secondary_instance_id"),
                "parent_instance_id": c.get("parent_instance_id"),
                "target_order_key": list(c["target_order_key"])
                if isinstance(c.get("target_order_key"), tuple)
                else c.get("target_order_key"),
                "lengths": c.get("lengths"),
                "cuts": c.get("cuts"),
            }
            for c in (result.get("secondary_cuts") or [])
            if not c.get("parent_instance_id")
        ][:120]
        with open(PROJECT_ROOT / "debug-ef42ae.log", "a", encoding="utf-8") as _ef42_f:
            _ef42_f.write(
                _ef42_json.dumps(
                    {
                        "sessionId": "ef42ae",
                        "hypothesisId": "H1_H3",
                        "location": "core/optimization/_implementation.py:after_secondary_plate_assignments",
                        "message": "secondary assignments: parent null count, raw secondary_cuts without parent, SKU watchlist primary vs secondary",
                        "data": {
                            "n_secondary_assignments": len(_sec_asg),
                            "n_secondary_null_parent": len(_null_par),
                            "n_raw_secondary_cuts": len(result.get("secondary_cuts") or []),
                            "raw_secondary_no_parent_count": len(
                                [c for c in (result.get("secondary_cuts") or []) if not c.get("parent_instance_id")]
                            ),
                            "null_parent_assignments_head": [
                                {
                                    "unit_id": p.get("unit_id"),
                                    "parent_unit_id": p.get("parent_unit_id"),
                                    "length": p.get("length"),
                                    "width": p.get("width"),
                                    "plate_name": (p.get("plate_name") or "")[:120],
                                }
                                for p in _null_par[:100]
                            ],
                            "raw_secondary_no_parent_head": _raw_sec_no_parent,
                            "watchlist_secondary": _watch_sec,
                            "watchlist_primary_rows": _watch_prim[:60],
                        },
                        "timestamp": int(_ef42_time.time() * 1000),
                    },
                    ensure_ascii=False,
                    default=str,
                )
                + "\n"
            )
    except Exception:
        pass
    # #endregion
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
        _dbg_open_append(_DEBUG_LOG_COMMON).write(__import__('json').dumps({"sessionId": "debug-session", "runId": "run1", "hypothesisId": "H5", "location": "optimization.py:result_built", "message": "plate_assignments count", "data": {"len_plate_assignments": len(result['plate_assignments']), "total_plates": result.get('total_plates', 0), "demand_sum": sum(demand_2d.values())}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
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
        _dbg_open_append(_DEBUG_LOG_COMMON).write(__import__('json').dumps({"hypothesisId": "H_opt_plates_by_key", "location": "optimization.py:result_built", "message": "optimizer output plates by (length, width, load_code)", "data": {"plates_by_key": {str(list(k)): v for k, v in _by_key.items()}, "total": sum(_by_key.values())}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    return opt_ok(result, partial=(_solver_status != "Optimal"))


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
        return opt_error(
            ERROR_PULP_MISSING,
            "PuLP не установлен — 1D оптимизация недоступна.",
        )

    if not orders:
        return opt_error(ERROR_EMPTY_ORDERS_1D, "Пустой словарь заказов по ширине (1D).")
    
    # Преобразуем заказы в список ширин с количеством
    target_widths = sorted(orders.keys())
    
    # Допустимый диапазон для каждой ширины (±20 мм)
    tolerance = 20
    
    # Генерируем все возможные варианты первичных резов (из плиты заданной ширины)
    # Для каждой целевой ширины создаём варианты: target_width + остаток
    primary_cut_options = generate_primary_cut_options_1d(
        target_widths=target_widths,
        plate_width=plate_width,
        min_useful_width=min_useful_width,
    )

    # Генерируем варианты вторичных резов (из остатков)
    # Для каждого возможного остатка смотрим, на какие ширины его можно разрезать
    secondary_cut_options = generate_secondary_cut_options_1d(
        primary_cut_options=primary_cut_options,
        target_widths=target_widths,
        tolerance=tolerance,
    )
    possible_rests = set(opt['rest'] for opt in primary_cut_options if opt['rest'] > 0)
    
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
            _code = (
                ERROR_SOLVER_UNDEFINED
                if _solver_status_1d == "Undefined"
                else ERROR_SOLVER_INFEASIBLE
            )
            return opt_error(
                _code,
                f"1D решатель: {_solver_status_1d}.",
                solver_status=_solver_status_1d,
            )

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
        qty = _opt_1d_pulp_nonneg_qty(value, x_prim[i], context=f"x_prim[{i}]")
        if qty > 0:
            result['primary_cuts'].append({
                'width': opt['main'],
                'rest': opt['rest'],
                'qty': qty,
            })
            result['total_plates'] += qty

    for i, opt in enumerate(secondary_cut_options):
        qty = _opt_1d_pulp_nonneg_qty(value, x_sec[i], context=f"x_sec[{i}]")
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

    return opt_ok(result, partial=(_solver_status_1d != "Optimal"))


# Публичный API: оркестратор; реэкспорт на уровень пакета через __init__.py (OPT-008).
from core.optimization.orchestrator import optimize_with_cascading_longitudinal_cuts  # noqa: E402


__all__ = (
    "DEFAULT_CONFIG",
    "GeometryConfig",
    "KERF_WIDTH_MM",
    "LOAD_TO_REINFORCEMENT_MAP",
    "NARROWING_TABLE",
    "OLD_CONFIG",
    "OPT_CASCADING_PLAN",
    "OPT_CASCADING_PLAN_BY_LOAD",
    "OPT_PLAN",
    "OPT_WIDTH_PRIORITY",
    "OptimizationConfig",
    "Piece",
    "Track",
    "apply_width_optimization",
    "build_order_info_list",
    "build_two_d_cutting_ilp",
    "filter_secondary_cut_options_2d",
    "first_fit_decreasing",
    "generate_primary_cut_options_1d",
    "generate_primary_cut_options_2d",
    "generate_raw_secondary_cut_options_2d",
    "generate_secondary_cut_options_1d",
    "optimize_cuts_pulp",
    "optimize_tracks",
    "optimize_with_cascading_longitudinal_cuts",
    "pack_tracks",
    "verify_coverage",
    "_append_actions",
    "_build_proportional_slot_lists",
    "_build_residual_balance_constraints",
    "_get_next_order_info",
    "_group_plate_lengths",
    "_next_slot_info",
    "_optimize_1d_widths_only",
    "_optimize_2d_with_lengths",
    "_peek_order_info",
    "_residual_phys_key",
)
