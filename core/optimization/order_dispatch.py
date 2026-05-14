#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Order aggregation, canonical keys, and proportional slot attribution (no PuLP)."""

from __future__ import annotations

from typing import Any, Mapping

from core.config_and_data import canonical_plate_key
from core.optimization.ports.order_data import PlateOrderDataPort

from core.optimization.debug_log import (
    _DEBUG_LOG_5b5324,
    _DEBUG_LOG_COMMON,
    _dbg_open_append,
    _opt_debug_enabled,
)


def build_order_info_list(
    orders_2d: list[Mapping[str, Any]],
    order_data: PlateOrderDataPort,
) -> dict[tuple[Any, ...], list[dict[str, Any]]]:
    """
    Маппинг (length, width, load_code) → список записей КП с qty_remaining.
    Ключи в формате canonical_plate_key; load_code через order_data.normalize_load_code.
    """
    order_info_list: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
    for order in orders_2d:
        load_code = order_data.normalize_load_code(order.get("load_code", 800))
        key = canonical_plate_key(order["length"], order["width"], load_code)
        if key not in order_info_list:
            order_info_list[key] = []
        order_info_list[key].append({
            "kp_id": order.get("kp_id"),
            "customer": order.get("customer", "неизвестно"),
            "kp_date": order.get("kp_date", "неизвестно"),
            "plate_name": order.get("plate_name", ""),
            "load_code": load_code,
            "reinforcement": order.get("reinforcement", 0),
            "qty_remaining": order.get("qty", 1),
            "concrete_grade": order.get("concrete_grade"),
        })
    return order_info_list


def _get_next_order_info(
    order_info_list: dict[tuple[Any, ...], list[dict[str, Any]]],
    key: tuple[Any, ...],
) -> dict[str, Any]:
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
    if _opt_debug_enabled():
        try:
            with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
                _f.write(__import__("json").dumps({"sessionId": "5b5324", "hypothesisId": "H_get_next", "location": "optimization:_get_next_order_info:entry", "message": "key requested", "data": {"key": list(key) if isinstance(key, tuple) else key}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
    # #endregion
    entries = order_info_list.get(key, [])
    for entry in entries:
        if entry.get("qty_remaining", 0) > 0:
            entry["qty_remaining"] -= 1
            out = {
                "kp_id": entry.get("kp_id"),
                "customer": entry.get("customer"),
                "kp_date": entry.get("kp_date"),
                "plate_name": entry.get("plate_name"),
                "load_code": entry.get("load_code"),
                "reinforcement": entry.get("reinforcement"),
                "concrete_grade": entry.get("concrete_grade"),
                "identity_match_type": "exact",
            }
            # #region agent log (session 5b5324) _get_next_order_info return exact
            if _opt_debug_enabled():
                try:
                    with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
                        _f.write(__import__("json").dumps({"sessionId": "5b5324", "hypothesisId": "H_get_next", "location": "optimization:_get_next_order_info:return", "message": "exact match", "data": {"match_type": "exact", "kp_id": out.get("kp_id"), "plate_name": (out.get("plate_name") or "")[:60]}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
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
                    if entry.get("qty_remaining", 0) > 0:
                        entry["qty_remaining"] -= 1
                        if _opt_debug_enabled():
                            try:
                                _req_lc = key[2] if len(key) >= 3 else None
                                _found_lc = candidate_key[2] if len(candidate_key) >= 3 else None
                                with _dbg_open_append(_DEBUG_LOG_COMMON) as _f:
                                    _f.write(__import__("json").dumps({
                                        "hypothesisId": "H2_fallback",
                                        "location": "optimization.py:_get_next_order_info",
                                        "message": "fallback used (length, width)",
                                        "data": {
                                            "requested_key": list(key),
                                            "found_key": list(candidate_key),
                                            "fallback_reason": "load_code_mismatch",
                                            "requested_load_code": _req_lc,
                                            "found_load_code": _found_lc,
                                            "kp_id": entry.get("kp_id"),
                                            "plate_name": (entry.get("plate_name") or "")[:50],
                                        },
                                        "timestamp": __import__("time").time(),
                                    }, ensure_ascii=False) + "\n")
                            except Exception:
                                pass
                        out_fb = {
                            "kp_id": entry.get("kp_id"),
                            "customer": entry.get("customer"),
                            "kp_date": entry.get("kp_date"),
                            "plate_name": entry.get("plate_name"),
                            "load_code": entry.get("load_code"),
                            "reinforcement": entry.get("reinforcement"),
                            "concrete_grade": entry.get("concrete_grade"),
                            "identity_match_type": "fallback_same_length_width",
                        }
                        # #region agent log (session 5b5324) fallback_same_length_width
                        if _opt_debug_enabled():
                            try:
                                with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
                                    _f.write(__import__("json").dumps({"sessionId": "5b5324", "hypothesisId": "H_get_next", "location": "optimization:_get_next_order_info:return", "message": "fallback_same_length_width", "data": {"requested_key": list(key), "found_key": list(candidate_key), "kp_id": out_fb.get("kp_id"), "plate_name": (out_fb.get("plate_name") or "")[:60]}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
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
                    if entry.get("qty_remaining", 0) > 0:
                        entry["qty_remaining"] -= 1
                        out_n = {
                            "kp_id": entry.get("kp_id"),
                            "customer": entry.get("customer"),
                            "kp_date": entry.get("kp_date"),
                            "plate_name": entry.get("plate_name"),
                            "load_code": entry.get("load_code"),
                            "reinforcement": entry.get("reinforcement"),
                            "concrete_grade": entry.get("concrete_grade"),
                            "identity_match_type": "fallback_neighbor_length",
                        }
                        # #region agent log (session 5b5324) fallback_neighbor_length
                        if _opt_debug_enabled():
                            try:
                                with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
                                    _f.write(__import__("json").dumps({"sessionId": "5b5324", "hypothesisId": "H_get_next", "location": "optimization:_get_next_order_info:return", "message": "fallback_neighbor_length", "data": {"requested_key": list(key), "found_key": list(candidate_key), "kp_id": out_n.get("kp_id"), "plate_name": (out_n.get("plate_name") or "")[:60]}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
                            except Exception:
                                pass
                        # #endregion
                        return out_n
    # #region agent log (session 5b5324) _get_next_order_info return empty
    if _opt_debug_enabled():
        try:
            with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
                _f.write(__import__("json").dumps({"sessionId": "5b5324", "hypothesisId": "H_get_next", "location": "optimization:_get_next_order_info:return", "message": "empty", "data": {"key": list(key) if isinstance(key, tuple) else key}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
    # #endregion
    return {}


def _build_proportional_slot_lists(
    orders_2d: list[Mapping[str, Any]],
    demand_2d: dict[tuple[Any, ...], int],
) -> tuple[dict[tuple[Any, ...], list[dict[str, Any]]], dict[tuple[Any, ...], int]]:
    """
    Строит пропорциональные слоты атрибуции по ключу (length, width, load_code).
    Возвращает (slot_lists, slot_cursors).
    slot_lists[key] — список из demand_2d[key] атрибуций, пропорционально qty заказов
    (floor + остаток по убыванию qty). Курсоры инициализированы в 0.
    """
    groups: dict[tuple[Any, Any, Any], list[Mapping[str, Any]]] = {}
    for order in orders_2d:
        key = (order["length"], order["width"], order.get("load_code", 800))
        groups.setdefault(key, []).append(order)

    slot_lists: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
    for key, need in demand_2d.items():
        entries = groups.get(key, [])
        total_qty = sum(o.get("qty", 1) for o in entries)
        if total_qty == 0:
            slot_lists[key] = []
            continue
        shares = [int(need * o.get("qty", 1) / total_qty) for o in entries]
        remainder = need - sum(shares)
        for idx in sorted(range(len(entries)), key=lambda i: -entries[i].get("qty", 1)):
            if remainder <= 0:
                break
            shares[idx] += 1
            remainder -= 1
        slots: list[dict[str, Any]] = []
        for entry, share in zip(entries, shares):
            info = {
                k: entry.get(k)
                for k in ("kp_id", "customer", "kp_date", "plate_name", "load_code", "reinforcement", "concrete_grade")
            }
            slots.extend([info] * share)
        slot_lists[key] = slots

    cursors = {key: 0 for key in slot_lists}
    return slot_lists, cursors


def _next_slot_info(
    slot_lists: dict[tuple[Any, ...], list[dict[str, Any]]],
    slot_cursors: dict[tuple[Any, ...], int],
    key: tuple[Any, ...],
) -> dict[str, Any]:
    """
    Возвращает следующую атрибуцию по ключу из предрасчитанных слотов и сдвигает курсор.
    При исчерпании слотов возвращает пустой dict, чтобы не дублировать identity.
    """
    slots = slot_lists.get(key, [])
    idx = slot_cursors.get(key, 0)
    if not slots or idx >= len(slots):
        return {}
    entry = dict(slots[idx])
    entry["identity_match_type"] = "slot_proportional"
    slot_cursors[key] = idx + 1
    return entry


def _peek_order_info(
    order_info_list: dict[tuple[Any, ...], list[dict[str, Any]]],
    key: tuple[Any, ...],
) -> dict[str, Any]:
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
        if entry.get("qty_remaining", 0) > 0:
            return {
                "kp_id": entry.get("kp_id"),
                "customer": entry.get("customer"),
                "kp_date": entry.get("kp_date"),
                "plate_name": entry.get("plate_name"),
                "load_code": entry.get("load_code"),
                "reinforcement": entry.get("reinforcement"),
                "concrete_grade": entry.get("concrete_grade"),
            }
    # Fallback по (length, width) без load_code
    if len(key) == 3:
        length, width, _ = key
        for candidate_key, candidate_entries in order_info_list.items():
            if len(candidate_key) >= 2 and candidate_key[0] == length and candidate_key[1] == width:
                for entry in candidate_entries:
                    if entry.get("qty_remaining", 0) > 0:
                        return {
                            "kp_id": entry.get("kp_id"),
                            "customer": entry.get("customer"),
                            "kp_date": entry.get("kp_date"),
                            "plate_name": entry.get("plate_name"),
                            "load_code": entry.get("load_code"),
                            "reinforcement": entry.get("reinforcement"),
                            "concrete_grade": entry.get("concrete_grade"),
                        }
    return {}
