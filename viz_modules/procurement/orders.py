from __future__ import annotations

from collections.abc import Mapping
from typing import TYPE_CHECKING, Any

from .plan_snapshot import OrderRequestedRow, parse_cascading_plan

if TYPE_CHECKING:
    from core.optimization.layout_runtime_snapshot import OptPlanFrozenSnapshot


def _orders_payload(rows: list[OrderRequestedRow]) -> list[dict[str, Any]]:
    orders_copy: list[dict[str, Any]] = []
    for order in rows:
        orders_copy.append(
            {
                "length": float(order.length),
                "width": order.width,
                "qty": int(order.qty),
                "load_code": order.load_code,
                "length_dm_raw": order.length_dm_raw,
            }
        )
    return orders_copy


def _orders_from_validated_maps(
    cascading_plan_by_load: Mapping[Any, Mapping[str, Any] | None] | None,
    cascading_plan_single: Mapping[str, Any] | None,
) -> list[dict[str, Any]] | None:
    orders_copy: list[dict[str, Any]] = []

    if cascading_plan_by_load:
        for _load_key, plan_raw in cascading_plan_by_load.items():
            snap = parse_cascading_plan(plan_raw if isinstance(plan_raw, Mapping) else {})
            if snap.orders_requested:
                orders_copy.extend(_orders_payload(snap.orders_requested))

    if orders_copy:
        return orders_copy

    snap_single = parse_cascading_plan(cascading_plan_single if isinstance(cascading_plan_single, Mapping) else None)
    if snap_single.orders_requested:
        return _orders_payload(snap_single.orders_requested)
    return None


def get_orders_from_opt_plan(opt_snapshot: OptPlanFrozenSnapshot | None = None):
    """
    Возвращает заказы (length/width/qty) из плана оптимизатора.

    Приоритет:
    1. Карта ``OPT_CASCADING_PLAN_BY_LOAD`` (или её копия из ``opt_snapshot``), после валидации снимка.
    2. Иначе общий ``OPT_CASCADING_PLAN`` / ``opt_snapshot.opt_cascading_plan``.
    3. Если заказов нет — ``None``.

    Явный ``opt_snapshot`` убирает чтение TLS-глобалов и задаёт контракт для вызывающего кода ([A7], [S4]).
    """
    if opt_snapshot is not None:
        return _orders_from_validated_maps(opt_snapshot.opt_cascading_plan_by_load, opt_snapshot.opt_cascading_plan)

    try:
        from core.optimization import OPT_CASCADING_PLAN, OPT_CASCADING_PLAN_BY_LOAD
    except ImportError:
        return None

    return _orders_from_validated_maps(OPT_CASCADING_PLAN_BY_LOAD, OPT_CASCADING_PLAN)
