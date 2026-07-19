#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Явный контракт результата оптимизации (успех / ошибка / частичный план).

Служебные ключи с префиксом ``_opt_`` не смешиваются с полями плана
(primary_cuts, plate_assignments, …) и стабильно читаются API/UI.
"""

from __future__ import annotations

from typing import Any, Final, Literal

OPT_STATUS_KEY: Final = "_opt_status"
OPT_ERROR_CODE_KEY: Final = "_opt_error_code"
OPT_ERROR_MESSAGE_KEY: Final = "_opt_error_message"
OPT_SOLVER_STATUS_KEY: Final = "_opt_solver_status"
OPT_DETAIL_KEY: Final = "_opt_detail"

OptimizationResultStatus = Literal["ok", "error", "partial"]

# Коды для клиента / логов (стабильные строки)
ERROR_PULP_MISSING: Final = "pulp_missing"
ERROR_EMPTY_ORDERS_2D: Final = "empty_orders_2d"
ERROR_EMPTY_ORDERS_1D: Final = "empty_orders_1d"
ERROR_NO_INPUT: Final = "no_input"
ERROR_SOLVER_INFEASIBLE: Final = "solver_infeasible"
ERROR_SOLVER_UNDEFINED: Final = "solver_undefined"


def opt_error(
    code: str,
    message: str | None = None,
    *,
    solver_status: str | None = None,
    detail: dict[str, Any] | None = None,
) -> dict[str, Any]:
    out: dict[str, Any] = {
        OPT_STATUS_KEY: "error",
        OPT_ERROR_CODE_KEY: code,
        OPT_ERROR_MESSAGE_KEY: message or code,
    }
    if solver_status is not None:
        out[OPT_SOLVER_STATUS_KEY] = solver_status
    if detail:
        out[OPT_DETAIL_KEY] = detail
    return out


def opt_ok(payload: dict[str, Any], *, partial: bool = False) -> dict[str, Any]:
    merged = dict(payload)
    merged[OPT_STATUS_KEY] = "partial" if partial else "ok"
    return merged


def is_optimization_success(d: dict[str, Any] | None) -> bool:
    """План пригоден для downstream (дорожки, визуализация, КП)."""
    if not d:
        return False
    status = d.get(OPT_STATUS_KEY)
    if status == "error":
        return False
    if status in ("ok", "partial"):
        return True
    # Обратная совместимость: старые ответы без маркера
    return bool(
        d.get("total_plates")
        or d.get("primary_cuts")
        or d.get("secondary_cuts")
        or d.get("plate_assignments")
        or d.get("actions")
    )


def optimization_error_code(d: dict[str, Any] | None) -> str | None:
    if not d or d.get(OPT_STATUS_KEY) != "error":
        return None
    raw = d.get(OPT_ERROR_CODE_KEY)
    return str(raw) if raw is not None else None


def optimization_error_message(d: dict[str, Any] | None) -> str | None:
    if not d or d.get(OPT_STATUS_KEY) != "error":
        return None
    msg = d.get(OPT_ERROR_MESSAGE_KEY)
    return str(msg) if msg is not None else None
