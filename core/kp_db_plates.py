#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Plate lifecycle persistence — re-export shim (A1 decomposition)."""

from core.kp_db_audit import audit_append
from core.kp_db_plates_common import (
    _deduct_kp_plate_qty,
    _fetch_kp_plate_row_by_id,
    _insert_completed_plate,
    _normalize_plate_name,
    _purge_zero_qty_plates,
    _record_plate_completion,
    insert_kp_plate_remainder_row,
)
from core.kp_db_plates_completion import (
    check_and_update_kp_completion,
    move_plates_to_completed,
)
from core.kp_db_plates_planning import (
    mark_plates_as_planned,
    recover_stuck_plates,
    return_lost_plates_to_production,
    return_plan_plates_to_production,
    return_plate_rows_for_plan,
    return_plates_to_production,
)
from core.kp_db_plates_queries import (
    get_all_completed_plates,
    get_all_plates_in_production,
    get_completed_plates_by_day,
    get_completed_plates_for_kp,
    get_completed_plates_stats,
    get_remaining_plates_for_kp,
)

__all__ = [
    "audit_append",
    "_normalize_plate_name",
    "_fetch_kp_plate_row_by_id",
    "_deduct_kp_plate_qty",
    "_insert_completed_plate",
    "_record_plate_completion",
    "_purge_zero_qty_plates",
    "insert_kp_plate_remainder_row",
    "move_plates_to_completed",
    "check_and_update_kp_completion",
    "get_remaining_plates_for_kp",
    "mark_plates_as_planned",
    "return_plates_to_production",
    "return_plate_rows_for_plan",
    "return_plan_plates_to_production",
    "recover_stuck_plates",
    "return_lost_plates_to_production",
    "get_completed_plates_for_kp",
    "get_completed_plates_stats",
    "get_completed_plates_by_day",
    "get_all_plates_in_production",
    "get_all_completed_plates",
]
