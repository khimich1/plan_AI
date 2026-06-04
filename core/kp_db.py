#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Facade for commercial offers (KP) persistence in plita.db.

Bounded-context modules:
- core.kp_db_schema — DDL / ensure_schema
- core.kp_db_offers — KP CRUD, XLSX, search
- core.kp_db_managers — managers CRUD + seed
- core.kp_db_plates — plate lifecycle
- core.kp_db_rests — plate rests
- core.kp_db_common — _connect, DEFAULT_DB, audit helper
"""

from __future__ import annotations

from core.kp_db_common import DEFAULT_DB, _audit_append, _connect

from core.kp_db_schema import ensure_schema, init_schema

from core.kp_db_nomenclature import (
    enrich_order_data_with_nomenclature,
    fill_plate_nomenclature_cache,
    lookup_nomenclature_by_plate_name,
)

from core.kp_db_offers import (
    clear_all_kp,
    clear_all_plates_data,
    delete_kp_by_id,
    get_all_kp_by_status,
    get_all_kp_list,
    get_db_stats,
    get_kp_by_id,
    get_kp_completion_percentage,
    get_kp_plates_in_plan_percentage,
    get_kp_total_length,
    get_next_kp_number,
    get_xlsx_file,
    save_kp_to_db,
    save_xlsx_file,
    search_kp_by_customer_name,
    update_kp_discount,
    update_kp_execution_date,
    update_kp_logistics_cost,
    update_kp_status,
)

from core.kp_db_managers import (
    add_manager,
    delete_manager,
    get_all_managers,
    get_manager_by_email,
    get_manager_by_id,
    init_default_managers,
    update_manager,
)

from core.kp_db_plates import (
    _normalize_plate_name,
    check_and_update_kp_completion,
    get_all_completed_plates,
    get_all_plates_in_production,
    get_completed_plates_by_day,
    get_completed_plates_for_kp,
    get_completed_plates_stats,
    get_remaining_plates_for_kp,
    mark_plates_as_planned,
    move_plates_to_completed,
    recover_stuck_plates,
    return_lost_plates_to_production,
    return_plan_plates_to_production,
    return_plate_rows_for_plan,
    return_plates_to_production,
)

from core.kp_db_rests import (
    complete_plate_rest,
    create_plate_rest,
    discard_plate_rest,
    find_matching_rests,
    get_all_plate_rests,
    get_available_rests,
    mark_rest_as_used,
)

__all__ = [
    "DEFAULT_DB",
    "_audit_append",
    "_connect",
    "add_manager",
    "check_and_update_kp_completion",
    "clear_all_kp",
    "clear_all_plates_data",
    "complete_plate_rest",
    "create_plate_rest",
    "delete_kp_by_id",
    "delete_manager",
    "discard_plate_rest",
    "ensure_schema",
    "enrich_order_data_with_nomenclature",
    "fill_plate_nomenclature_cache",
    "find_matching_rests",
    "get_all_completed_plates",
    "get_all_kp_by_status",
    "get_all_kp_list",
    "get_all_managers",
    "get_all_plate_rests",
    "get_all_plates_in_production",
    "get_completed_plates_by_day",
    "get_completed_plates_for_kp",
    "get_completed_plates_stats",
    "get_db_stats",
    "get_kp_by_id",
    "get_kp_completion_percentage",
    "get_kp_plates_in_plan_percentage",
    "get_kp_total_length",
    "get_manager_by_email",
    "get_manager_by_id",
    "get_next_kp_number",
    "get_remaining_plates_for_kp",
    "get_xlsx_file",
    "init_default_managers",
    "init_schema",
    "lookup_nomenclature_by_plate_name",
    "mark_plates_as_planned",
    "mark_rest_as_used",
    "move_plates_to_completed",
    "recover_stuck_plates",
    "return_lost_plates_to_production",
    "return_plan_plates_to_production",
    "return_plate_rows_for_plan",
    "return_plates_to_production",
    "save_kp_to_db",
    "save_xlsx_file",
    "search_kp_by_customer_name",
    "update_kp_discount",
    "update_kp_execution_date",
    "update_kp_logistics_cost",
    "update_kp_status",
    "update_manager",
]
