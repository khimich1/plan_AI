"""Thin facade over ``core.kp_db`` for Telegram handlers (A3 layer boundary)."""

from __future__ import annotations

from core import kp_db
from core.kp_db_nomenclature import enrich_order_data_with_nomenclature

# Low-level (temporary until queries move to bot services)
DEFAULT_DB = kp_db.DEFAULT_DB
_connect = kp_db._connect
ensure_schema = kp_db.ensure_schema
init_schema = kp_db.init_schema

# Offers / archive / admin
get_kp_by_id = kp_db.get_kp_by_id
get_all_kp_list = kp_db.get_all_kp_list
get_all_kp_by_status = kp_db.get_all_kp_by_status
update_kp_discount = kp_db.update_kp_discount
update_kp_status = kp_db.update_kp_status
update_kp_execution_date = kp_db.update_kp_execution_date
delete_kp_by_id = kp_db.delete_kp_by_id
get_kp_completion_percentage = kp_db.get_kp_completion_percentage
get_kp_plates_in_plan_percentage = kp_db.get_kp_plates_in_plan_percentage
get_kp_total_length = kp_db.get_kp_total_length
get_next_kp_number = kp_db.get_next_kp_number
get_manager_by_id = kp_db.get_manager_by_id
clear_all_kp = kp_db.clear_all_kp
clear_all_plates_data = kp_db.clear_all_plates_data
get_db_stats = kp_db.get_db_stats

# Plates / production / rests
move_plates_to_completed = kp_db.move_plates_to_completed
check_and_update_kp_completion = kp_db.check_and_update_kp_completion
mark_rest_as_used = kp_db.mark_rest_as_used
create_plate_rest = kp_db.create_plate_rest
return_plan_plates_to_production = kp_db.return_plan_plates_to_production
get_all_plates_in_production = kp_db.get_all_plates_in_production
get_all_completed_plates = kp_db.get_all_completed_plates
mark_plates_as_planned = kp_db.mark_plates_as_planned
recover_stuck_plates = kp_db.recover_stuck_plates
get_all_plate_rests = kp_db.get_all_plate_rests

__all__ = [
    "DEFAULT_DB",
    "_connect",
    "check_and_update_kp_completion",
    "clear_all_kp",
    "clear_all_plates_data",
    "create_plate_rest",
    "delete_kp_by_id",
    "ensure_schema",
    "enrich_order_data_with_nomenclature",
    "get_all_completed_plates",
    "get_all_kp_by_status",
    "get_all_kp_list",
    "get_all_plate_rests",
    "get_all_plates_in_production",
    "get_db_stats",
    "get_kp_by_id",
    "get_kp_completion_percentage",
    "get_kp_plates_in_plan_percentage",
    "get_kp_total_length",
    "get_manager_by_id",
    "get_next_kp_number",
    "init_schema",
    "mark_plates_as_planned",
    "mark_rest_as_used",
    "move_plates_to_completed",
    "recover_stuck_plates",
    "return_plan_plates_to_production",
    "update_kp_discount",
    "update_kp_execution_date",
    "update_kp_status",
]
