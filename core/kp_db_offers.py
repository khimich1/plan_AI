#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Commercial offers (KP) persistence — thin re-export facade (A5 slice 5)."""

from __future__ import annotations

from core.kp.offers_read import (
    get_all_kp_by_status,
    get_all_kp_list,
    get_db_stats,
    get_kp_by_id,
    get_kp_completion_percentage,
    get_kp_plates_in_plan_percentage,
    get_kp_total_length,
    get_next_kp_number,
    get_xlsx_file,
    search_kp_by_customer_name,
)
from core.kp.offers_write import (
    clear_all_kp,
    clear_all_plates_data,
    delete_kp_by_id,
    save_kp_to_db,
    save_xlsx_file,
    update_kp_discount,
    update_kp_execution_date,
    update_kp_logistics_cost,
    update_kp_status,
)

__all__ = [
    "clear_all_kp",
    "clear_all_plates_data",
    "delete_kp_by_id",
    "get_all_kp_by_status",
    "get_all_kp_list",
    "get_db_stats",
    "get_kp_by_id",
    "get_kp_completion_percentage",
    "get_kp_plates_in_plan_percentage",
    "get_kp_total_length",
    "get_next_kp_number",
    "get_xlsx_file",
    "save_kp_to_db",
    "save_xlsx_file",
    "search_kp_by_customer_name",
    "update_kp_discount",
    "update_kp_execution_date",
    "update_kp_logistics_cost",
    "update_kp_status",
]
