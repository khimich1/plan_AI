"""Rest matching domain orchestration (A2 rests slice)."""

from __future__ import annotations

import sqlite3
from typing import Any

from core.domain.rest_matching import classify_match_type, compute_cut_cost
from core.domain.rest_matching_types import RestMatch
from core.kp_db_common import DEFAULT_DB, _connect
from core.kp_db_rests import _fetch_available_rests_candidates


class RestMatchingService:
    """Find warehouse rests that can yield plates of the requested size."""

    @staticmethod
    def find_matching_rests_on_cursor(
        cur: sqlite3.Cursor,
        length_m: float,
        width_mm: int,
        qty_needed: int,
    ) -> list[RestMatch]:
        results: list[RestMatch] = []
        qty_collected = 0

        for row in _fetch_available_rests_candidates(cur, length_m, width_mm):
            if qty_collected >= qty_needed:
                break

            rest_length = float(row["length_m"])
            rest_width = int(row["rest_width_mm"])
            rest_qty = int(row["qty"])
            match_type = classify_match_type(
                rest_length,
                rest_width,
                length_m=length_m,
                width_mm=width_mm,
            )
            cut_cost = compute_cut_cost(match_type, length_m=length_m)
            can_take = min(rest_qty, qty_needed - qty_collected)
            customer = row["customer_name"]
            results.append(
                {
                    "rest_id": int(row["id"]),
                    "rest_length": rest_length,
                    "rest_width_mm": rest_width,
                    "rest_qty_available": rest_qty,
                    "qty_to_use": can_take,
                    "match_type": match_type,
                    "cut_cost": cut_cost,
                    "source_plate_name": row["source_plate_name"],
                    "source_kp_id": int(row["kp_id"]),
                    "source_customer": customer if customer is not None else "неизвестно",
                }
            )
            qty_collected += can_take

        return results

    @staticmethod
    def find_matching_rests(
        length_m: float,
        width_mm: int,
        qty_needed: int,
        db_path: str = DEFAULT_DB,
    ) -> list[dict[str, Any]]:
        conn = _connect(db_path)
        try:
            conn.execute("PRAGMA foreign_keys = ON")
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            return list(
                RestMatchingService.find_matching_rests_on_cursor(
                    cur, length_m, width_mm, qty_needed
                )
            )
        finally:
            conn.close()
