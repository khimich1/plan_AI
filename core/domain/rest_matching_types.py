"""Types for plate rest matching (A2 rests slice)."""

from __future__ import annotations

from typing import Literal, TypedDict

MatchType = Literal["exact", "width_cut", "length_cut", "both_cuts"]


class RestMatch(TypedDict):
    rest_id: int
    rest_length: float
    rest_width_mm: int
    rest_qty_available: int
    qty_to_use: int
    match_type: MatchType
    cut_cost: float
    source_plate_name: str
    source_kp_id: int
    source_customer: str
