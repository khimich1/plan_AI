"""DTO types for plate day completion (A2)."""

from __future__ import annotations

from typing import TypedDict


class UnmovedPlateInfo(TypedDict):
    kp_id: int
    plate_name: str
    qty: int
    length_m: float
    width_m: float
    load_class: int


class CompletePlatesResult(TypedDict):
    completed_count: int
    unmoved: list[UnmovedPlateInfo]
