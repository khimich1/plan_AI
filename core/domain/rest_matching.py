"""Domain rules for matching plate rests to required plate dimensions (A2)."""

from __future__ import annotations

from core.domain.rest_matching_types import MatchType

# TODO: move to Settings when config slice is ready (audit A20)
LONG_CUT_PRICE_PER_M = 460.0
TRANSVERSE_CUT_PRICE = 1200.0

_LENGTH_TOLERANCE_M = 0.01


def classify_match_type(
    rest_length: float,
    rest_width_mm: int,
    *,
    length_m: float,
    width_mm: int,
) -> MatchType:
    length_match = abs(rest_length - length_m) < _LENGTH_TOLERANCE_M
    width_match = rest_width_mm == width_mm
    if length_match and width_match:
        return "exact"
    if length_match and not width_match:
        return "width_cut"
    if not length_match and width_match:
        return "length_cut"
    return "both_cuts"


def compute_cut_cost(match_type: MatchType, *, length_m: float) -> float:
    if match_type == "exact":
        return 0.0
    if match_type == "width_cut":
        return LONG_CUT_PRICE_PER_M * length_m
    if match_type == "length_cut":
        return TRANSVERSE_CUT_PRICE
    return LONG_CUT_PRICE_PER_M * length_m + TRANSVERSE_CUT_PRICE
