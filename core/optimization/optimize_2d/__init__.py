# -*- coding: utf-8 -*-
"""2D plate optimization split into phases (prep/solve, extraction, post-process).

OPT_* thread-local state is not updated inside this subpackage; see ``with_lengths`` module docstring.
"""

from core.optimization.optimize_2d.extract_cuts import extract_two_d_phase_b
from core.optimization.optimize_2d.prep_solve import run_two_d_phase_a
from core.optimization.optimize_2d.state import TwoDPhaseAState, norm_demand_key

__all__ = [
    "TwoDPhaseAState",
    "extract_two_d_phase_b",
    "norm_demand_key",
    "run_two_d_phase_a",
]
