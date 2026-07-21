"""Domain errors for the production planning pipeline."""

from __future__ import annotations


class PlanBuildError(RuntimeError):
    """Valid input that cannot be turned into a production plan."""
