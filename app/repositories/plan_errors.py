from __future__ import annotations


class PlanVersionConflict(Exception):
    """Optimistic lock failure: expected plan version does not match stored version."""

    def __init__(self, plan_id: str, expected_version: int) -> None:
        self.plan_id = plan_id
        self.expected_version = expected_version
        super().__init__(
            f"Plan {plan_id!r} version conflict (expected version {expected_version})"
        )
