from __future__ import annotations

from core.execution_terms import normalize_execution_terms_to_ddmmyyyy


class ExecutionTermsService:
    def normalize(self, raw_input: str) -> str:
        return normalize_execution_terms_to_ddmmyyyy(raw_input)
