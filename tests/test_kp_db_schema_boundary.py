"""Guard: schema init only in kp_db_schema + documented entrypoints (A4)."""

from __future__ import annotations

from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent / "core"

_SLICES = (
    "kp_db_offers.py",
    "kp_db_managers.py",
    "kp_db_plates.py",
    "kp_db_rests.py",
)

_FORBIDDEN_PATTERNS = ("ensure_schema(", "_init_schema(")


@pytest.mark.parametrize("module_name", _SLICES)
def test_persistence_slices_do_not_call_schema_init(module_name: str) -> None:
    text = (_ROOT / module_name).read_text(encoding="utf-8")
    for pattern in _FORBIDDEN_PATTERNS:
        assert pattern not in text, (
            f"{module_name} must not call {pattern!r}; "
            "use startup ensure_schema or test fixtures (make_iso_db)."
        )


def test_schema_module_defines_ensure_schema() -> None:
    text = (_ROOT / "kp_db_schema.py").read_text(encoding="utf-8")
    assert "def ensure_schema(" in text
    assert "def _init_schema_impl(" in text
