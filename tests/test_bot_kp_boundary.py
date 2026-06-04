"""Guard: bot handlers must not import core.kp_db directly (A3 boundary)."""

from __future__ import annotations

from pathlib import Path

import pytest

_HANDLERS_DIR = Path(__file__).resolve().parent.parent / "bot" / "handlers"
_FORBIDDEN = "from core import kp_db"


@pytest.mark.parametrize("handler_file", sorted(_HANDLERS_DIR.glob("*.py")), ids=lambda p: p.name)
def test_handler_does_not_import_core_kp_db_directly(handler_file: Path) -> None:
    text = handler_file.read_text(encoding="utf-8")
    assert _FORBIDDEN not in text, (
        f"{handler_file.name} must use "
        "'from bot.services import kp_persistence as kp_db' instead of core.kp_db"
    )
