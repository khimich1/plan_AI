"""Guard: app layer must not import monolithic core.kp_db (A3 wave 2)."""

from __future__ import annotations

import re
from pathlib import Path

_FORBIDDEN = re.compile(r"^\s*from\s+core\s+import\s+kp_db\s*$", re.MULTILINE)
_TARGETS = [
    Path(__file__).resolve().parents[1] / "app" / "repositories" / "kp_repository.py",
    Path(__file__).resolve().parents[1] / "app" / "repositories" / "kp_archive_repository.py",
    Path(__file__).resolve().parents[1] / "app" / "services" / "admin_service.py",
]


def test_app_kp_modules_do_not_import_core_kp_db_monolith() -> None:
    offenders: list[str] = []
    for path in _TARGETS:
        text = path.read_text(encoding="utf-8")
        if _FORBIDDEN.search(text):
            offenders.append(str(path.name))
    assert not offenders, (
        "Use core.kp_db_offers / kp_db_plates slice modules instead of "
        f"'from core import kp_db': {offenders}"
    )
