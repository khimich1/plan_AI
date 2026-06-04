"""Guard: production services must not import monolithic core.kp_db (A3)."""

from __future__ import annotations

import re
from pathlib import Path

_FORBIDDEN = re.compile(r"^\s*from\s+core\s+import\s+kp_db\s*$", re.MULTILINE)
_ROOT = Path(__file__).resolve().parents[1] / "app" / "services"


def test_production_services_do_not_import_core_kp_db_monolith() -> None:
    offenders: list[str] = []
    for path in sorted(_ROOT.glob("production_*.py")):
        text = path.read_text(encoding="utf-8")
        if _FORBIDDEN.search(text):
            offenders.append(str(path.relative_to(_ROOT.parents[1])))
    assert not offenders, (
        "Use core.kp_db_plates / kp_db_schema / slice modules instead of "
        f"'{_FORBIDDEN}': {offenders}"
    )
