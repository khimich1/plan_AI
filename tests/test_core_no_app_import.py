"""Guard: core package must not import app (layer boundary)."""

from __future__ import annotations

import re
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1] / "core"
_FROM_APP = re.compile(r"^\s*from\s+app(?:\.|\s)", re.MULTILINE)
_IMPORT_APP = re.compile(r"^\s*import\s+app(?:\.|\s)", re.MULTILINE)


def test_core_py_files_do_not_import_app() -> None:
    offenders: list[str] = []
    for path in sorted(_ROOT.rglob("*.py")):
        if path.name == "__pycache__":
            continue
        text = path.read_text(encoding="utf-8")
        if _FROM_APP.search(text) or _IMPORT_APP.search(text):
            offenders.append(str(path.relative_to(_ROOT.parents[0])))
    assert not offenders, f"core must not depend on app: {offenders}"
