"""Guard: no JSON plan file I/O on app/planning hot paths (WP1 / A11)."""
from __future__ import annotations

import re
from pathlib import Path

import pytest

_PROJECT_ROOT = Path(__file__).resolve().parents[1]
_PLANNING_DIR = _PROJECT_ROOT / "app" / "planning"
_PLAN_REPOSITORY = _PROJECT_ROOT / "app" / "repositories" / "plan_repository.py"

# Relative to app/planning/ — only migration helpers belong here, not runtime hot paths.
_ALLOWLIST: frozenset[str] = frozenset()

_FORBIDDEN_IN_PLANNING: tuple[tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"\bjson\.dump\s*\("), "json.dump("),
    (re.compile(r"\bopen\s*\([^)]*plans"), "open(...plans"),
    (re.compile(r"\.write_text\s*\("), ".write_text("),
)

_IMPORT_FROM_PLAN_STORAGE = re.compile(
    r"^\s*from\s+app\.planning\.plan_storage\s+import\s+(.+)$",
    re.MULTILINE,
)

_PERSISTENCE_IMPORTS = frozenset(
    {"save_plan", "load_plan", "save_plans_metadata", "delete_plan"},
)


def _iter_planning_modules() -> list[tuple[Path, str]]:
    modules: list[tuple[Path, str]] = []
    for path in sorted(_PLANNING_DIR.rglob("*.py")):
        rel = path.relative_to(_PLANNING_DIR).as_posix()
        if rel in _ALLOWLIST:
            continue
        modules.append((path, rel))
    return modules


_PLANNING_MODULES = _iter_planning_modules()


@pytest.mark.parametrize(
    ("module_path", "rel_path"),
    _PLANNING_MODULES,
    ids=[rel for _, rel in _PLANNING_MODULES],
)
def test_planning_modules_avoid_plan_file_io(module_path: Path, rel_path: str) -> None:
    text = module_path.read_text(encoding="utf-8")
    offenders: list[str] = []
    for pattern, label in _FORBIDDEN_IN_PLANNING:
        if pattern.search(text):
            offenders.append(label)
    assert not offenders, f"app/planning/{rel_path} must not use plan file I/O: {offenders}"


def test_plan_repository_does_not_import_plan_storage_persistence() -> None:
    """PlanRepository may import pure helpers from plan_storage, not persistence API."""
    text = _PLAN_REPOSITORY.read_text(encoding="utf-8")
    match = _IMPORT_FROM_PLAN_STORAGE.search(text)
    if match is None:
        return
    imported = match.group(1)
    offenders = sorted(sym for sym in _PERSISTENCE_IMPORTS if sym in imported)
    assert not offenders, (
        f"plan_repository.py must not import plan_storage persistence {offenders}; "
        "use SQLite methods directly or core/production/plan_utils.py"
    )
