"""P4 WP3 A4: grep gate — web/core must not use PEP 562 proxy (config_and_data as cfg)."""

from __future__ import annotations

import re
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]

_PROXY_ALIAS = re.compile(r"\bconfig_and_data\s+as\s+cfg\b")
_PROXY_ATTR = re.compile(
    r"\bcfg\.(PLATES_|PLATE_LOAD|PLATE_EXACT|PLATE_LENGTH|LONGITUDINAL|LENGTH_TRIMS|"
    r"UNUSED_STRIPS|SCRAP_STRIPS|USABLE_STRIPS|WASTE_AREA|PLATE_METADATA|"
    r"PLATE_MAX_REINFORCEMENT|PLATE_NOMENCLATURE|LAST_PARSE)"
)
_SCAN_ROOTS = ("app", "core")
_EXCLUDE = frozenset({"core/config_and_data.py", "core/plate_runtime_state.py"})


def _iter_py_sources(root: str) -> list[Path]:
    base = REPO_ROOT / root
    return sorted(p for p in base.rglob("*.py") if p.is_file())


def _relative(path: Path) -> str:
    return path.relative_to(REPO_ROOT).as_posix()


def test_app_and_core_have_no_config_and_data_cfg_alias() -> None:
    offenders: list[str] = []
    for root in _SCAN_ROOTS:
        for path in _iter_py_sources(root):
            rel = _relative(path)
            if rel in _EXCLUDE:
                continue
            text = path.read_text(encoding="utf-8")
            if _PROXY_ALIAS.search(text):
                offenders.append(rel)
    assert offenders == [], f"config_and_data as cfg in web/core: {offenders}"


def test_app_and_core_have_no_cfg_proxy_attribute_access() -> None:
    offenders: list[str] = []
    for root in _SCAN_ROOTS:
        for path in _iter_py_sources(root):
            rel = _relative(path)
            if rel in _EXCLUDE:
                continue
            text = path.read_text(encoding="utf-8")
            if _PROXY_ATTR.search(text):
                offenders.append(rel)
    assert offenders == [], f"cfg.PLATES_* / cfg.PLATE_* proxy access in web/core: {offenders}"
