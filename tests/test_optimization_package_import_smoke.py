"""DIP-001: дымовой тест импорта пакета core.optimization без циклических сбоев.

Проверяем, что все имена из __all__ реально доступны на уровне пакета после рефакторинга
направления зависимостей (orchestrator не тянется из _implementation).
"""
from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.optimization as opt  # noqa: E402


def test_package_all_symbols_resolve() -> None:
    assert hasattr(opt, "__all__")
    assert "optimize_with_cascading_longitudinal_cuts" in opt.__all__
    for name in opt.__all__:
        assert hasattr(opt, name), f"missing public attribute: {name!r}"
