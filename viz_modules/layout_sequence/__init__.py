# -*- coding: utf-8 -*-
"""
Пакет построения последовательности раскладки плит.

Не создавайте рядом файл ``viz_modules/layout_sequence.py``: при импорте
``viz_modules.layout_sequence`` Python отдал бы приоритет модулю вместо этого
пакета (дрейф копий логики). См. тест ``test_viz_layout_sequence_is_package_not_top_level_py``.
"""

from __future__ import annotations

from viz_modules.layout_sequence.builder import build_layout_sequence
from viz_modules.layout_sequence.debug_trace import _agent_seq_debug
from viz_modules.layout_sequence.from_plan import _build_sequence_from_plan
from viz_modules.layout_sequence.secondary_ops import (
    merge_atomic_secondaries_by_shared_parent as _merge_atomic_secondaries_by_shared_parent,
)

__all__ = [
    "build_layout_sequence",
    "_build_sequence_from_plan",
    "_agent_seq_debug",
    "_merge_atomic_secondaries_by_shared_parent",
]
