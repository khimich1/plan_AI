from __future__ import annotations

import sys
import types

from ._implementation import *
from ._implementation import __all__ as _impl_all

# Публичная точка входа (OPT-007): после загрузки графа подмодулей — без цикла _implementation → orchestrator (DIP-001).
from .orchestrator import optimize_with_cascading_longitudinal_cuts

__all__ = (*_impl_all, "optimize_with_cascading_longitudinal_cuts")


class _OptimizationModule(types.ModuleType):
    """Перенаправляет присваивания OPT_* в потоколокальное хранилище (OPT-005)."""

    _TLS_PUBLIC = frozenset({
        "OPT_PLAN",
        "OPT_CASCADING_PLAN",
        "OPT_CASCADING_PLAN_BY_LOAD",
        "OPT_WIDTH_PRIORITY",
        "LOAD_TO_REINFORCEMENT_MAP",
    })

    def __setattr__(self, name: str, value) -> None:
        if name in self._TLS_PUBLIC:
            from core.optimization import context as _opt_ctx

            _opt_ctx.tls_set(name, value)
            return
        super().__setattr__(name, value)


sys.modules[__name__].__class__ = _OptimizationModule
