from __future__ import annotations

from collections.abc import Callable

from .types import NomenclatureCacheFiller

_filler: NomenclatureCacheFiller | None = None


def _lazy_default_nomenclature_cache_fill() -> None:
    from core.kp_db import fill_plate_nomenclature_cache

    fill_plate_nomenclature_cache()


def get_default_nomenclature_cache_filler() -> NomenclatureCacheFiller:
    if _filler is not None:
        return _filler
    return _lazy_default_nomenclature_cache_fill


def set_nomenclature_cache_filler(f: Callable[[], None] | None) -> None:
    global _filler
    _filler = f
