from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Protocol, Union


class PriceLookupFn(Protocol):
    def __call__(
        self, length_m: float, load_code: float | int = 8, db_path: str = ...
    ) -> Optional[float]: ...


class RawMaterialCostFn(Protocol):
    def __call__(self, plate_name: str, db_path: str = ...) -> Optional[float]: ...


class ReinforcementFn(Protocol):
    def __call__(
        self,
        length_m: float,
        load_code: int | float,
        source: str = 'erm',
        db_path: Union[str, Path] = ...,
        allow_fallback: bool = True,
    ) -> float | None: ...


@dataclass(frozen=True)
class ProcurementDeps:
    db_path: str
    get_price: PriceLookupFn
    get_raw_material_cost: RawMaterialCostFn
    get_reinforcement: ReinforcementFn


def resolve_procurement_deps(deps: ProcurementDeps | None) -> ProcurementDeps:
    if deps is not None:
        return deps
    from .adapters_default import default_procurement_deps

    return default_procurement_deps()
