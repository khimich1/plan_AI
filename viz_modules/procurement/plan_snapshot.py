"""Явный контракт снимка каскадного плана оптимизатора для procurement.

[S4] / [Q3]: один слой нормализации и проверки схемы на границе между OPT JSON и расчётом
закупок; ниже по стеку используются согласованные словари из `.model_dump()`.

Не заменяет `OptPlanFrozenSnapshot` (глублёкая заморозка TLS в layout): там материализуйте
`layout_runtime_snapshot.OptPlanFrozenSnapshot`, затем передавайте его карты в
`get_orders_from_opt_plan(opt_snapshot=...)` чтобы избежать неявного «текущего» глобала.
"""

from __future__ import annotations

from collections.abc import Mapping, MutableMapping
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, ValidationError, field_validator

__all__ = [
    "CascadingPlanSnapshot",
    "OrderRequestedRow",
    "PlanSnapshotValidationError",
    "PrimaryCutSnapshot",
    "SecondaryCutSnapshot",
    "normalize_optimizer_plan_mapping",
    "parse_cascading_plan",
    "parse_plan_by_load",
    "snapshot_to_trim_dict",
]


class PlanSnapshotValidationError(ValueError):
    """Обёртка над pydantic ValidationError для публичного API procurement."""

    def __init__(self, message: str, *, errors: list[Any] | None = None) -> None:
        super().__init__(message)
        self.errors = errors or []


class OrderRequestedRow(BaseModel):
    model_config = ConfigDict(extra="ignore")

    length: float = 0.0
    width: float | int = 0
    qty: int = Field(default=1, ge=0)
    load_code: float | int | None = None
    length_dm_raw: str = ""

    @field_validator("length", mode="before")
    @classmethod
    def _coerce_length(cls, v: Any) -> float:
        if v is None:
            return 0.0
        return float(v)

    @field_validator("width", mode="before")
    @classmethod
    def _coerce_width(cls, v: Any) -> float | int:
        if v is None:
            return 0
        if isinstance(v, bool):  # bool is subclass of int
            return int(v)
        if isinstance(v, int):
            return v
        return float(v)

    @field_validator("qty", mode="before")
    @classmethod
    def _coerce_qty(cls, v: Any) -> int:
        if v is None:
            return 1
        return int(v)

    @field_validator("load_code", mode="before")
    @classmethod
    def _coerce_load(cls, v: Any) -> float | int | None:
        if v is None or v == "":
            return None
        if isinstance(v, bool):
            return int(v)
        if isinstance(v, int):
            return v
        return float(v)

    @field_validator("length_dm_raw", mode="before")
    @classmethod
    def _strip_ldr(cls, v: Any) -> str:
        if v is None:
            return ""
        return str(v).strip()


class PrimaryCutSnapshot(BaseModel):
    model_config = ConfigDict(extra="ignore")

    width: int | None = None
    lengths: list[float] = Field(default_factory=list)
    qty: int = 0
    rest: int = 0
    load_code: float | int | None = None
    primary_instance_id: str | None = None

    @field_validator("lengths", mode="before")
    @classmethod
    def _lengths(cls, v: Any) -> list[float]:
        if v is None:
            return []
        if not isinstance(v, (list, tuple)):
            return [float(v)]
        out: list[float] = []
        for x in v:
            try:
                if x is not None:
                    out.append(float(x))
            except (TypeError, ValueError):
                continue
        return out

    @field_validator("width", mode="before")
    @classmethod
    def _width(cls, v: Any) -> int | None:
        if v is None:
            return None
        try:
            return int(v)
        except (TypeError, ValueError):
            return None

    @field_validator("qty", "rest", mode="before")
    @classmethod
    def _qty_rest(cls, v: Any) -> int:
        if v is None:
            return 0
        try:
            return int(v)
        except (TypeError, ValueError):
            return 0

    @field_validator("load_code", mode="before")
    @classmethod
    def _coerce_load(cls, v: Any) -> float | int | None:
        if v is None or v == "":
            return None
        if isinstance(v, bool):
            return int(v)
        if isinstance(v, int):
            return v
        return float(v)


class SecondaryCutSnapshot(BaseModel):
    model_config = ConfigDict(extra="ignore")

    source: int = 0
    source_lengths: list[float] = Field(default_factory=list)
    lengths: list[float] = Field(default_factory=list)
    cuts: list[float | int] = Field(default_factory=list)
    qty: int = 0
    pieces: int = 1
    waste: float = 0.0
    type: str | None = None
    load_code: float | int | None = None
    parent_instance_id: str | None = None

    @field_validator("source_lengths", "lengths", mode="before")
    @classmethod
    def _float_list(cls, v: Any) -> list[float]:
        if v is None:
            return []
        if not isinstance(v, (list, tuple)):
            try:
                return [float(v)]
            except (TypeError, ValueError):
                return []
        out: list[float] = []
        for x in v:
            try:
                if x is not None:
                    out.append(float(x))
            except (TypeError, ValueError):
                continue
        return out

    @field_validator("cuts", mode="before")
    @classmethod
    def _cuts(cls, v: Any) -> list[float | int]:
        if v is None:
            return []
        if not isinstance(v, (list, tuple)):
            try:
                f = float(v)
                return [int(f) if f == int(f) else f]
            except (TypeError, ValueError):
                return []
        out: list[float | int] = []
        for x in v:
            try:
                if x is None:
                    continue
                if isinstance(x, bool):
                    out.append(int(x))
                    continue
                if isinstance(x, int):
                    out.append(x)
                    continue
                f = float(x)
                out.append(int(f) if f == int(f) else f)
            except (TypeError, ValueError):
                continue
        return out

    @field_validator("source", "qty", "pieces", mode="before")
    @classmethod
    def _ints(cls, v: Any) -> int:
        if v is None:
            return 0
        try:
            return int(v)
        except (TypeError, ValueError):
            return 0

    @field_validator("waste", mode="before")
    @classmethod
    def _waste(cls, v: Any) -> float:
        if v is None:
            return 0.0
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    @field_validator("load_code", mode="before")
    @classmethod
    def _coerce_load(cls, v: Any) -> float | int | None:
        if v is None or v == "":
            return None
        if isinstance(v, bool):
            return int(v)
        if isinstance(v, int):
            return v
        return float(v)


class CascadingPlanSnapshot(BaseModel):
    """Допустимое подмножество ключей OPT-плана, используемое procurement."""

    model_config = ConfigDict(extra="ignore")

    orders_requested: list[OrderRequestedRow] = Field(default_factory=list)
    primary_cuts: list[PrimaryCutSnapshot] = Field(default_factory=list)
    secondary_cuts: list[SecondaryCutSnapshot] = Field(default_factory=list)
    transverse_cuts: list[Any] = Field(default_factory=list)
    total_plates: int | None = None
    waste_width: float | int | None = None

    @field_validator("orders_requested", "primary_cuts", "secondary_cuts", "transverse_cuts", mode="before")
    @classmethod
    def _none_to_empty(cls, v: Any) -> Any:
        return [] if v is None else v


def normalize_optimizer_plan_mapping(raw: Mapping[str, Any] | None) -> dict[str, Any]:
    """Подготовка сырого dict: None → пустые списки для известных полей."""
    if raw is None:
        return {}
    if not isinstance(raw, Mapping):
        raise PlanSnapshotValidationError(f"plan must be a mapping, got {type(raw)!r}")
    data = dict(raw)
    for key in ("orders_requested", "primary_cuts", "secondary_cuts", "transverse_cuts"):
        if data.get(key) is None:
            data[key] = []
    return data


def parse_cascading_plan(raw: Mapping[str, Any] | None) -> CascadingPlanSnapshot:
    """Валидация и нормализация одного плана (слой целостности для downstream-кода)."""
    normalized = normalize_optimizer_plan_mapping(raw)
    try:
        return CascadingPlanSnapshot.model_validate(normalized)
    except ValidationError as e:
        raise PlanSnapshotValidationError(str(e), errors=e.errors()) from e


def snapshot_to_trim_dict(snapshot: CascadingPlanSnapshot | Mapping[str, Any]) -> dict[str, Any]:
    """Словарь, совместимый с `_calc_trim_components` / производственными обходами плана."""
    if isinstance(snapshot, Mapping):
        snapshot = parse_cascading_plan(snapshot)
    return snapshot.model_dump(mode="python")


def parse_plan_by_load(
    raw: Mapping[Any, Mapping[str, Any] | MutableMapping[str, Any] | None] | None,
) -> dict[Any, CascadingPlanSnapshot]:
    """Валидация карты нагрузка→план; пустые значения становятся пустым снимком."""
    if not raw:
        return {}
    out: dict[Any, CascadingPlanSnapshot] = {}
    for key, plan_raw in raw.items():
        if plan_raw is None:
            out[key] = CascadingPlanSnapshot()
            continue
        out[key] = parse_cascading_plan(plan_raw)
    return out
