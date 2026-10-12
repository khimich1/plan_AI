"""Контракт ручек, превью и записи себестоимости цельных свай.

Объём сваи и килограммы арматуры в запросы не входят: их нельзя прислать.
Разложение одной марки отдаёт их только для чтения.
"""

from __future__ import annotations

from pydantic import BaseModel, ConfigDict, Field


class PileCostArticlePrice(BaseModel):
    model_config = ConfigDict(extra="forbid")

    article: str
    price: float


class PileCostRecipeInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    concrete_grade: str
    cement_kg_m3: float
    sand_kg_m3: float
    crushed_kg_m3: float


class PileCostOverheadInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    general: float
    payroll: float
    forming: float
    other: float


class PileCostInputsBody(BaseModel):
    """Ручки. Объём и килограммы нормы сюда не входят."""

    model_config = ConfigDict(extra="forbid")

    prices: list[PileCostArticlePrice]
    recipes: list[PileCostRecipeInput]
    overhead: PileCostOverheadInput


class PileCostGradePrice(BaseModel):
    concrete_grade: str
    buyer_price: float | None = None


class PileCostMarkSummary(BaseModel):
    """Краткая строка списка: имя, сечение и пять цен, без разложения."""

    mark: str
    section_cm: int
    grades: list[PileCostGradePrice]


class PileCostPricedLine(BaseModel):
    article: str
    kg: float
    rub: float


class PileCostGradeBreakdown(BaseModel):
    concrete_grade: str
    materials_rub: float
    expenses_rub: float
    cost: float
    buyer_price: float
    surcharge_applied: bool
    cement_kg: float
    sand_kg: float
    crushed_kg: float
    additive_kg: float
    cement_rub: float
    sand_rub: float
    crushed_rub: float
    additive_rub: float
    crushed_article: str
    lines: list[PileCostPricedLine]


class PileCostMarkBreakdown(BaseModel):
    """Разложение пяти классов одной марки. Килограммы только для чтения."""

    mark: str
    section_cm: int
    volume_m3: float
    grades: list[PileCostGradeBreakdown]


class PileCostState(BaseModel):
    inputs_sha256: str
    inputs: PileCostInputsBody
    marks: list[PileCostMarkSummary]
    breakdown: PileCostMarkBreakdown | None = None


class PileCostSaved(BaseModel):
    inputs_sha256: str
    inputs: PileCostInputsBody


class PileCostDiffExample(BaseModel):
    key: str
    old_price: float | None = None
    new_price: float | None = None


class PileCostDiffExamples(BaseModel):
    changed: list[PileCostDiffExample] = Field(default_factory=list)
    new: list[PileCostDiffExample] = Field(default_factory=list)
    missing: list[PileCostDiffExample] = Field(default_factory=list)
    unchanged: list[PileCostDiffExample] = Field(default_factory=list)


class PileCostPreview(BaseModel):
    inputs_sha256: str
    changed: int
    new: int
    missing: int
    unchanged: int
    examples: PileCostDiffExamples


class PileCostApplyRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    inputs_sha256: str | None = None
