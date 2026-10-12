"""Себестоимость цельной сваи: чистая формула, без Excel и SQLite."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass

ADDITIVE_OF_CEMENT = 0.006
INFLATION = 0.03
MARGIN = 1.2
VAT = 1.22
SECTION_SURCHARGE = 1.02
SURCHARGE_SECTIONS_CM = frozenset({30, 35, 40})

ARTICLE_CEMENT = "cement"
ARTICLE_SAND = "sand"
ARTICLE_CRUSHED_20 = "crushed_20"
ARTICLE_CRUSHED_5 = "crushed_5"
ARTICLE_ADDITIVE = "additive"
ARTICLE_WIRE_5 = "wire_5"
ARTICLE_AI_10 = "ai_10"
ARTICLE_A500C_10 = "a500c_10"
ARTICLE_A500C_12 = "a500c_12"

GRANITE_GRADE = "B30_granite"


@dataclass(frozen=True, slots=True)
class Recipe:
    """Килограммы цемента, песка и щебня на кубометр бетона."""

    cement_kg_m3: float
    sand_kg_m3: float
    crushed_kg_m3: float


@dataclass(frozen=True, slots=True)
class NormLine:
    """Килограммы одной статьи нормы: металл или «Щ.5»."""

    article: str
    kg: float


@dataclass(frozen=True, slots=True)
class PricedLine:
    """Строка нормы после умножения на цену своей статьи."""

    article: str
    kg: float
    rub: float


@dataclass(frozen=True, slots=True)
class PriceBook:
    """Руб/кг по коду статьи."""

    prices: Mapping[str, float]

    def __post_init__(self) -> None:
        object.__setattr__(self, "prices", {str(key): float(value) for key, value in self.prices.items()})

    def unit_price(self, article: str) -> float:
        if article not in self.prices:
            raise ValueError(f"нет цены статьи {article}")
        price = self.prices[article]
        if price <= 0:
            raise ValueError(f"цена статьи {article} должна быть больше нуля")
        return price


@dataclass(frozen=True, slots=True)
class Overhead:
    """Четыре статьи расходов, руб/м³, до инфляции."""

    general: float
    payroll: float
    forming: float
    other: float


@dataclass(frozen=True, slots=True)
class PileCost:
    """Разложение себестоимости и цена для покупателя без округления до рубля."""

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
    lines: tuple[PricedLine, ...]

    @property
    def metal_rub(self) -> float:
        return sum(line.rub for line in self.lines if line.article != ARTICLE_CRUSHED_5)


def _crushed_article(concrete_grade: str) -> str:
    if concrete_grade == GRANITE_GRADE:
        return ARTICLE_CRUSHED_5
    return ARTICLE_CRUSHED_20


def _additive_kg(cement_kg: float, entered_kg: float | None) -> float:
    if entered_kg is None:
        return cement_kg * ADDITIVE_OF_CEMENT
    if entered_kg < 0:
        raise ValueError("килограммы добавки не могут быть отрицательными")
    return entered_kg


def _price_lines(lines: Sequence[NormLine], prices: PriceBook) -> tuple[PricedLine, ...]:
    priced: list[PricedLine] = []
    for line in lines:
        if line.kg < 0:
            raise ValueError(f"килограммы статьи {line.article} не могут быть отрицательными")
        unit = prices.unit_price(line.article)
        priced.append(PricedLine(article=line.article, kg=line.kg, rub=line.kg * unit))
    return tuple(priced)


def _expenses_rub(volume_m3: float, overhead: Overhead) -> float:
    rate = (overhead.general + overhead.payroll + overhead.forming + overhead.other) * (1 + INFLATION)
    return volume_m3 * rate


def buyer_price(
    *,
    section_cm: int,
    volume_m3: float,
    concrete_grade: str,
    recipe: Recipe,
    lines: Sequence[NormLine],
    prices: PriceBook,
    overhead: Overhead,
    additive_kg: float | None = None,
) -> PileCost:
    """Себестоимость и цена покупателя. Отрицательный объём или цена — ValueError."""
    if volume_m3 < 0:
        raise ValueError("объём не может быть отрицательным")

    cement_kg = volume_m3 * recipe.cement_kg_m3
    sand_kg = volume_m3 * recipe.sand_kg_m3
    crushed_kg = volume_m3 * recipe.crushed_kg_m3
    additive = _additive_kg(cement_kg, additive_kg)
    crushed_article = _crushed_article(concrete_grade)

    cement_rub = cement_kg * prices.unit_price(ARTICLE_CEMENT)
    sand_rub = sand_kg * prices.unit_price(ARTICLE_SAND)
    crushed_rub = crushed_kg * prices.unit_price(crushed_article)
    additive_rub = additive * prices.unit_price(ARTICLE_ADDITIVE)
    priced_lines = _price_lines(lines, prices)
    materials = cement_rub + sand_rub + crushed_rub + additive_rub + sum(line.rub for line in priced_lines)
    expenses = _expenses_rub(volume_m3, overhead)
    cost = materials + expenses
    applied = section_cm in SURCHARGE_SECTIONS_CM
    buyer = cost * MARGIN * VAT * (SECTION_SURCHARGE if applied else 1.0)

    return PileCost(
        materials_rub=materials,
        expenses_rub=expenses,
        cost=cost,
        buyer_price=buyer,
        surcharge_applied=applied,
        cement_kg=cement_kg,
        sand_kg=sand_kg,
        crushed_kg=crushed_kg,
        additive_kg=additive,
        cement_rub=cement_rub,
        sand_rub=sand_rub,
        crushed_rub=crushed_rub,
        additive_rub=additive_rub,
        crushed_article=crushed_article,
        lines=priced_lines,
    )
