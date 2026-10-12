"""PC-001: чистая формула цены цельной сваи.

Входы С30.20-3 и С40.30-1 — ячейки, которые лист «Стоимость» от 06.10.2026
умножает на цену. Ожидаемые рубли — критерии спеки, не подгонка под код.
"""

from __future__ import annotations

import pytest

from core.pile_cost import (
    ARTICLE_A500C_10,
    ARTICLE_A500C_12,
    ARTICLE_ADDITIVE,
    ARTICLE_AI_10,
    ARTICLE_CEMENT,
    ARTICLE_CRUSHED_20,
    ARTICLE_CRUSHED_5,
    ARTICLE_SAND,
    ARTICLE_WIRE_5,
    NormLine,
    Overhead,
    PriceBook,
    Recipe,
    buyer_price,
)

# Руб/кг листа «Цены» от 06.10.2026. А500С Ø10 и Ø12 в файле стоят одинаково.
_PRICES_2026_10_06: dict[str, float] = {
    ARTICLE_CEMENT: 7.54,
    ARTICLE_CRUSHED_20: 1.80,
    ARTICLE_CRUSHED_5: 1.97,
    ARTICLE_SAND: 0.66,
    ARTICLE_ADDITIVE: 59.0,
    ARTICLE_WIRE_5: 72.7,
    ARTICLE_AI_10: 50.57,
    ARTICLE_A500C_10: 51.39,
    ARTICLE_A500C_12: 51.39,
}

_OVERHEAD = Overhead(general=3186, payroll=3400, forming=40, other=606)

_RECIPE_B15 = Recipe(cement_kg_m3=320, sand_kg_m3=990, crushed_kg_m3=950)
_RECIPE_B20 = Recipe(cement_kg_m3=340, sand_kg_m3=950, crushed_kg_m3=950)
_RECIPE_B30_GRANITE = Recipe(cement_kg_m3=400, sand_kg_m3=920, crushed_kg_m3=950)

# С30.20-3: объём 0,13; спираль Вр-I Ø5 3,1; поперечная А-I Ø10 3,15; рабочая А500С Ø10 8.
_S30_LINES = (
    (ARTICLE_WIRE_5, 3.1),
    (ARTICLE_AI_10, 3.15),
    (ARTICLE_A500C_10, 8.0),
)

# С40.30-1: объём 0,37; спираль 6,5; поперечная А-I Ø10 20. Рабочей строки формула не умножает.
_S40_LINES = (
    (ARTICLE_WIRE_5, 6.5),
    (ARTICLE_AI_10, 20.0),
)

_TOLERANCE = 0.01


def _book(**overrides: float) -> PriceBook:
    prices = dict(_PRICES_2026_10_06)
    prices.update(overrides)
    return PriceBook(prices)


def _lines(*pairs: tuple[str, float]) -> tuple[NormLine, ...]:
    return tuple(NormLine(article, kg) for article, kg in pairs)


def test_s30_20_3_b15_matches_spec_without_surcharge() -> None:
    result = buyer_price(
        section_cm=20,
        volume_m3=0.13,
        concrete_grade="B15",
        recipe=_RECIPE_B15,
        lines=_lines(*_S30_LINES),
        prices=_book(),
        overhead=_OVERHEAD,
    )

    assert result.materials_rub == pytest.approx(1431.4179, abs=_TOLERANCE)
    assert result.expenses_rub == pytest.approx(968.3648, abs=_TOLERANCE)
    assert result.cost == pytest.approx(2399.7827, abs=_TOLERANCE)
    assert result.buyer_price == pytest.approx(3513.2818728, abs=_TOLERANCE)
    assert result.surcharge_applied is False


def test_s30_20_3_b30_granite_uses_crushed_stone_197_without_surcharge() -> None:
    result = buyer_price(
        section_cm=20,
        volume_m3=0.13,
        concrete_grade="B30_granite",
        recipe=_RECIPE_B30_GRANITE,
        lines=_lines(*_S30_LINES),
        prices=_book(),
        overhead=_OVERHEAD,
    )

    assert result.buyer_price == pytest.approx(3655.4166552, abs=_TOLERANCE)
    assert result.surcharge_applied is False
    assert result.crushed_article == ARTICLE_CRUSHED_5
    assert result.crushed_rub == pytest.approx(0.13 * 950 * 1.97, abs=_TOLERANCE)

    granite_ignores_crushed_20 = buyer_price(
        section_cm=20,
        volume_m3=0.13,
        concrete_grade="B30_granite",
        recipe=_RECIPE_B30_GRANITE,
        lines=_lines(*_S30_LINES),
        prices=_book(**{ARTICLE_CRUSHED_20: 9.99}),
        overhead=_OVERHEAD,
    )
    assert granite_ignores_crushed_20.buyer_price == pytest.approx(
        result.buyer_price, abs=_TOLERANCE
    )


def test_s40_30_1_b15_applies_section_surcharge() -> None:
    result = buyer_price(
        section_cm=30,
        volume_m3=0.37,
        concrete_grade="B15",
        recipe=_RECIPE_B15,
        lines=_lines(*_S40_LINES),
        prices=_book(),
        overhead=_OVERHEAD,
    )

    assert result.buyer_price == pytest.approx(9033.108758784, abs=_TOLERANCE)
    assert result.surcharge_applied is True


def test_b15_recipe_does_not_change_b20() -> None:
    shared = dict(
        section_cm=20,
        volume_m3=0.13,
        lines=_lines(*_S30_LINES),
        prices=_book(),
        overhead=_OVERHEAD,
    )
    b20_before = buyer_price(concrete_grade="B20", recipe=_RECIPE_B20, **shared)
    b15_before = buyer_price(concrete_grade="B15", recipe=_RECIPE_B15, **shared)

    moved_b15 = Recipe(cement_kg_m3=400, sand_kg_m3=990, crushed_kg_m3=950)
    b20_after = buyer_price(concrete_grade="B20", recipe=_RECIPE_B20, **shared)
    b15_after = buyer_price(concrete_grade="B15", recipe=moved_b15, **shared)

    assert b20_after.buyer_price == pytest.approx(b20_before.buyer_price, abs=_TOLERANCE)
    assert b15_after.buyer_price != pytest.approx(b15_before.buyer_price, abs=_TOLERANCE)


def test_cement_price_does_not_change_metal_rubles() -> None:
    base = buyer_price(
        section_cm=20,
        volume_m3=0.13,
        concrete_grade="B15",
        recipe=_RECIPE_B15,
        lines=_lines(*_S30_LINES),
        prices=_book(),
        overhead=_OVERHEAD,
    )
    dearer_cement = buyer_price(
        section_cm=20,
        volume_m3=0.13,
        concrete_grade="B15",
        recipe=_RECIPE_B15,
        lines=_lines(*_S30_LINES),
        prices=_book(**{ARTICLE_CEMENT: 10.0}),
        overhead=_OVERHEAD,
    )

    assert dearer_cement.metal_rub == pytest.approx(base.metal_rub, abs=_TOLERANCE)
    assert dearer_cement.cement_rub != pytest.approx(base.cement_rub, abs=_TOLERANCE)
    assert dearer_cement.materials_rub != pytest.approx(base.materials_rub, abs=_TOLERANCE)


def test_a500c_10_price_does_not_change_mark_that_has_only_a500c_12() -> None:
    only_12 = _lines((ARTICLE_A500C_12, 8.0))
    cheap = _book()
    dear_10 = _book(**{ARTICLE_A500C_10: 80.0})
    kwargs = dict(
        section_cm=20,
        volume_m3=0.13,
        concrete_grade="B15",
        recipe=_RECIPE_B15,
        lines=only_12,
        overhead=_OVERHEAD,
    )

    assert buyer_price(prices=cheap, **kwargs).buyer_price == pytest.approx(
        buyer_price(prices=dear_10, **kwargs).buyer_price,
        abs=_TOLERANCE,
    )

    with_10 = dict(kwargs, lines=_lines((ARTICLE_A500C_10, 8.0)))
    assert buyer_price(prices=cheap, **with_10).buyer_price != pytest.approx(
        buyer_price(prices=dear_10, **with_10).buyer_price,
        abs=_TOLERANCE,
    )


def test_empty_lines_count_only_concrete_and_expenses() -> None:
    result = buyer_price(
        section_cm=20,
        volume_m3=0.13,
        concrete_grade="B15",
        recipe=_RECIPE_B15,
        lines=(),
        prices=_book(),
        overhead=_OVERHEAD,
    )

    assert result.lines == ()
    assert result.metal_rub == pytest.approx(0, abs=_TOLERANCE)
    assert result.materials_rub == pytest.approx(635.6324, abs=_TOLERANCE)
    assert result.expenses_rub == pytest.approx(968.3648, abs=_TOLERANCE)
    assert result.cost == pytest.approx(1603.9972, abs=_TOLERANCE)
    assert result.buyer_price == pytest.approx(2348.2519008, abs=_TOLERANCE)
    assert result.surcharge_applied is False


def test_hand_entered_additive_replaces_cement_fraction() -> None:
    default = buyer_price(
        section_cm=20,
        volume_m3=0.13,
        concrete_grade="B15",
        recipe=_RECIPE_B15,
        lines=(),
        prices=_book(),
        overhead=_OVERHEAD,
    )
    entered = buyer_price(
        section_cm=20,
        volume_m3=0.13,
        concrete_grade="B15",
        recipe=_RECIPE_B15,
        lines=(),
        prices=_book(),
        overhead=_OVERHEAD,
        additive_kg=4.6,
    )

    assert default.additive_kg == pytest.approx(0.13 * 320 * 0.006, abs=_TOLERANCE)
    assert entered.additive_kg == pytest.approx(4.6, abs=_TOLERANCE)
    assert entered.additive_rub == pytest.approx(4.6 * 59, abs=_TOLERANCE)
    assert entered.materials_rub != pytest.approx(default.materials_rub, abs=_TOLERANCE)


@pytest.mark.parametrize("section_cm", [15, 20])
def test_sections_15_and_20_skip_surcharge(section_cm: int) -> None:
    result = buyer_price(
        section_cm=section_cm,
        volume_m3=0.13,
        concrete_grade="B15",
        recipe=_RECIPE_B15,
        lines=_lines(*_S30_LINES),
        prices=_book(),
        overhead=_OVERHEAD,
    )
    assert result.surcharge_applied is False
    assert result.buyer_price == pytest.approx(3513.2818728, abs=_TOLERANCE)


@pytest.mark.parametrize("section_cm", [30, 35, 40])
def test_sections_30_35_40_multiply_by_102(section_cm: int) -> None:
    result = buyer_price(
        section_cm=section_cm,
        volume_m3=0.13,
        concrete_grade="B15",
        recipe=_RECIPE_B15,
        lines=_lines(*_S30_LINES),
        prices=_book(),
        overhead=_OVERHEAD,
    )
    assert result.surcharge_applied is True
    assert result.buyer_price == pytest.approx(3513.2818728 * 1.02, abs=_TOLERANCE)


def test_negative_volume_is_an_error() -> None:
    with pytest.raises(ValueError):
        buyer_price(
            section_cm=20,
            volume_m3=-0.13,
            concrete_grade="B15",
            recipe=_RECIPE_B15,
            lines=(),
            prices=_book(),
            overhead=_OVERHEAD,
        )


@pytest.mark.parametrize("price", [0, -1.5])
def test_non_positive_price_is_an_error(price: float) -> None:
    with pytest.raises(ValueError):
        buyer_price(
            section_cm=20,
            volume_m3=0.13,
            concrete_grade="B15",
            recipe=_RECIPE_B15,
            lines=_lines(*_S30_LINES),
            prices=_book(**{ARTICLE_CEMENT: price}),
            overhead=_OVERHEAD,
        )


def test_b15_does_not_use_crushed_5_price() -> None:
    result = buyer_price(
        section_cm=20,
        volume_m3=0.13,
        concrete_grade="B15",
        recipe=_RECIPE_B15,
        lines=_lines(*_S30_LINES),
        prices=_book(**{ARTICLE_CRUSHED_5: 9.99}),
        overhead=_OVERHEAD,
    )
    assert result.crushed_article == ARTICLE_CRUSHED_20
    assert result.materials_rub == pytest.approx(1431.4179, abs=_TOLERANCE)
    assert result.buyer_price == pytest.approx(3513.2818728, abs=_TOLERANCE)
