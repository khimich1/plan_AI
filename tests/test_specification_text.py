"""Абзацы спецификации: доли, пропись, день/дня/дней, условие предоплаты."""

from __future__ import annotations

from datetime import date

import pytest

from core.specification_text import (
    SpecificationChoice,
    SpecificationTextError,
    calendar_days_phrase,
    payment_tranche_kopecks,
    render_specification,
    rubles_in_words,
    total_amount_words,
)

_FULL_CLAUSE = (
    "при условии поступления предварительной оплаты на расчетный счет поставщика "
    "в полном объеме в установленный в настоящей спецификации срок"
)
_PARTIAL_CLAUSE = (
    "при условии поступления авансов, указанных в условиях оплаты настоящей спецификации, "
    "на расчетный счет поставщика в установленный в них срок"
)
_DUE = date(2026, 10, 1)
_FIRST = date(2026, 9, 18)
_SECOND = date(2026, 10, 1)


def _choice(**overrides: object) -> SpecificationChoice:
    base = dict(
        payment="prepay_100",
        term="by_date",
        term_date=_DUE,
        delivery="pickup",
    )
    base.update(overrides)
    return SpecificationChoice(**base)  # type: ignore[arg-type]


def test_whole_tranche_words_end_with_rubles_and_skip_zero_kopecks() -> None:
    words = rubles_in_words(100_000)

    assert words == "одна тысяча рублей"
    assert words.endswith("рублей")
    assert "00 копеек" not in words
    assert "копе" not in words


def test_total_line_words_keep_kopecks_and_start_with_a_capital() -> None:
    assert total_amount_words(100_000) == "Одна тысяча рублей 00 копеек"
    assert total_amount_words(355_480_200) == (
        "Три миллиона пятьсот пятьдесят четыре тысячи восемьсот два рубля 00 копеек"
    )
    assert total_amount_words(101) == "Один рубль 01 копейка"
    assert rubles_in_words(100_000) == "одна тысяча рублей"


def test_kopecks_are_written_only_when_present() -> None:
    assert rubles_in_words(100) == "один рубль"
    assert rubles_in_words(200) == "два рубля"
    assert rubles_in_words(500) == "пять рублей"
    assert rubles_in_words(2_100) == "двадцать один рубль"
    assert rubles_in_words(1_100) == "одиннадцать рублей"
    assert rubles_in_words(101) == "один рубль 01 копейка"
    assert rubles_in_words(102) == "один рубль 02 копейки"
    assert rubles_in_words(105) == "один рубль 05 копеек"


def test_day_agreement_one_three_five_twenty_one() -> None:
    assert calendar_days_phrase(1) == "1 (один) календарный день"
    assert calendar_days_phrase(3) == "3 (три) календарных дня"
    assert calendar_days_phrase(5) == "5 (пять) календарных дней"
    assert calendar_days_phrase(21) == "21 (двадцать один) календарный день"


def test_deferral_uses_day_agreement() -> None:
    for days, expected in (
        (1, "Отсрочка платежа — 1 (один) календарный день."),
        (3, "Отсрочка платежа — 3 (три) календарных дня."),
        (5, "Отсрочка платежа — 5 (пять) календарных дней."),
        (21, "Отсрочка платежа — 21 (двадцать один) календарный день."),
    ):
        paragraphs = render_specification(
            _choice(payment="deferral", payment_days=days),
            payable_total=100_000,
        )
        assert paragraphs.payment == expected


def test_split_50_50_and_share_remainder_sum_to_total() -> None:
    half = payment_tranche_kopecks("split_50_50", 1001)
    share = payment_tranche_kopecks("split_share", 1001, share_percent=20)

    assert half == (501, 500)
    assert sum(half) == 1001
    assert share == (501, 200, 300)
    assert sum(share) == 1001


def test_prepay_100_paragraph_has_amount_and_words() -> None:
    paragraphs = render_specification(_choice(), payable_total=100_000)

    assert paragraphs.payment == (
        "1 000,00 (одна тысяча рублей) — предварительная оплата в размере 100%."
    )


def test_split_50_50_paragraph_uses_remainder_and_within_is_spelled() -> None:
    paragraphs = render_specification(
        _choice(
            payment="split_50_50",
            payment_date=_FIRST,
            payment_days=3,
            term="after_payment",
            term_days=10,
            term_date=None,
        ),
        payable_total=100_000,
    )

    assert paragraphs.payment == (
        "500,00 (пятьсот рублей) — предварительная оплата в размере 50% до 18.09.2026 включительно. "
        "500,00 (пятьсот рублей) — за 3 (три) дня до планируемого начала отгрузки."
    )
    assert "в течении" not in paragraphs.term
    assert "в течение" in paragraphs.term


def test_split_share_paragraph_remainder_is_within_calendar_days() -> None:
    paragraphs = render_specification(
        _choice(
            payment="split_share",
            payment_date=_FIRST,
            second_payment_date=_SECOND,
            second_share_percent=20,
            payment_days=5,
        ),
        payable_total=100_000,
    )

    assert paragraphs.payment == (
        "500,00 (пятьсот рублей) — предварительная оплата в размере 50% до 18.09.2026 включительно. "
        "200,00 (двести рублей) — предварительная оплата в размере 20% до 01.10.2026 включительно. "
        "300,00 (триста рублей) — остаток в течение 5 (пять) календарных дней после поставки партии. "
        "Партия — товар по одному УПД."
    )
    assert "в течении" not in paragraphs.payment


def test_odd_kopeck_tranches_appear_in_the_paragraph() -> None:
    paragraphs = render_specification(
        _choice(
            payment="split_50_50",
            payment_date=_FIRST,
            payment_days=1,
        ),
        payable_total=1001,
    )

    assert "5,01 (пять рублей 01 копейка)" in paragraphs.payment
    assert "5,00 (пять рублей)" in paragraphs.payment
    assert "за 1 (один) день" in paragraphs.payment
    assert "00 копеек" not in paragraphs.payment


def test_full_prepay_clause_sticks_only_to_deadline() -> None:
    by_date = render_specification(_choice(payment="prepay_100"), payable_total=100_000)
    after_money = render_specification(
        _choice(payment="prepay_100", term="after_payment", term_days=10, term_date=None),
        payable_total=100_000,
    )
    custom = render_specification(
        _choice(payment="custom", custom_text="Два транша по договорённости."),
        payable_total=100_000,
    )

    assert by_date.term == (
        "Поставщик обязуется поставить товар не позднее 1 октября 2026 года "
        f"{_FULL_CLAUSE}."
    )
    assert _FULL_CLAUSE not in after_money.term
    assert "при условии поступления" not in after_money.term
    assert custom.payment == "Два транша по договорённости."
    assert custom.term.endswith(f"{_FULL_CLAUSE}.")


def test_partial_prepay_clause_is_narrow_and_absent_for_deferral() -> None:
    partial = render_specification(
        _choice(payment="split_50_50", payment_date=_FIRST, payment_days=3),
        payable_total=100_000,
    )
    deferred = render_specification(
        _choice(payment="deferral", payment_days=45),
        payable_total=100_000,
    )
    after_money = render_specification(
        _choice(
            payment="deferral",
            payment_days=45,
            term="after_payment",
            term_days=10,
            term_date=None,
        ),
        payable_total=100_000,
    )

    assert _PARTIAL_CLAUSE in partial.term
    assert _FULL_CLAUSE not in partial.term
    assert deferred.term == "Поставщик обязуется поставить товар не позднее 1 октября 2026 года."
    assert "при условии поступления" not in deferred.term
    assert "при условии поступления" not in after_money.term
    assert after_money.payment == "Отсрочка платежа — 45 (сорок пять) календарных дней."


def test_pile_rhythm_is_absent_unless_the_choice_and_the_kp_have_piles() -> None:
    plates = render_specification(_choice(), payable_total=100_000, has_piles=False)

    assert "свай" not in plates.term
    assert "в неделю" not in plates.term

    piles = render_specification(
        _choice(
            payment="deferral",
            payment_days=45,
            term="pile_rhythm",
            term_date=date(2026, 9, 20),
            pile_count=12,
            pile_unit="week",
        ),
        payable_total=100_000,
        has_piles=True,
    )

    assert piles.term == "Начало поставок с 20.09.2026. По 12 свай в неделю."
    assert "при условии" not in piles.term

    per_day = render_specification(
        _choice(
            payment="deferral",
            payment_days=45,
            term="pile_rhythm",
            term_date=date(2026, 9, 20),
            pile_count=4,
            pile_unit="day",
        ),
        payable_total=100_000,
        has_piles=True,
    )
    assert per_day.term == "Начало поставок с 20.09.2026. По 4 свай в день."

    with pytest.raises(SpecificationTextError, match="сва"):
        render_specification(
            _choice(
                term="pile_rhythm",
                term_date=date(2026, 9, 20),
                pile_count=12,
                pile_unit="week",
            ),
            payable_total=100_000,
            has_piles=False,
        )


def test_pickup_and_site_delivery_paragraphs() -> None:
    pickup = render_specification(_choice(delivery="pickup"), payable_total=100_000)
    site = render_specification(
        _choice(delivery="site", delivery_address="г. Кострома, ул. Лесная, 4"),
        payable_total=100_000,
    )

    assert "150020, г. Ярославль, пр. Домостроителей, дом 1, стр. 3" in pickup.delivery
    assert "обособленное подразделение «ЖБК СТАРТ»" in pickup.delivery
    assert site.delivery == (
        "Поставка на объект покупателя по адресу: г. Кострома, ул. Лесная, 4. "
        "Стоимость доставки включена в стоимость товара."
    )


def test_days_outside_1_to_365_are_refused() -> None:
    with pytest.raises(SpecificationTextError):
        render_specification(_choice(payment="deferral", payment_days=0), payable_total=100)
    with pytest.raises(SpecificationTextError):
        render_specification(_choice(payment="deferral", payment_days=366), payable_total=100)

    edge = render_specification(
        _choice(payment="deferral", payment_days=365),
        payable_total=100,
    )
    assert edge.payment.startswith("Отсрочка платежа — 365 (")
