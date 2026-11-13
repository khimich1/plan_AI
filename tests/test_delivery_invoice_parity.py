"""Контракт delivery_invoice_lines: два эталона 1С и края.

Цель строки — сегодняшняя сумма к оплате (прайс после скидки КП плюс доля
доставки, до копейки). Скидка документа = округлённая база − сумма целей.
"""

from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP

import pytest

from core.delivery_invoice_parity import ParityLine, delivery_invoice_lines

_KOPECK = Decimal("0.01")
_PRICE_A = Decimal("8949.283209360001")
_PRICE_B = Decimal("13312.012003200003")
_QTY = [200, 30]
_DELIVERY = [Decimal("146087.10"), Decimal("21912.90")]


def _target(price: Decimal, discount_percent: str, qty: int, delivery: Decimal) -> Decimal:
    factor = Decimal(1) - Decimal(discount_percent) / Decimal(100)
    return (price * factor * qty + delivery).quantize(_KOPECK, rounding=ROUND_HALF_UP)


def _lines(
    prices: list[Decimal],
    qty: list[int],
    targets: list[Decimal],
    delivery: list[Decimal] | None = None,
) -> list[ParityLine] | None:
    if delivery is None:
        delivery = [Decimal(0)] * len(prices)
    return delivery_invoice_lines(
        prices=prices,
        qty=qty,
        target_payable=targets,
        delivery_rub=delivery,
    )


def test_reference_discount_25_matches_invoice_478() -> None:
    """Скидка КП 25 %: ручная скидка 17,33 %, итог 1 809 912,75."""
    targets = [
        _target(_PRICE_A, "25", 200, _DELIVERY[0]),
        _target(_PRICE_B, "25", 30, _DELIVERY[1]),
    ]
    assert targets == [Decimal("1488479.58"), Decimal("321433.17")]
    assert sum(targets) == Decimal("1809912.75")

    rows = _lines([_PRICE_A, _PRICE_B], _QTY, targets, _DELIVERY)
    assert rows is not None
    assert [row.discount_percent for row in rows] == [Decimal("17.33"), Decimal("17.33")]
    assert [row.price for row in rows] == [Decimal("8949.28"), Decimal("13312.01")]
    assert [row.discount_amount for row in rows] == [Decimal("310110.40"), Decimal("69193.15")]
    assert [row.payable for row in rows] == [Decimal("1479745.60"), Decimal("330167.15")]
    assert [row.vat for row in rows] == [Decimal("266839.37"), Decimal("59538.34")]
    assert [row.net for row in rows] == [Decimal("1212906.23"), Decimal("270628.81")]
    assert [row.invoice_unit_price for row in rows] == [Decimal("7398.73"), Decimal("11005.57")]
    assert sum(row.payable for row in rows) == sum(targets)
    assert sum(row.discount_amount for row in rows) == Decimal("2189216.30") - sum(targets)


def test_reference_discount_5_matches_invoice_481() -> None:
    """Скидка КП 5 %: наценка −2,67 %, итог 2 247 756,15."""
    targets = [
        _target(_PRICE_A, "5", 200, _DELIVERY[0]),
        _target(_PRICE_B, "5", 30, _DELIVERY[1]),
    ]
    assert targets == [Decimal("1846450.91"), Decimal("401305.24")]
    assert sum(targets) == Decimal("2247756.15")

    rows = _lines([_PRICE_A, _PRICE_B], _QTY, targets, _DELIVERY)
    assert rows is not None
    assert [row.discount_percent for row in rows] == [Decimal("-2.67"), Decimal("-2.67")]
    assert [row.price for row in rows] == [Decimal("8949.28"), Decimal("13312.01")]
    assert [row.discount_amount for row in rows] == [Decimal("-47860.92"), Decimal("-10678.93")]
    assert [row.payable for row in rows] == [Decimal("1837716.92"), Decimal("410039.23")]
    assert [row.vat for row in rows] == [Decimal("331391.58"), Decimal("73941.50")]
    assert [row.net for row in rows] == [Decimal("1506325.34"), Decimal("336097.73")]
    assert [row.invoice_unit_price for row in rows] == [Decimal("9188.58"), Decimal("13667.97")]
    assert sum(row.payable for row in rows) == sum(targets)
    assert sum(row.discount_amount for row in rows) == Decimal("2189216.30") - sum(targets)


def test_single_line_takes_whole_positive_and_negative_discount() -> None:
    positive = _lines(
        [Decimal("100.00")],
        [2],
        [Decimal("150.00")],
    )
    assert positive is not None
    assert positive[0].discount_amount == Decimal("50.00")
    assert positive[0].discount_percent == Decimal("25.00")
    assert positive[0].payable == Decimal("150.00")

    negative = _lines(
        [Decimal("100.00")],
        [2],
        [Decimal("250.00")],
    )
    assert negative is not None
    assert negative[0].discount_amount == Decimal("-50.00")
    assert negative[0].discount_percent == Decimal("-25.00")
    assert negative[0].payable == Decimal("250.00")
    assert negative[0].invoice_unit_price == Decimal("125.00")


def test_zero_qty_line_is_outside_base_and_does_not_take_remainder() -> None:
    """Остаток копеек — на последнюю строку с количеством больше нуля."""
    rows = _lines(
        [Decimal("1.00"), Decimal("1.00"), Decimal("1.00"), Decimal("9.00")],
        [1, 1, 1, 0],
        [Decimal("1.00"), Decimal("1.00"), Decimal("0.99"), Decimal("0.00")],
    )
    assert rows is not None
    assert [row.discount_amount for row in rows] == [
        Decimal("0.00"),
        Decimal("0.00"),
        Decimal("0.01"),
        Decimal("0.00"),
    ]
    assert rows[3].payable == Decimal("0.00")
    assert rows[3].price == Decimal("9.00")
    assert rows[3].invoice_unit_price == Decimal("9.00")
    assert rows[3].discount_percent == rows[0].discount_percent
    assert sum(row.payable for row in rows) == Decimal("2.99")
    assert sum(row.discount_amount for row in rows) == Decimal("3.00") - Decimal("2.99")


def test_zero_document_base_returns_none() -> None:
    assert _lines([Decimal("10.00")], [0], [Decimal("0")]) is None
    assert _lines([Decimal("0.001"), Decimal("0")], [4, 3], [Decimal("1"), Decimal("2")]) is None


def test_zero_document_discount_invoice_price_equals_rounded_list_price() -> None:
    rows = _lines(
        [Decimal("10.004"), Decimal("10.005")],
        [3, 1],
        [Decimal("30.00"), Decimal("10.01")],
    )
    assert rows is not None
    assert [row.discount_amount for row in rows] == [Decimal("0.00"), Decimal("0.00")]
    assert [row.discount_percent for row in rows] == [Decimal("0.00"), Decimal("0.00")]
    assert [row.invoice_unit_price for row in rows] == [Decimal("10.00"), Decimal("10.01")]
    assert [row.price for row in rows] == [row.invoice_unit_price for row in rows]


def test_lengths_must_match() -> None:
    with pytest.raises(ValueError):
        delivery_invoice_lines(
            prices=[Decimal("1")],
            qty=[1, 2],
            target_payable=[Decimal("1")],
            delivery_rub=[Decimal("0")],
        )
