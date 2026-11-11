"""Скидка или наценка документа, когда в строках уже есть доля доставки.

Чистая арифметика: без Excel, без базы и без чтения файлов.
Долю доставки сюда передают только чтобы сверить длину списков.
В сумму строки она не прибавляется: она уже входит в цель.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP

_KOPECK = Decimal("0.01")
_HUNDRED = Decimal(100)
_VAT_INCLUDED = Decimal(22) / Decimal(122)


@dataclass(frozen=True)
class ParityLine:
    price: Decimal
    discount_percent: Decimal
    discount_amount: Decimal
    net: Decimal
    vat: Decimal
    payable: Decimal
    invoice_unit_price: Decimal


def delivery_invoice_lines(
    *,
    prices: list[Decimal],
    qty: list[int],
    target_payable: list[Decimal],
    delivery_rub: list[Decimal],
) -> list[ParityLine] | None:
    """Строки варианта А. Скидка может быть отрицательной.

    None — база документа равна нулю, вызвать сегодняшнюю формулу.
    """
    _require_same_length(prices, qty, target_payable, delivery_rub)
    rounded = [_kopecks(price) for price in prices]
    counts = [_non_negative_qty(count) for count in qty]
    bases = [rounded[index] * counts[index] for index in range(len(rounded))]
    document_base = _document_base(bases, counts)
    if document_base == 0:
        return None
    discount = document_base - sum(_kopecks(amount) for amount in target_payable)
    percent = _kopecks(discount / document_base * _HUNDRED)
    amounts = _line_discounts(discount, bases, counts, document_base)
    return [
        _parity_line(rounded[index], percent, amounts[index], counts[index])
        for index in range(len(rounded))
    ]


def _document_base(bases: list[Decimal], counts: list[int]) -> Decimal:
    return sum(
        (base for base, count in zip(bases, counts) if count > 0),
        start=Decimal(0),
    )


def _line_discounts(
    discount: Decimal,
    bases: list[Decimal],
    counts: list[int],
    document_base: Decimal,
) -> list[Decimal]:
    """Пропорция по базе. Остаток копеек — последней строке с количеством."""
    amounts = [Decimal(0)] * len(bases)
    positive = [index for index, count in enumerate(counts) if count > 0]
    allocated = Decimal(0)
    for index in positive[:-1]:
        share = _kopecks(discount * bases[index] / document_base)
        amounts[index] = share
        allocated += share
    amounts[positive[-1]] = discount - allocated
    return amounts


def _parity_line(
    price: Decimal,
    percent: Decimal,
    discount: Decimal,
    qty: int,
) -> ParityLine:
    payable = price * qty - discount
    vat = _kopecks(payable * _VAT_INCLUDED)
    unit = _kopecks(payable / qty) if qty > 0 else price
    return ParityLine(
        price=price,
        discount_percent=percent,
        discount_amount=discount,
        net=payable - vat,
        vat=vat,
        payable=payable,
        invoice_unit_price=unit,
    )


def _require_same_length(
    prices: list[Decimal],
    qty: list[int],
    target_payable: list[Decimal],
    delivery_rub: list[Decimal],
) -> None:
    width = len(prices)
    if len(qty) != width or len(target_payable) != width or len(delivery_rub) != width:
        raise ValueError("Списки цен, количеств, целей и доставки разной длины")


def _non_negative_qty(raw: int) -> int:
    try:
        count = int(raw)
    except (TypeError, ValueError):
        return 0
    return max(0, count)


def _kopecks(amount: Decimal) -> Decimal:
    return Decimal(amount).quantize(_KOPECK, rounding=ROUND_HALF_UP)
