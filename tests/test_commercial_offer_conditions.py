from __future__ import annotations

from core.commercial_offer import (
    _format_delivery_conditions_text,
    _format_payment_conditions_text,
)


def test_format_delivery_conditions_custom() -> None:
    assert _format_delivery_conditions_text("Самовывоз") == "1. Условия поставки: Самовывоз"


def test_format_delivery_conditions_default() -> None:
    assert _format_delivery_conditions_text("") == "1. Условия поставки:"
    assert _format_delivery_conditions_text(None) == "1. Условия поставки:"
    assert _format_delivery_conditions_text("   ") == "1. Условия поставки:"


def test_format_payment_conditions_custom() -> None:
    assert _format_payment_conditions_text("Оплата 50/50") == "2. Условия оплаты: Оплата 50/50"


def test_format_payment_conditions_default() -> None:
    assert _format_payment_conditions_text("") == (
        "2. Условия оплаты: Предварительная оплата в размере 100%"
    )
    assert _format_payment_conditions_text(None) == (
        "2. Условия оплаты: Предварительная оплата в размере 100%"
    )
