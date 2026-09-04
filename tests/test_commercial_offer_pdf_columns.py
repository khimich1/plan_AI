"""PDF KP money columns must fit formatted amounts without overflow."""

from __future__ import annotations

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics

from core.commercial_offer import (
    FONT_NORMAL,
    _MONEY_FIT_SAMPLE,
    _TABLE_BODY_SIZE,
    _TABLE_CELL_PAD,
    _format_money_ru,
    _money_col_width,
    generate_commercial_offer_pdf,
)


def test_money_col_width_fits_sample_and_large_sum() -> None:
    col = _money_col_width()
    usable = col - 2 * _TABLE_CELL_PAD

    for sample in (_MONEY_FIT_SAMPLE, _format_money_ru(56_330_456.66), "56 330 456,66"):
        text_w = pdfmetrics.stringWidth(sample, FONT_NORMAL, _TABLE_BODY_SIZE)
        assert text_w <= usable, f"{sample!r} width {text_w} > usable {usable} (col={col})"


def test_pile_layout_name_width_positive_on_a4() -> None:
    """Fixed pile columns + money cols leave room for Наименование on A4."""
    content_width = A4[0] - 2 * (10 * mm)
    no_width = 10 * mm
    grade_width = 32 * mm
    qty_width = 14 * mm
    money = _money_col_width()
    fixed = no_width + grade_width + qty_width + money + money
    name_width = content_width - fixed
    assert name_width > 0
    assert fixed + name_width == content_width


def test_pdf_generate_pile_order_with_large_sums() -> None:
    order = [
        {
            "mark": "C110.30-8",
            "name": "C110.30-8",
            "qty": 683,
            "concrete_grade": "B30_granite",
            "unit_price": 22317.40,
            "product_type": "piles",
            "product_kind": "pile",
        },
        {
            "mark": "C90.30-6",
            "name": "C90.30-6",
            "qty": 4946,
            "concrete_grade": "B30_granite",
            "unit_price": 11386.79,
            "product_type": "piles",
            "product_kind": "pile",
        },
    ]
    buf = generate_commercial_offer_pdf(
        order_data=order,
        offer_number="COLFIT",
        offer_date="02.09.2026",
        customer_name='ООО «СК «БЛОК»',
        logistics_cost=0.0,
    )
    data = buf.getvalue()
    assert data[:4] == b"%PDF"
    assert len(data) > 100
