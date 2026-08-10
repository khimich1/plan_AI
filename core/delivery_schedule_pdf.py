#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF-документ графика поставки (reportlab, стиль как commercial_offer).

Макет MVP (R5): шапка + таблица партий/позиций, не этажи-колонки ЯРПРОФИТ.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from core.commercial_offer import FONT_BOLD, FONT_NORMAL
from core.delivery_schedule_xlsx import (
    DOC_HEADERS,
    document_header_lines,
    iter_document_table_rows,
    schedule_as_dict,
)

logger = logging.getLogger(__name__)


def build_document(schedule_view_or_dict: Any, path: str | Path) -> Path:
    """Собирает PDF-документ графика поставки (шапка + таблица партий).

    ``schedule_view_or_dict`` — ``DeliveryScheduleView`` или dict; опционально
    ключ ``customer_name`` из КП для шапки.
    """
    data = schedule_as_dict(schedule_view_or_dict)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        str(path),
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=15 * mm,
    )
    styles = getSampleStyleSheet()
    style_title = ParagraphStyle(
        "DeliveryScheduleTitle",
        parent=styles["Normal"],
        fontName=FONT_BOLD,
        fontSize=14,
        leading=18,
        spaceAfter=3 * mm,
    )
    style_meta = ParagraphStyle(
        "DeliveryScheduleMeta",
        parent=styles["Normal"],
        fontName=FONT_NORMAL,
        fontSize=10,
        leading=13,
        spaceAfter=1 * mm,
    )
    style_cell = ParagraphStyle(
        "DeliveryScheduleCell",
        parent=styles["Normal"],
        fontName=FONT_NORMAL,
        fontSize=9,
        leading=11,
    )
    style_cell_bold = ParagraphStyle(
        "DeliveryScheduleCellBold",
        parent=style_cell,
        fontName=FONT_BOLD,
    )

    story: list[Any] = []
    header_lines = document_header_lines(data)
    if header_lines:
        story.append(Paragraph(escape(header_lines[0]), style_title))
        for line in header_lines[1:]:
            story.append(Paragraph(escape(line), style_meta))
    story.append(Spacer(1, 4 * mm))

    table_data: list[list[Any]] = [
        [Paragraph(escape(h), style_cell_bold) for h in DOC_HEADERS]
    ]
    for values in iter_document_table_rows(data):
        table_data.append(
            [
                Paragraph(escape(str(v if v is not None else "")), style_cell)
                for v in values
            ]
        )

    col_widths = [
        12 * mm,
        40 * mm,
        24 * mm,
        24 * mm,
        24 * mm,
        40 * mm,
        18 * mm,
    ]
    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.Color(0.92, 0.92, 0.92)),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(table)
    doc.build(story)
    return path
