#!/usr/bin/env python3
"""Convert docs/specs/1c-integration-tz.md to PDF (reportlab + Windows Arial)."""

from __future__ import annotations

import re
import sys
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import HRFlowable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SRC = ROOT / "docs" / "specs" / "1c-integration-tz.md"
DEFAULT_DST = ROOT / "docs" / "specs" / "1c-integration-tz.pdf"
FONT_REG = "Arial"
FONT_BOLD = "Arial-Bold"
FONT_MONO = "Consolas"


def _register_fonts() -> None:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.pdfmetrics import registerFontFamily
    from reportlab.pdfbase.ttfonts import TTFont

    fonts_dir = Path("C:/Windows/Fonts")
    pdfmetrics.registerFont(TTFont(FONT_REG, str(fonts_dir / "arial.ttf")))
    pdfmetrics.registerFont(TTFont(FONT_BOLD, str(fonts_dir / "arialbd.ttf")))
    pdfmetrics.registerFont(TTFont(FONT_MONO, str(fonts_dir / "consola.ttf")))
    registerFontFamily(
        FONT_REG,
        normal=FONT_REG,
        bold=FONT_BOLD,
        italic=FONT_REG,
        boldItalic=FONT_BOLD,
    )


def _styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "title",
            parent=base["Title"],
            fontName=FONT_BOLD,
            fontSize=16,
            leading=20,
            alignment=TA_CENTER,
            spaceAfter=6,
        ),
        "subtitle": ParagraphStyle(
            "subtitle",
            parent=base["Normal"],
            fontName=FONT_BOLD,
            fontSize=12,
            leading=15,
            alignment=TA_CENTER,
            spaceAfter=12,
        ),
        "meta": ParagraphStyle(
            "meta",
            parent=base["Normal"],
            fontName=FONT_REG,
            fontSize=10,
            leading=13,
            spaceAfter=3,
        ),
        "h2": ParagraphStyle(
            "h2",
            parent=base["Heading2"],
            fontName=FONT_BOLD,
            fontSize=12,
            leading=15,
            spaceBefore=10,
            spaceAfter=6,
        ),
        "h3": ParagraphStyle(
            "h3",
            parent=base["Heading3"],
            fontName=FONT_BOLD,
            fontSize=11,
            leading=14,
            spaceBefore=8,
            spaceAfter=4,
        ),
        "body": ParagraphStyle(
            "body",
            parent=base["Normal"],
            fontName=FONT_REG,
            fontSize=10,
            leading=14,
            alignment=TA_JUSTIFY,
            spaceAfter=4,
        ),
        "bullet": ParagraphStyle(
            "bullet",
            parent=base["Normal"],
            fontName=FONT_REG,
            fontSize=10,
            leading=14,
            leftIndent=14,
            bulletIndent=0,
            spaceAfter=2,
        ),
        "footer": ParagraphStyle(
            "footer",
            parent=base["Normal"],
            fontName=FONT_REG,
            fontSize=9,
            leading=12,
            textColor=colors.grey,
            alignment=TA_LEFT,
            spaceBefore=12,
        ),
    }


def _escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _inline_md(text: str) -> str:
    text = _escape(text.strip())
    text = re.sub(r"\*\*(.+?)\*\*", rf"<b>\1</b>", text)
    text = re.sub(r"`(.+?)`", rf'<font name="{FONT_MONO}">\1</font>', text)
    text = re.sub(r"\*(.+?)\*", rf"<i>\1</i>", text)
    return text


def _parse_table(lines: list[str], start: int) -> tuple[Table | None, int]:
    header = [c.strip() for c in lines[start].strip().strip("|").split("|")]
    if start + 1 >= len(lines) or not re.match(r"^\|?[\s\-:|]+\|?$", lines[start + 1].strip()):
        return None, start
    rows = [header]
    i = start + 2
    while i < len(lines) and lines[i].strip().startswith("|"):
        rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
        i += 1
    data = [[Paragraph(_inline_md(c), _styles()["body"]) for c in row] for row in rows]
    table = Table(data, colWidths=[5.5 * cm, 10.5 * cm], hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD),
                ("FONTNAME", (0, 1), (-1, -1), FONT_REG),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f0f0f0")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return table, i


def build_pdf(src: Path, dst: Path) -> None:
    _register_fonts()
    st = _styles()
    lines = src.read_text(encoding="utf-8").splitlines()
    story: list = []

    i = 0
    while i < len(lines):
        raw = lines[i].rstrip()
        stripped = raw.strip()

        if not stripped:
            i += 1
            continue

        if stripped == "---":
            story.append(Spacer(1, 4))
            story.append(HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey))
            story.append(Spacer(1, 4))
            i += 1
            continue

        if stripped.startswith("# "):
            story.append(Paragraph(_inline_md(stripped[2:]), st["title"]))
            i += 1
            continue

        if stripped.startswith("## "):
            story.append(Paragraph(_inline_md(stripped[3:]), st["subtitle"] if not story else st["h2"]))
            i += 1
            continue

        if stripped.startswith("### "):
            story.append(Paragraph(_inline_md(stripped[4:]), st["h3"]))
            i += 1
            continue

        if stripped.startswith("|"):
            table, i = _parse_table(lines, i)
            if table:
                story.append(Spacer(1, 4))
                story.append(table)
                story.append(Spacer(1, 6))
                continue

        if stripped.startswith("- "):
            story.append(Paragraph(f"• {_inline_md(stripped[2:])}", st["bullet"]))
            i += 1
            continue

        if re.match(r"^\d+\.\s", stripped):
            story.append(Paragraph(_inline_md(stripped), st["bullet"]))
            i += 1
            continue

        if stripped.startswith("*") and stripped.endswith("*"):
            story.append(Paragraph(_inline_md(stripped.strip("*")), st["footer"]))
            i += 1
            continue

        story.append(Paragraph(_inline_md(stripped), st["body"]))
        i += 1

    dst.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(dst),
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
        title="Техническое задание — интеграция 1С",
        author="Шишов",
    )
    doc.build(story)


def main() -> int:
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SRC
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_DST
    if not src.is_file():
        print(f"Source not found: {src}", file=sys.stderr)
        return 1
    build_pdf(src, dst)
    print(f"Written: {dst}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
