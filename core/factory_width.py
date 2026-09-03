"""Factory cut-width ranges and plate-mark rewrite for commercial offers."""

from __future__ import annotations

import re

FACTORY_WIDTH_RANGES_MM: tuple[tuple[int, int], ...] = (
    (260, 320),
    (460, 530),
    (660, 720),
    (860, 920),
    (1020, 1080),
    (1200, 1200),
)

# «п» optional: user input often omits it («68-11-10 1»), display names keep it.
_MARK_WIDTH_RE = re.compile(
    r"(?P<head>.*?)(?P<length>\d+(?:[.,]\d+)?)-(?P<width>\d+(?:[.,]\d+)?)-"
    r"(?P<load>\d+(?:[.,]\d+)?)(?P<pe>\s*п)?(?P<tail>.*)",
    re.IGNORECASE | re.DOTALL,
)


def width_m_to_mm(width_m: float) -> int:
    return int(round(float(width_m) * 1000))


def is_factory_width_mm(width_mm: int) -> bool:
    return any(lo <= width_mm <= hi for lo, hi in FACTORY_WIDTH_RANGES_MM)


def suggest_factory_width_mm(width_mm: int) -> list[int]:
    if is_factory_width_mm(width_mm):
        return []
    lower = [hi for lo, hi in FACTORY_WIDTH_RANGES_MM if hi < width_mm]
    upper = [lo for lo, hi in FACTORY_WIDTH_RANGES_MM if lo > width_mm]
    out: list[int] = []
    if lower:
        out.append(max(lower))
    if upper:
        out.append(min(upper))
    return out


def format_factory_width_label(width_mm: int) -> str:
    """Format mm as the W token in L-W-N (720 → «7,2», 1200 → «12»)."""
    width_m = width_mm / 1000.0
    width_dm = round(width_m * 10, 2)
    if abs(width_dm - round(width_dm)) < 1e-6:
        return str(int(round(width_dm)))
    return f"{width_dm:.2f}".rstrip("0").rstrip(".").replace(".", ",")


def rewrite_plate_line_width(line: str, new_width_mm: int) -> str:
    """Replace only the W part of L-W-N. Qty, load, and prefix stay as-is."""
    text = str(line or "")
    match = _MARK_WIDTH_RE.fullmatch(text)
    if not match:
        raise ValueError(f"Не удалось найти ширину в марке: {line!r}")
    label = format_factory_width_label(new_width_mm)
    pe = match.group("pe") or ""
    return (
        f"{match.group('head')}{match.group('length')}-{label}-"
        f"{match.group('load')}{pe}{match.group('tail')}"
    )
