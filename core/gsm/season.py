"""Season mode for GSM fuel norms: manual switch journal, single source of truth.

The accountant switches the season explicitly (button → POST settings/season);
each switch is appended to the ``season_switches`` setting as
``[{"date": "YYYY-MM-DD", "mode": "winter"|"summer"}]`` sorted by date.
There is no calendar heuristic: before the first switch everything is summer.
"""

from __future__ import annotations

import json
from collections.abc import Sequence
from datetime import date

SEASON_MODES = frozenset({"summer", "winter"})
DEFAULT_SEASON_MODE = "summer"

SeasonSwitch = tuple[date, str]


def season_mode_for(day: date, switches: Sequence[SeasonSwitch]) -> str:
    """Mode of the latest switch with ``date <= day``; summer when none.

    ``switches`` must be sorted by date ascending (see ``parse_season_switches``);
    on equal dates the later entry wins.
    """
    mode = DEFAULT_SEASON_MODE
    for switch_date, switch_mode in switches:
        if switch_date > day:
            break
        mode = switch_mode
    return mode


def norm_for(
    day: date,
    *,
    norm_summer: float,
    norm_winter: float,
    switches: Sequence[SeasonSwitch],
) -> float:
    """Winter norm iff ``season_mode_for(day, switches) == "winter"``."""
    return norm_winter if season_mode_for(day, switches) == "winter" else norm_summer


def parse_season_switches(raw: str | None) -> tuple[SeasonSwitch, ...]:
    """Parse the ``season_switches`` setting JSON into date-sorted tuples.

    Raises ``ValueError`` on invalid JSON or malformed items; callers wrap it
    in their domain error (``gsm_settings_invalid``).
    """
    if raw is None or raw == "":
        return ()
    try:
        parsed = json.loads(raw)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"season_switches is not valid JSON: {raw!r}") from exc
    if not isinstance(parsed, list):
        raise ValueError("season_switches must be a JSON array")
    switches: list[SeasonSwitch] = []
    for item in parsed:
        if not isinstance(item, dict):
            raise ValueError(f"season_switches item must be an object: {item!r}")
        try:
            day = date.fromisoformat(str(item["date"]))
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"season_switches item has invalid date: {item!r}") from exc
        mode = str(item.get("mode"))
        if mode not in SEASON_MODES:
            raise ValueError(f"season_switches item has invalid mode: {item!r}")
        switches.append((day, mode))
    switches.sort(key=lambda s: s[0])
    return tuple(switches)


def serialize_season_switches(switches: Sequence[SeasonSwitch]) -> str:
    """Inverse of ``parse_season_switches`` (kept sorted by date)."""
    ordered = sorted(switches, key=lambda s: s[0])
    payload = [{"date": d.isoformat(), "mode": m} for d, m in ordered]
    return json.dumps(payload, ensure_ascii=False)
