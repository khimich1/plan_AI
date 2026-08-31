"""Unit tests for core.gsm.geo (pure sphere math, no I/O).

TDD / Task 2 orch-2026-08-15-gsm-geo-lookahead: these tests are written
FIRST and must fail until ``core/gsm/geo.py`` exists.

Public surface pinned by this file
-----------------------------------
``core.gsm.geo``
    frozen+slots DTO: ``GeoPoint(lat, lon)``.
    ``haversine_km(a, b) -> float``
        sphere distance, R=6371.0, ``round(..., 2)``.
    ``bearing_deg(a, b) -> float``
        azimuth a→b in degrees ``[0, 360)``.
    ``angle_diff_deg(x, y) -> float``
        minimal azimuth difference in ``[0, 180]``.
    ``point_to_segment_km(point, a, b) -> float``
        shortest distance from point to segment A–B, km
        (local equidistant projection, GeoPoint lat/lon API).
"""

from __future__ import annotations

import ast
import dataclasses
import math
from pathlib import Path

import pytest

from core.gsm.geo import (
    GeoPoint,
    angle_diff_deg,
    bearing_deg,
    haversine_km,
    point_to_segment_km,
)

REPO_ROOT = Path(__file__).resolve().parents[1]
GSM_GEO_PATH = REPO_ROOT / "core" / "gsm" / "geo.py"

# Control points from spec / orchestration brief (Kostroma Kuznetskaya / centre).
KOSTROMA = GeoPoint(lat=57.766, lon=40.927)
YAROSLAVL = GeoPoint(lat=57.630, lon=39.870)
MOSCOW = GeoPoint(lat=55.755, lon=37.617)

KOSTROMA_YAROSLAVL_KM = 63.5
KOSTROMA_YAROSLAVL_BEARING = 257.0
KOSTROMA_MOSCOW_KM = 300.0
KOSTROMA_MOSCOW_BEARING = 224.0

# ±5% of 257° ≈ 12.85°; brief also allows ±13°.
BEARING_ABS_TOL = 13.0


# ---------------------------------------------------------------------------
# 1. GeoPoint DTO
# ---------------------------------------------------------------------------


def test_geo_point_is_frozen_slots_dataclass() -> None:
    assert dataclasses.is_dataclass(GeoPoint)
    params = getattr(GeoPoint, "__dataclass_params__", None)
    assert params is not None and params.frozen
    assert hasattr(GeoPoint, "__slots__")
    fields = {f.name for f in dataclasses.fields(GeoPoint)}
    assert fields == {"lat", "lon"}


def test_geo_point_is_immutable() -> None:
    with pytest.raises(dataclasses.FrozenInstanceError):
        KOSTROMA.lat = 0.0  # type: ignore[misc]


# ---------------------------------------------------------------------------
# 2. haversine / bearing — control pairs (±5%)
# ---------------------------------------------------------------------------


def test_haversine_kostroma_to_yaroslavl() -> None:
    dist = haversine_km(KOSTROMA, YAROSLAVL)
    assert dist == pytest.approx(KOSTROMA_YAROSLAVL_KM, rel=0.05)


def test_bearing_kostroma_to_yaroslavl() -> None:
    az = bearing_deg(KOSTROMA, YAROSLAVL)
    assert 0.0 <= az < 360.0
    assert az == pytest.approx(KOSTROMA_YAROSLAVL_BEARING, rel=0.05, abs=BEARING_ABS_TOL)


def test_haversine_kostroma_to_moscow() -> None:
    dist = haversine_km(KOSTROMA, MOSCOW)
    assert dist == pytest.approx(KOSTROMA_MOSCOW_KM, rel=0.05)


def test_bearing_kostroma_to_moscow() -> None:
    az = bearing_deg(KOSTROMA, MOSCOW)
    assert 0.0 <= az < 360.0
    assert az == pytest.approx(KOSTROMA_MOSCOW_BEARING, rel=0.05, abs=BEARING_ABS_TOL)


def test_haversine_rounds_to_two_decimals() -> None:
    dist = haversine_km(KOSTROMA, YAROSLAVL)
    assert dist == round(dist, 2)


def test_haversine_same_point_is_zero() -> None:
    assert haversine_km(KOSTROMA, KOSTROMA) == 0.0


def test_haversine_is_symmetric() -> None:
    assert haversine_km(KOSTROMA, YAROSLAVL) == haversine_km(YAROSLAVL, KOSTROMA)
    assert haversine_km(KOSTROMA, MOSCOW) == haversine_km(MOSCOW, KOSTROMA)


def test_bearing_reverse_is_about_plus_180() -> None:
    forward = bearing_deg(KOSTROMA, YAROSLAVL)
    reverse = bearing_deg(YAROSLAVL, KOSTROMA)
    assert angle_diff_deg(forward, reverse) == pytest.approx(180.0, abs=1.0)


# ---------------------------------------------------------------------------
# 3. angle_diff_deg — wrap-around and symmetry
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("x", "y", "expected"),
    [
        (0.0, 360.0, 0.0),
        (360.0, 0.0, 0.0),
        (179.0, 181.0, 2.0),
        (181.0, 179.0, 2.0),
        (10.0, 350.0, 20.0),
        (350.0, 10.0, 20.0),
        (0.0, 180.0, 180.0),
        (90.0, 90.0, 0.0),
    ],
)
def test_angle_diff_deg_boundaries(x: float, y: float, expected: float) -> None:
    assert angle_diff_deg(x, y) == pytest.approx(expected, abs=1e-9)


def test_angle_diff_deg_is_symmetric() -> None:
    pairs = ((0.0, 360.0), (179.0, 181.0), (10.0, 350.0), (45.0, 200.0))
    for x, y in pairs:
        assert angle_diff_deg(x, y) == angle_diff_deg(y, x)


def test_angle_diff_deg_range() -> None:
    for x, y in ((0.0, 359.0), (1.0, 180.0), (270.0, 90.0), (359.0, 1.0)):
        diff = angle_diff_deg(x, y)
        assert 0.0 <= diff <= 180.0


# ---------------------------------------------------------------------------
# 4. point_to_segment_km — on-segment vs far aside
# ---------------------------------------------------------------------------


def test_point_to_segment_endpoint_is_zero() -> None:
    assert point_to_segment_km(KOSTROMA, KOSTROMA, YAROSLAVL) == pytest.approx(0.0, abs=0.01)
    assert point_to_segment_km(YAROSLAVL, KOSTROMA, YAROSLAVL) == pytest.approx(0.0, abs=0.01)


def test_point_to_segment_midpoint_is_near_zero() -> None:
    mid = GeoPoint(
        lat=(KOSTROMA.lat + YAROSLAVL.lat) / 2.0,
        lon=(KOSTROMA.lon + YAROSLAVL.lon) / 2.0,
    )
    assert point_to_segment_km(mid, KOSTROMA, YAROSLAVL) < 1.0


def test_point_to_segment_far_aside_is_large() -> None:
    """Moscow is far off the Kostroma–Yaroslavl chord (~63 km west)."""
    dist = point_to_segment_km(MOSCOW, KOSTROMA, YAROSLAVL)
    assert dist > 200.0


def test_point_to_segment_degenerate_equals_point_distance() -> None:
    """A==B: distance is the projected offset of the point from A."""
    dist = point_to_segment_km(YAROSLAVL, KOSTROMA, KOSTROMA)
    assert dist == pytest.approx(haversine_km(YAROSLAVL, KOSTROMA), rel=0.02)


def test_point_to_segment_beyond_endpoint_snaps_to_vertex() -> None:
    """Point past Yaroslavl along the same heading is closest to Yaroslavl."""
    # ~1° further west of Yaroslavl (Kostroma→Yaroslavl is westward).
    beyond = GeoPoint(lat=YAROSLAVL.lat, lon=YAROSLAVL.lon - 1.0)
    dist = point_to_segment_km(beyond, KOSTROMA, YAROSLAVL)
    assert dist == pytest.approx(haversine_km(beyond, YAROSLAVL), rel=0.05)


# ---------------------------------------------------------------------------
# 5. Purity: no sqlite / requests / app / pathlib I/O
# ---------------------------------------------------------------------------


_FORBIDDEN_ROOTS = frozenset(
    {
        "sqlite3",
        "requests",
        "app",
        "pathlib",
        "httpx",
        "urllib",
        "urllib3",
        "aiohttp",
        "socket",
    }
)


def _imported_roots(tree: ast.AST) -> set[str]:
    roots: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                roots.add(alias.name.split(".", 1)[0])
        elif isinstance(node, ast.ImportFrom):
            module = node.module or ""
            if module:
                roots.add(module.split(".", 1)[0])
    return roots


def test_geo_module_is_pure_no_io() -> None:
    assert GSM_GEO_PATH.is_file(), f"missing module file: {GSM_GEO_PATH}"
    source = GSM_GEO_PATH.read_text(encoding="utf-8")
    tree = ast.parse(source)
    roots = _imported_roots(tree)
    forbidden = roots & _FORBIDDEN_ROOTS
    assert not forbidden, f"core.gsm.geo must stay pure, got imports: {sorted(forbidden)}"
    assert "open(" not in source
    assert "scripts" not in roots
