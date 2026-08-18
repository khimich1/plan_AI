"""Pure sphere geometry for GSM routes: distance, bearing, point-to-segment."""

from __future__ import annotations

from dataclasses import dataclass
from math import asin, atan2, cos, hypot, radians, sin, sqrt

_EARTH_RADIUS_KM = 6371.0
_EARTH_RADIUS_M = _EARTH_RADIUS_KM * 1000.0
_PI = 3.141592653589793


@dataclass(frozen=True, slots=True)
class GeoPoint:
    lat: float
    lon: float


def haversine_km(a: GeoPoint, b: GeoPoint) -> float:
    """Расстояние по сфере, км."""
    r = _EARTH_RADIUS_KM
    dlat = radians(b.lat - a.lat)
    dlon = radians(b.lon - a.lon)
    h = sin(dlat / 2) ** 2 + cos(radians(a.lat)) * cos(radians(b.lat)) * sin(dlon / 2) ** 2
    return round(2 * r * asin(sqrt(h)), 2)


def bearing_deg(a: GeoPoint, b: GeoPoint) -> float:
    """Азимут a→b в градусах [0, 360)."""
    dlon = radians(b.lon - a.lon)
    x = sin(dlon) * cos(radians(b.lat))
    y = (
        cos(radians(a.lat)) * sin(radians(b.lat))
        - sin(radians(a.lat)) * cos(radians(b.lat)) * cos(dlon)
    )
    return (atan2(x, y) * 180.0 / _PI + 360.0) % 360.0


def angle_diff_deg(x: float, y: float) -> float:
    """Минимальная разница двух азимутов [0, 180]."""
    d = abs(x - y) % 360.0
    return min(d, 360.0 - d)


def _local_xy_m(point: GeoPoint, origin: GeoPoint) -> tuple[float, float]:
    """Эквидистантная проекция в метры относительно origin."""
    x = radians(point.lon - origin.lon) * cos(radians(origin.lat)) * _EARTH_RADIUS_M
    y = radians(point.lat - origin.lat) * _EARTH_RADIUS_M
    return x, y


def point_to_segment_km(point: GeoPoint, a: GeoPoint, b: GeoPoint) -> float:
    """Кратчайшее расстояние от точки до отрезка A–B, км."""
    ax, ay = _local_xy_m(a, point)
    bx, by = _local_xy_m(b, point)
    abx, aby = bx - ax, by - ay
    ab2 = abx * abx + aby * aby
    if ab2 <= 0.0:
        dist_m = hypot(ax, ay)
    else:
        # P в начале координат; вектор A→P = (-ax, -ay)
        t = ((-ax) * abx + (-ay) * aby) / ab2
        t = max(0.0, min(1.0, t))
        cx = ax + t * abx
        cy = ay + t * aby
        dist_m = hypot(cx, cy)
    return dist_m / 1000.0
