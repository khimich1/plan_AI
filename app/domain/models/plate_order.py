from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


LoadKey = tuple[float, float, int | float, str]
ExactWidthKey = tuple[float, str]


@dataclass
class PlateOrder:
    plates_1_2: list[float] = field(default_factory=list)
    plates_1_5_to_1_2: list[float] = field(default_factory=list)
    plates_1_0: list[float] = field(default_factory=list)
    plates_1_08: list[float] = field(default_factory=list)
    plates_0_46: list[float] = field(default_factory=list)
    plates_0_32: list[float] = field(default_factory=list)
    plates_0_72: list[float] = field(default_factory=list)
    plates_0_70: list[float] = field(default_factory=list)
    plates_0_86: list[float] = field(default_factory=list)
    plates_0_74: list[float] = field(default_factory=list)
    plates_0_88: list[float] = field(default_factory=list)
    plates_0_48: list[float] = field(default_factory=list)
    plates_0_50: list[float] = field(default_factory=list)
    plates_0_34: list[float] = field(default_factory=list)
    plate_load_details: dict[LoadKey, int] = field(default_factory=dict)
    plate_length_dm_raw: dict[LoadKey, str] = field(default_factory=dict)
    plate_exact_widths: dict[ExactWidthKey, float] = field(default_factory=dict)
    nomenclature_cache: dict[LoadKey, dict[str, Any]] = field(default_factory=dict)
    longitudinal_cuts: int = 0
    length_trims: int = 0
    unused_strips_0_3_m_total: float = 0.0
    scrap_strips_0_2_m_total: float = 0.0
    usable_strips_0_74_m_total: float = 0.0
    usable_strips_0_88_m_total: float = 0.0
    usable_strips_0_48_m_total: float = 0.0
    usable_strips_0_50_m_total: float = 0.0
    usable_strips_0_34_m_total: float = 0.0
    scrap_strips_0_12_m_total: float = 0.0
    waste_area_m2: float = 0.0

    def to_dict(self) -> dict[str, Any]:
        return {
            "plates_1_2": list(self.plates_1_2),
            "plates_1_5_to_1_2": list(self.plates_1_5_to_1_2),
            "plates_1_0": list(self.plates_1_0),
            "plates_1_08": list(self.plates_1_08),
            "plates_0_46": list(self.plates_0_46),
            "plates_0_32": list(self.plates_0_32),
            "plates_0_72": list(self.plates_0_72),
            "plates_0_70": list(self.plates_0_70),
            "plates_0_86": list(self.plates_0_86),
            "plates_0_74": list(self.plates_0_74),
            "plates_0_88": list(self.plates_0_88),
            "plates_0_48": list(self.plates_0_48),
            "plates_0_50": list(self.plates_0_50),
            "plates_0_34": list(self.plates_0_34),
            "plate_load_details": [[list(k), v] for k, v in self.plate_load_details.items()],
            "plate_length_dm_raw": [[list(k), v] for k, v in self.plate_length_dm_raw.items()],
            "plate_exact_widths": [[list(k), v] for k, v in self.plate_exact_widths.items()],
            "nomenclature_cache": [[list(k), v] for k, v in self.nomenclature_cache.items()],
            "longitudinal_cuts": self.longitudinal_cuts,
            "length_trims": self.length_trims,
            "unused_strips_0_3_m_total": self.unused_strips_0_3_m_total,
            "scrap_strips_0_2_m_total": self.scrap_strips_0_2_m_total,
            "usable_strips_0_74_m_total": self.usable_strips_0_74_m_total,
            "usable_strips_0_88_m_total": self.usable_strips_0_88_m_total,
            "usable_strips_0_48_m_total": self.usable_strips_0_48_m_total,
            "usable_strips_0_50_m_total": self.usable_strips_0_50_m_total,
            "usable_strips_0_34_m_total": self.usable_strips_0_34_m_total,
            "scrap_strips_0_12_m_total": self.scrap_strips_0_12_m_total,
            "waste_area_m2": self.waste_area_m2,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "PlateOrder":
        def _parse_load_key(items: list[Any]) -> LoadKey:
            raw = str(items[3]).strip() if len(items) > 3 and items[3] is not None else ""
            load_value = items[2]
            if isinstance(load_value, (int, float)):
                normalized_load: int | float = float(load_value)
            else:
                normalized_load = int(str(load_value))
            return (float(items[0]), float(items[1]), normalized_load, raw)

        return cls(
            plates_1_2=list(data.get("plates_1_2", [])),
            plates_1_5_to_1_2=list(data.get("plates_1_5_to_1_2", [])),
            plates_1_0=list(data.get("plates_1_0", [])),
            plates_1_08=list(data.get("plates_1_08", [])),
            plates_0_46=list(data.get("plates_0_46", [])),
            plates_0_32=list(data.get("plates_0_32", [])),
            plates_0_72=list(data.get("plates_0_72", [])),
            plates_0_70=list(data.get("plates_0_70", [])),
            plates_0_86=list(data.get("plates_0_86", [])),
            plates_0_74=list(data.get("plates_0_74", [])),
            plates_0_88=list(data.get("plates_0_88", [])),
            plates_0_48=list(data.get("plates_0_48", [])),
            plates_0_50=list(data.get("plates_0_50", [])),
            plates_0_34=list(data.get("plates_0_34", [])),
            plate_load_details={_parse_load_key(k): int(v) for k, v in data.get("plate_load_details", [])},
            plate_length_dm_raw={_parse_load_key(k): str(v) for k, v in data.get("plate_length_dm_raw", [])},
            plate_exact_widths={(float(k[0]), str(k[1])): float(v) for k, v in data.get("plate_exact_widths", [])},
            nomenclature_cache={_parse_load_key(k): dict(v) for k, v in data.get("nomenclature_cache", [])},
            longitudinal_cuts=int(data.get("longitudinal_cuts", 0)),
            length_trims=int(data.get("length_trims", 0)),
            unused_strips_0_3_m_total=float(data.get("unused_strips_0_3_m_total", 0)),
            scrap_strips_0_2_m_total=float(data.get("scrap_strips_0_2_m_total", 0)),
            usable_strips_0_74_m_total=float(data.get("usable_strips_0_74_m_total", 0)),
            usable_strips_0_88_m_total=float(data.get("usable_strips_0_88_m_total", 0)),
            usable_strips_0_48_m_total=float(data.get("usable_strips_0_48_m_total", 0)),
            usable_strips_0_50_m_total=float(data.get("usable_strips_0_50_m_total", 0)),
            usable_strips_0_34_m_total=float(data.get("usable_strips_0_34_m_total", 0)),
            scrap_strips_0_12_m_total=float(data.get("scrap_strips_0_12_m_total", 0)),
            waste_area_m2=float(data.get("waste_area_m2", 0)),
        )

    @classmethod
    def from_legacy(cls, legacy_order: Any) -> "PlateOrder":
        if isinstance(legacy_order, cls):
            return legacy_order
        if hasattr(legacy_order, "to_dict"):
            return cls.from_dict(legacy_order.to_dict())
        raise TypeError("Unsupported legacy plate order")

    @classmethod
    def from_orders_2d(cls, orders_2d: list[dict[str, Any]]) -> "PlateOrder":
        order = cls()
        width_to_list = {
            1200: order.plates_1_2,
            1080: order.plates_1_08,
            1000: order.plates_1_0,
            320: order.plates_0_32,
            460: order.plates_0_46,
            700: order.plates_0_70,
            720: order.plates_0_72,
            860: order.plates_0_86,
            880: order.plates_0_88,
            740: order.plates_0_74,
            480: order.plates_0_48,
            500: order.plates_0_50,
            340: order.plates_0_34,
        }
        for plate in orders_2d:
            length = float(plate["length"])
            width_mm = int(plate["width"])
            width_m = width_mm / 1000.0
            qty = int(plate.get("qty", 1))
            load_code = plate.get("load_code", 8)
            raw = (plate.get("length_dm_raw") or "").strip()
            key: LoadKey = (round(length, 3), round(width_m, 3), float(load_code), raw)
            order.plate_load_details[key] = order.plate_load_details.get(key, 0) + qty
            order.plate_length_dm_raw[key] = raw
            target = width_to_list.get(width_mm)
            if target is None and abs(width_m - 1.2) < 0.01:
                target = order.plates_1_2
            if target is not None:
                target.extend([length] * qty)
        order.recompute_totals()
        return order

    def to_orders_2d(self) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        for key, qty in self.plate_load_details.items():
            result.append(
                {
                    "length": key[0],
                    "width": int(round(key[1] * 1000)),
                    "qty": int(qty),
                    "load_code": key[2],
                    "length_dm_raw": key[3],
                }
            )
        return result

    def recompute_totals(self) -> None:
        self.longitudinal_cuts = (
            len(self.plates_1_5_to_1_2)
            + len(self.plates_1_0)
            + len(self.plates_1_08)
            + len(self.plates_0_46)
            + len(self.plates_0_32)
            + len(self.plates_0_72)
            + len(self.plates_0_70)
            + len(self.plates_0_86)
        )
        self.length_trims = 0
        self.unused_strips_0_3_m_total = 0.0
        self.scrap_strips_0_2_m_total = 0.0
        self.usable_strips_0_74_m_total = round(sum(self.plates_0_46), 1)
        self.usable_strips_0_88_m_total = round(sum(self.plates_0_32), 1)
        self.usable_strips_0_48_m_total = round(sum(self.plates_0_72), 1)
        self.usable_strips_0_50_m_total = round(sum(self.plates_0_70), 1)
        self.usable_strips_0_34_m_total = round(sum(self.plates_0_86), 1)
        self.scrap_strips_0_12_m_total = round(sum(self.plates_1_08), 1)
        self.waste_area_m2 = round(0.12 * self.scrap_strips_0_12_m_total, 2)

