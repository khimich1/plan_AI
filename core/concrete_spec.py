"""Заводская таблица морозостойкости и водонепроницаемости бетона."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class ConcreteSpec:
    """Гранитная пара таблицы и водонепроницаемость рядового столбца, если он есть."""

    frost_resistance: str
    waterproofness: str
    concrete_aggregate: str
    ordinary_waterproofness: str | None


# F, W гранита, W рядового (None — отдельного пункта нет).
_TABLE: dict[str, tuple[str, str, str | None]] = {
    "B7.5": ("F50", "W2", "W2"),
    "B12.5": ("F50", "W2", "W2"),
    "B15": ("F100", "W4", "W4"),
    "B20": ("F100", "W4", "W4"),
    "B22.5": ("F200", "W6", "W4"),
    "B25": ("F200", "W8", "W6"),
    "B30": ("F300", "W10", None),
    "M400": ("F300", "W10", None),
    "B35": ("F300", "W10", None),
    "M500": ("F300", "W12", None),
    "B40": ("F300", "W12", None),
    "B45": ("F300", "W12", None),
}


def _lookup_key(concrete_grade: str) -> str | None:
    text = str(concrete_grade).strip().replace("М", "M").replace("м", "M")
    if not text:
        return None
    key = text.upper().replace("_", ".")
    if key.endswith(".GRANITE"):
        key = key[: -len(".GRANITE")]
    return key


def suggest_concrete_spec(concrete_grade: str) -> ConcreteSpec | None:
    """Гранитная пара заводской таблицы. None, если марку или класс не узнали.

    B25 → F200 / W8 / granite
    B30_granite и B30 → F300 / W10 / granite
    М500 и M500 → F300 / W12 / granite
    """
    key = _lookup_key(concrete_grade)
    if key is None:
        return None
    row = _TABLE.get(key)
    if row is None:
        return None
    frost, granite_w, ordinary_w = row
    return ConcreteSpec(
        frost_resistance=frost,
        waterproofness=granite_w,
        concrete_aggregate="granite",
        ordinary_waterproofness=ordinary_w,
    )


_SPEC_FIELDS = (
    "frost_resistance",
    "waterproofness",
    "concrete_aggregate",
    "concrete_spec_source",
)

GOST_FROST_RESISTANCE = frozenset(
    {"F50", "F75", "F100", "F150", "F200", "F300", "F400", "F500", "F600", "F800", "F1000"}
)
GOST_WATERPROOFNESS = frozenset(
    {"W2", "W4", "W6", "W8", "W10", "W12", "W14", "W16", "W18", "W20"}
)


def _line_id(line: dict[str, Any]) -> str:
    return str(line.get("line_id") or "").strip()


def _snapshot_empty(line: dict[str, Any]) -> bool:
    """Старая строка без снимка: полей нет или все четыре NULL/пустые."""
    for key in _SPEC_FIELDS:
        value = line.get(key)
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return False
    return True


def _strip_spec(line: dict[str, Any]) -> None:
    for key in _SPEC_FIELDS:
        line.pop(key, None)


def _write_table_pair(line: dict[str, Any], aggregate: str) -> None:
    spec = suggest_concrete_spec(str(line.get("concrete_grade") or ""))
    if spec is None:
        _strip_spec(line)
        return
    use_ordinary = aggregate == "ordinary" and spec.ordinary_waterproofness is not None
    line["frost_resistance"] = spec.frost_resistance
    line["waterproofness"] = (
        spec.ordinary_waterproofness if use_ordinary else spec.waterproofness
    )
    line["concrete_aggregate"] = "ordinary" if use_ordinary else "granite"
    line["concrete_spec_source"] = "table"


def _copy_manual(line: dict[str, Any], previous: dict[str, Any]) -> None:
    line["frost_resistance"] = previous.get("frost_resistance")
    line["waterproofness"] = previous.get("waterproofness")
    aggregate = previous.get("concrete_aggregate")
    line["concrete_aggregate"] = "" if aggregate is None else aggregate
    line["concrete_spec_source"] = "manual"


def build_concrete_spec_update(
    *,
    concrete_grade: str,
    concrete_spec_source: str,
    concrete_aggregate: str | None = None,
    frost_resistance: str | None = None,
    waterproofness: str | None = None,
) -> dict[str, str]:
    """Четыре поля снимка для PATCH. Для table F и W берёт справочник, не клиент."""
    source = (concrete_spec_source or "").strip()
    if source == "manual":
        frost = (frost_resistance or "").strip()
        water = (waterproofness or "").strip()
        if frost not in GOST_FROST_RESISTANCE or water not in GOST_WATERPROOFNESS:
            raise ValueError(
                "Марка морозостойкости или водонепроницаемости вне списка ГОСТ."
            )
        return {
            "frost_resistance": frost,
            "waterproofness": water,
            "concrete_aggregate": "",
            "concrete_spec_source": "manual",
        }
    if source != "table":
        raise ValueError("Источник пары должен быть table или manual.")
    aggregate = (concrete_aggregate or "").strip()
    spec = suggest_concrete_spec(concrete_grade)
    if spec is None or aggregate not in {"granite", "ordinary"}:
        raise ValueError("Для этого бетона нет такой пары таблицы.")
    if aggregate == "ordinary":
        if spec.ordinary_waterproofness is None:
            raise ValueError("Для этого бетона нет рядовой пары таблицы.")
        waterproof = spec.ordinary_waterproofness
    else:
        waterproof = spec.waterproofness
    return {
        "frost_resistance": spec.frost_resistance,
        "waterproofness": waterproof,
        "concrete_aggregate": aggregate,
        "concrete_spec_source": "table",
    }


def prepare_plate_line(line: dict[str, Any]) -> dict[str, Any]:
    """Марка плиты: явное значение не затирается, иначе resolve_concrete_grade.

    Слепой запасной М400 не считается основанием для подсказки F/W.
    """
    from core.concrete_grade_resolver import (
        _pb_number_fallback,
        normalize_concrete_grade,
        resolve_concrete_grade_from_order,
    )

    filled = dict(line)
    name = str(filled.get("name") or filled.get("plate_name") or filled.get("mark") or "")
    length_raw = filled.get("length_m")
    try:
        length_m = float(length_raw) if length_raw not in (None, "") else None
    except (TypeError, ValueError):
        length_m = None
    explicit = normalize_concrete_grade(filled.get("concrete_grade"))
    evidenced = explicit is not None or _pb_number_fallback(name, length_m) is not None
    if explicit:
        filled["concrete_grade"] = explicit
    else:
        filled["concrete_grade"] = resolve_concrete_grade_from_order(
            {
                "concrete_grade": None,
                "plate_name": name,
                "length": length_m,
                "load_code": filled.get("load_class", 800),
            }
        )
    filled["_concrete_grade_evidenced"] = evidenced
    return filled


def suppress_unevidenced_plate_specs(
    previous: list[dict[str, Any]],
    lines: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """У новой плиты без марки и номера ПБ не выдумывать F/W из запасного М400."""
    previous_by_id: dict[str, dict[str, Any]] = {}
    for raw in previous:
        if not isinstance(raw, dict):
            continue
        line_id = _line_id(raw)
        if line_id and line_id not in previous_by_id:
            previous_by_id[line_id] = raw

    cleaned: list[dict[str, Any]] = []
    for raw in lines:
        line = dict(raw)
        evidenced = bool(line.pop("_concrete_grade_evidenced", True))
        prior = previous_by_id.get(_line_id(line))
        if not evidenced and (prior is None or _snapshot_empty(prior)):
            _strip_spec(line)
        cleaned.append(line)
    return cleaned


def apply_line_specs(
    previous: list[dict[str, Any]],
    new: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Прикрепить снимок F/W к новым строкам по line_id.

    Нет пары — гранитная подсказка. Пустой снимок найденной строки не заполняется.
    Класс и цена новой строки не пересчитываются.
    """
    previous_by_id: dict[str, dict[str, Any]] = {}
    for raw in previous:
        if not isinstance(raw, dict):
            continue
        line_id = _line_id(raw)
        if line_id and line_id not in previous_by_id:
            previous_by_id[line_id] = raw

    attached: list[dict[str, Any]] = []
    for raw in new:
        line = dict(raw) if isinstance(raw, dict) else {}
        prior = previous_by_id.get(_line_id(line))
        if prior is None:
            _write_table_pair(line, "granite")
        elif _snapshot_empty(prior):
            _strip_spec(line)
        elif str(prior.get("concrete_spec_source") or "").strip() == "manual":
            _copy_manual(line, prior)
        else:
            aggregate = str(prior.get("concrete_aggregate") or "granite").strip() or "granite"
            _write_table_pair(line, aggregate)
        attached.append(line)
    return attached
