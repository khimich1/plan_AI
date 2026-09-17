#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сборка Excel-очереди «Очередь синхронизации GUID и цен.xlsx».

Рабочая очередь вместо разового отчёта «Сопоставление GUID 1С и программа.xlsx»:
на каждом листе колонки «Состояние»/«Действие», лист «Задачи» — две очереди
(🏭 завести в 1С / 💰 ввести цену). Дубли 1С только помечаются, не разрешаются.

Запуск из корня репозитория:

    .venv/bin/python scripts/build_guid_queue.py
"""

from __future__ import annotations

import argparse
import re
import sqlite3
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional

import xlrd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.pile_catalog import parse_bridge_pile_geometry, parse_pile_mark
from core.project_paths import PRICE_DB_PATH

DEFAULT_GUID_DIR = PROJECT_ROOT / "банк знаний" / "Новая папка" / "GUID"
DEFAULT_OUT = DEFAULT_GUID_DIR / "Очередь синхронизации GUID и цен.xlsx"

STATE_READY = "✅ готово"
STATE_DUP = "⚠️ дубль, ждёт бухгалтерию"
STATE_CREATE = "🏭 завести в 1С"
STATE_PRICE = "💰 ввести цену"
STATE_OUT = "⏸️ вне контура"
STATE_MAT = "🚫 вне контура"
STATE_PLATE_ONLY = "только в 1С"

ACTION_READY = "—"
ACTION_DUP = "бухгалтерия выбирает GUID — см. «Дубли GUID в 1С.xlsx»"
ACTION_CREATE = "завести"
ACTION_CREATE_PLAIN = "завести только plain"
ACTION_PRICE = "ввести цену за шт"
ACTION_NONE = "—"
ACTION_PLATE_MATRIX = "не очередь: покрыто матрицей длина×нагрузка"
ACTION_PLATE_SHORT = "открытый вопрос: коротыш вне матрицы (0,58/13/15/15,8 дм)"

SHORT_PLATE_LENGTHS = {0.58, 13.0, 15.0, 15.8}

LAT2CYR = str.maketrans(
    {
        "c": "с",
        "C": "С",
        "t": "т",
        "T": "Т",
        "b": "в",
        "B": "В",
        "p": "р",
        "P": "Р",
        "a": "а",
        "A": "А",
        "e": "е",
        "E": "Е",
        "h": "н",
        "H": "Н",
        "m": "м",
        "M": "М",
        "k": "к",
        "K": "К",
        "x": "х",
        "X": "Х",
        "o": "о",
        "O": "О",
        "y": "у",
        "Y": "У",
    }
)

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF")
SECTION_CREATE_FILL = PatternFill("solid", fgColor="F4B183")
SECTION_PRICE_FILL = PatternFill("solid", fgColor="FFE699")
STATE_FILLS = {
    STATE_READY: PatternFill("solid", fgColor="C6EFCE"),
    STATE_DUP: PatternFill("solid", fgColor="FFEB9C"),
    STATE_CREATE: PatternFill("solid", fgColor="FCE4D6"),
    STATE_PRICE: PatternFill("solid", fgColor="FFF2CC"),
    STATE_OUT: PatternFill("solid", fgColor="D9D9D9"),
    STATE_MAT: PatternFill("solid", fgColor="D9D9D9"),
    STATE_PLATE_ONLY: PatternFill("solid", fgColor="DDEBF7"),
}
STATE_SORT = {
    STATE_CREATE: 0,
    STATE_PRICE: 1,
    STATE_DUP: 2,
    STATE_PLATE_ONLY: 3,
    STATE_READY: 4,
    STATE_OUT: 5,
    STATE_MAT: 6,
}

PRICE_FILE_NAMES = {
    "plity": "Прайс плиты.xls",
    "svai": "Прайс сваи.xls",
    "svai_s": "Прайс сваи составные.xls",
    "most": "Прайс сваи мостовые.xls",
    "most_s": "Прайс сваи мостовые составные.xls",
    "bloki": "Прайс блоки.xls",
    "lm": "Прайс ЛМ, ЛП, ЛС, Перемычки, Прогоны, Тротуарные плиты, Фундаменты.xls",
}

# Очередь GPS-001 / bootstrap GPS-003: не-плитные группы с прайсом в pb.db.
NON_PLATE_QUEUE_GROUPS = ("piles", "bridge", "fbs", "marches", "steps")
QUEUE_GROUP_TO_KIND = {
    "piles": "pile",
    "bridge": "bridge_pile",
    "fbs": "fbs",
    "marches": "stair_flight",
    "steps": "stair_step",
}

BR_RX = re.compile(r"^[Cc]\s*(\d+)[-\.](\d+)\s*([TТBВb])\s*(\d+)$")
STEP_RX = re.compile(r"^ЛС\s*(\d+)(.*)$", re.I)
PILE_TOKEN_RE = re.compile(
    r"([СC]\s*\d+(?:[.,]\d+)?\s*\.\s*\d+\s*-\s*\d+(?:[.,]\d+)?у?)",
    re.I,
)
BRIDGE_1C_RE = re.compile(
    r"^сваимостовыес(\d+)\.(\d+)-([тв])(\d+)$",
)
STEP_NUM_RE = re.compile(r"лс-?(\d+)(.*)$")
PLATE_LEN_RE = re.compile(r"^Плиты\s+ПБ\s*(\d+(?:[.,]\d+)?)", re.I)
GRADE_PREF = ("B25", "B20", "B22_5", "B30", "B15", "B30_granite", "B7_5")


@dataclass
class Item1C:
    guid: str
    name: str
    price: Optional[float]
    unit: str


@dataclass
class MatchResult:
    status: str  # ok | ambig | none | u_only
    item: Optional[Item1C] = None
    u_item: Optional[Item1C] = None
    note: str = ""
    candidates: list[Item1C] = field(default_factory=list)


@dataclass
class QueueRow:
    group_key: str
    group_title: str
    state: str
    action: str
    our_label: str = ""
    our_source: str = ""
    name_1c: str = ""
    guid: str = ""
    guid_u: str = ""
    name_u: str = ""
    price_1c: Optional[float] = None
    unit: str = ""
    neighbor_price: Optional[float] = None
    neighbor_label: str = ""
    note: str = ""
    suggested_1c_name: str = ""


@dataclass
class GeoPrice:
    section: object
    length: float
    mark: str
    price: float


def norm(s: str) -> str:
    text = str(s).strip().lower().replace("ё", "е").translate(LAT2CYR)
    text = text.replace("–", "-").replace("—", "-").replace(",", ".").replace("*", "")
    return re.sub(r"\s+", "", text)


def fmt_price(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return round(float(value), 2)


def load_price_file(path: Path) -> list[Item1C]:
    wb = xlrd.open_workbook(str(path))
    sh = wb.sheet_by_name("Прайс-лист")
    rows: list[Item1C] = []
    for r in range(2, sh.nrows):
        name = str(sh.cell_value(r, 3)).strip()
        if not name:
            continue
        price = sh.cell_value(r, 7)
        rows.append(
            Item1C(
                guid=str(sh.cell_value(r, 1)).strip().lower(),
                name=name,
                price=price if isinstance(price, (int, float)) and price > 0 else None,
                unit=str(sh.cell_value(r, 8)).strip(),
            )
        )
    return rows


def is_composite_pile(name: str) -> bool:
    n = norm(name)
    return "нсв" in n or "всв" in n


def make_set(files: dict[str, list[Item1C]], keys: Iterable[str], predicate=None) -> dict[str, Item1C]:
    out: dict[str, Item1C] = {}
    for key in keys:
        for item in files[key]:
            if predicate and not predicate(item.name):
                continue
            out.setdefault(item.guid, item)
    return out


def load_1c(guid_dir: Path) -> tuple[dict[str, dict[str, Item1C]], dict[str, dict[str, list[Item1C]]]]:
    files = {
        key: load_price_file(guid_dir / fname) for key, fname in PRICE_FILE_NAMES.items()
    }
    plity_all = files["plity"]
    plates_by_guid = {it.guid: it for it in plity_all}
    c1c: dict[str, dict[str, Item1C]] = {
        "plates": plates_by_guid,
        "piles": {
            **make_set(files, ["svai"]),
            **{
                it.guid: it
                for it in plity_all
                if it.name.startswith("Сваи") and not is_composite_pile(it.name)
            },
        },
        "piles_comp": {
            **make_set(files, ["svai_s"]),
            **{
                it.guid: it
                for it in plity_all
                if it.name.startswith("Сваи") and is_composite_pile(it.name)
            },
        },
        "bridge": make_set(files, ["most", "most_s"]),
        "fbs": make_set(files, ["bloki"]),
        "marches": make_set(files, ["lm"], lambda n: n.startswith("Лестничные марши")),
        "steps": make_set(files, ["lm"], lambda n: n.startswith("Лестничные ступени")),
        "pads": make_set(files, ["lm"], lambda n: n.startswith("Лестничные площадки")),
        "peremy": make_set(files, ["lm"], lambda n: n.startswith("Перемычки")),
        "progony": make_set(
            files, ["lm"], lambda n: n.startswith("Прогоны") or n.startswith("Пр-")
        ),
        "fund": make_set(files, ["lm"], lambda n: n.startswith("ФЛ")),
        "materials": {
            it.guid: it
            for it in plity_all
            if not it.name.startswith(("Плиты", "Сваи"))
        },
    }
    lm_used: set[str] = set()
    for key in ("marches", "steps", "pads", "peremy", "progony", "fund"):
        lm_used |= set(c1c[key].keys())
    c1c["lm_other"] = {it.guid: it for it in files["lm"] if it.guid not in lm_used}

    name_idx: dict[str, dict[str, list[Item1C]]] = {key: defaultdict(list) for key in c1c}
    for key, items in c1c.items():
        for item in items.values():
            name_idx[key][norm(item.name)].append(item)
    return c1c, name_idx


def lookup(name_idx, group: str, candidates: Iterable[str]) -> tuple[str, list[Item1C], str]:
    for cand in candidates:
        hits = name_idx[group].get(norm(cand))
        if not hits:
            continue
        guids = {h.guid for h in hits}
        if len(guids) == 1:
            return "ok", hits, ""
        return (
            "ambig",
            hits,
            f"в 1С несколько GUID с таким наименованием ({len(guids)} шт) — ждёт бухгалтерию",
        )
    return "none", [], ""


def pile_base(mark: str) -> str:
    """Марка без ведущей латиницы C: как в исходном матчере _tmp_guid_match."""
    return re.sub(r"^[Cc]\s*", "С ", mark.strip().rstrip("* ").strip())


def suggested_pile_1c_name(mark: str) -> str:
    raw = mark.strip().rstrip("* ").strip()
    body = re.sub(r"^[СCсc]\s*", "", raw)
    return f"Сваи С {body}"


def match_plates(guid: str, c1c) -> MatchResult:
    item = c1c["plates"].get(guid)
    if item:
        return MatchResult("ok", item=item)
    return MatchResult("none", note="GUID отсутствует в новой выгрузке 1С")


def match_piles(mark: str, name_idx) -> MatchResult:
    """Plain GUID обязателен для ✅. Карточка «у» без plain → 🏭 завести только plain."""
    base = pile_base(mark)
    status, hits, note = lookup(name_idx, "piles", [f"Сваи {base}"])
    st_u, hits_u, note_u = lookup(name_idx, "piles", [f"Сваи {base}у"])
    u_item = hits_u[0] if st_u == "ok" else None

    if status == "ok":
        return MatchResult("ok", item=hits[0], u_item=u_item)
    if status == "ambig":
        return MatchResult("ambig", note=note, candidates=hits)
    if st_u == "ok":
        return MatchResult(
            "u_only",
            u_item=u_item,
            note="в 1С есть только усиленная карточка «у» — завести только plain",
        )
    if st_u == "ambig":
        return MatchResult("ambig", note=note_u, candidates=hits_u)
    return MatchResult("none", note="в выгрузке 1С отсутствует")


def match_bridge(mark: str, name_idx) -> MatchResult:
    mm = BR_RX.match(mark.strip())
    cands: list[str] = []
    if mm:
        length, section, kind, num = mm.groups()
        if kind in "TТ":
            cands = [f"Сваи мостовые С {length}.{section}-Т{num}"]
        else:
            cands = [
                f"Сваи мостовые С {length}.{section}-В{num}",
                f"Сваи С {length}0.{section}-ВСв.{num}",
                f"Сваи С {length}0.{section}-НСв.{num}",
            ]
    cands.append(f"Сваи мостовые {mark}")
    status, hits, note = lookup(name_idx, "bridge", cands)
    if status == "ok":
        return MatchResult("ok", item=hits[0])
    if status == "ambig":
        return MatchResult("ambig", note=note, candidates=hits)
    return MatchResult("none", note="такого варианта (Т/В, номер) нет в выгрузке 1С")


def match_fbs(mark: str, name_idx) -> MatchResult:
    status, hits, note = lookup(name_idx, "fbs", [f"Блоки {mark}", mark])
    if status == "ok":
        return MatchResult("ok", item=hits[0])
    if status == "ambig":
        return MatchResult("ambig", note=note, candidates=hits)
    return MatchResult("none", note=note or "в выгрузке 1С отсутствует")


def match_marches(display: str, name_idx) -> MatchResult:
    status, hits, note = lookup(name_idx, "marches", [display])
    if status == "ok":
        return MatchResult("ok", item=hits[0])
    if status == "ambig":
        return MatchResult("ambig", note=note, candidates=hits)
    return MatchResult("none", note=note or "в выгрузке 1С отсутствует")


def match_steps(display: str, mark: str, name_idx) -> MatchResult:
    mm = STEP_RX.match(mark.strip())
    cands = [display]
    analog_names: list[str] = []
    if mm:
        num, rest = mm.groups()
        rest = rest.strip()
        cands.append(f"Лестничные ступени ЛС-{num}{rest}")
        if re.search(r"лев$", rest, re.I):
            base = re.sub(r"лев$", "", rest, flags=re.I).rstrip()
            analog = f"Лестничные ступени ЛС-{num}{base} закл лев"
            analog_names.append(analog)
            cands.append(analog)
            if base in ("", "-1"):
                analog2 = f"Лестничные ступени ЛС-{num} закл лев"
                analog_names.append(analog2)
                cands.append(analog2)
    status, hits, note = lookup(name_idx, "steps", cands)
    if status == "ok":
        hit_name = hits[0].name
        exact = {norm(display)}
        if mm:
            exact.add(norm(cands[1]))
        if norm(hit_name) not in exact:
            return MatchResult(
                "none",
                note=(
                    f"возможный аналог в 1С: «{hit_name}» — не автомат, завести нашу карточку"
                ),
                candidates=hits,
            )
        return MatchResult("ok", item=hits[0])
    if status == "ambig":
        return MatchResult("ambig", note=note, candidates=hits)
    return MatchResult("none", note="в выгрузке 1С отсутствует (серия/вариант не заведены)")


def match_catalog_row(
    group_key: str,
    db_key: str,
    db_label: str,
    name_idx,
    c1c=None,
) -> MatchResult:
    """Сопоставить строку нашего прайса с выгрузкой 1С (тот же матчер, что у очереди)."""
    if group_key == "plates":
        if c1c is None:
            raise TypeError("c1c is required to match plates")
        matched = match_plates(db_key, c1c)
        if matched.status == "ok" and matched.item:
            hits = name_idx["plates"].get(norm(matched.item.name), [])
            if len({h.guid for h in hits}) > 1:
                return MatchResult(
                    "ambig",
                    item=matched.item,
                    note="дубль наименования в 1С — оба GUID уже в prays_plity",
                    candidates=hits,
                )
        return matched
    if group_key == "piles":
        return match_piles(db_key, name_idx)
    if group_key == "bridge":
        return match_bridge(db_key, name_idx)
    if group_key == "fbs":
        return match_fbs(db_key, name_idx)
    if group_key == "marches":
        return match_marches(db_key, name_idx)
    if group_key == "steps":
        return match_steps(db_key, db_label, name_idx)
    raise ValueError(f"Unknown group_key: {group_key!r}")


def iter_our_catalog(
    name_idx,
    db: dict,
    *,
    groups: Iterable[str] = NON_PLATE_QUEUE_GROUPS,
    c1c=None,
):
    """Yield (group_key, db_key, db_label, src, MatchResult) for our price-table rows."""
    for key in groups:
        for db_key, db_label, src in db.get(key) or ():
            yield key, db_key, db_label, src, match_catalog_row(
                key, db_key, db_label, name_idx, c1c
            )


def load_db(db_path: Path) -> dict[str, list[tuple[str, str, str]]]:
    pb = sqlite3.connect(f"file:{db_path}?mode=ro&immutable=1", uri=True)
    try:
        return {
            "plates": [
                (g.lower(), n, "prays_plity")
                for g, n in pb.execute(
                    'SELECT "Уникальный идентификатор (Номенклатура)", "Товар" FROM prays_plity'
                )
            ],
            "piles": [
                (m, m, "pile_prices")
                for (m,) in pb.execute("SELECT DISTINCT mark FROM pile_prices")
            ],
            "bridge": [
                (m, m, "bridge_pile_prices")
                for (m,) in pb.execute("SELECT DISTINCT mark FROM bridge_pile_prices")
            ],
            "fbs": [
                (m, m, "fbs_prices")
                for (m,) in pb.execute("SELECT DISTINCT mark FROM fbs_prices")
            ],
            "marches": [
                (d, m, "march_prices")
                for m, d in pb.execute("SELECT DISTINCT mark, display_name FROM march_prices")
            ],
            "steps": [
                (d, m, "step_prices")
                for m, d in pb.execute("SELECT DISTINCT mark, display_name FROM step_prices")
            ],
        }
    finally:
        pb.close()


def pick_grade_price(grades: dict[str, float]) -> Optional[float]:
    for code in GRADE_PREF:
        if code in grades:
            return grades[code]
    if not grades:
        return None
    return next(iter(grades.values()))


def load_our_unit_prices(db_path: Path) -> dict[str, dict[str, float]]:
    pb = sqlite3.connect(f"file:{db_path}?mode=ro&immutable=1", uri=True)
    out: dict[str, dict[str, float]] = {
        "piles": {},
        "bridge": {},
        "fbs": {},
        "marches": {},
        "steps": {},
    }
    try:
        pile_grades: dict[str, dict[str, float]] = defaultdict(dict)
        for mark, grade, price in pb.execute(
            "SELECT mark, concrete_grade, price FROM pile_prices"
        ):
            pile_grades[mark][grade] = float(price)
        out["piles"] = {m: pick_grade_price(g) or 0.0 for m, g in pile_grades.items()}

        bridge_grades: dict[str, dict[str, float]] = defaultdict(dict)
        for mark, grade, price in pb.execute(
            "SELECT mark, concrete_grade, price FROM bridge_pile_prices"
        ):
            bridge_grades[mark][grade] = float(price)
        out["bridge"] = {m: pick_grade_price(g) or 0.0 for m, g in bridge_grades.items()}

        fbs_grades: dict[str, dict[str, float]] = defaultdict(dict)
        for mark, grade, price in pb.execute(
            "SELECT mark, concrete_grade, price FROM fbs_prices"
        ):
            fbs_grades[mark][grade] = float(price)
        out["fbs"] = {m: pick_grade_price(g) or 0.0 for m, g in fbs_grades.items()}

        march_grades: dict[str, dict[str, float]] = defaultdict(dict)
        for mark, grade, price in pb.execute(
            "SELECT mark, concrete_grade, price FROM march_prices"
        ):
            march_grades[mark][grade] = float(price)
        out["marches"] = {m: pick_grade_price(g) or 0.0 for m, g in march_grades.items()}

        for mark, price in pb.execute("SELECT mark, price FROM step_prices"):
            out["steps"][mark] = float(price)
        return out
    finally:
        pb.close()


def pile_token_from_1c(name: str) -> Optional[str]:
    text = re.sub(r"^сваи\s*", "", str(name or "").strip(), flags=re.I)
    text = re.sub(r"\s*\([^)]*\)\s*", " ", text)
    match = PILE_TOKEN_RE.search(text)
    if match:
        return match.group(1).strip()
    match = re.search(
        r"([СC]\s*\d+(?:[.,]\d+)?)\.(20|30|35|40|45)(?:\.\d+)?",
        text,
    )
    if match:
        return f"{match.group(1)}.{match.group(2)}"
    return None


def is_reinforced_pile_name(name: str) -> bool:
    token = pile_token_from_1c(name)
    if token:
        return token.endswith(("у", "У"))
    n = norm(name)
    if not n.startswith("сваи") or "нсв" in n or "всв" in n:
        return False
    return bool(re.search(r"\dу(?:[а-яa-z]|$)", n))


def pile_geo_from_mark(mark: str) -> tuple[Optional[float], Optional[int]]:
    cleaned = mark.strip()
    cleaned = re.sub(r"[уУ]$", "", cleaned)
    cleaned = re.sub(r"-(\d+)\.\d+$", r"-\1", cleaned)  # -8.1 → тот же ствол, что -8
    cleaned = re.sub(r"-?[АA]\d+.*$", "", cleaned, flags=re.I)
    return parse_pile_mark(cleaned)


def pile_geo_from_1c(name: str) -> tuple[Optional[float], Optional[int]]:
    token = pile_token_from_1c(name)
    if token:
        return pile_geo_from_mark(token)
    match = re.search(
        r"[СC]\s*(\d+(?:[.,]\d+)?)\s*\.\s*(\d+)(?:\s*\.\s*\d+)?",
        str(name or ""),
    )
    if match:
        return parse_pile_mark(f"С{match.group(1)}.{match.group(2)}")
    return None, None


def bridge_geo_from_1c(name: str) -> tuple[Optional[float], Optional[int]]:
    n = norm(name)
    match = BRIDGE_1C_RE.match(n)
    if match:
        length_m = float(match.group(1))
        section_cm = float(match.group(2))
        return length_m, int(round(section_cm * 10))
    our_like = re.sub(r"^сваимостовые", "", n)
    return parse_bridge_pile_geometry(our_like)


def step_key(mark_or_name: str) -> Optional[tuple[str, float]]:
    n = norm(mark_or_name)
    n = re.sub(r"^лестничныеступени", "", n)
    match = STEP_NUM_RE.match(n)
    if not match:
        return None
    rest = match.group(2)
    rest = re.sub(r"в\d.*$", "", rest)
    rest = re.sub(r"закллев$", "лев", rest)
    rest = re.sub(r"закл$", "", rest)
    return rest, float(match.group(1))


def march_key(mark_or_name: str) -> Optional[tuple[str, float]]:
    n = norm(mark_or_name)
    n = re.sub(r"^лестничныемарши", "", n)
    if n.startswith("1лм"):
        prefix, body = "1лм", n[3:]
    elif n.startswith("лм"):
        prefix, body = "лм", n[2:]
    else:
        return None
    num = re.match(r"(\d+(?:\.\d+)?)", body)
    if not num:
        return None
    return prefix, float(num.group(1))


def nearest_neighbor(pool: list[GeoPrice], section, length: float) -> Optional[GeoPrice]:
    same = [item for item in pool if item.section == section]
    if not same:
        return None
    return min(same, key=lambda item: (abs(item.length - length), item.mark))


def build_neighbor_indexes(our_prices: dict[str, dict[str, float]]) -> dict[str, list[GeoPrice]]:
    indexes: dict[str, list[GeoPrice]] = {"piles": [], "bridge": [], "steps": [], "marches": []}
    for mark, price in our_prices["piles"].items():
        length_m, section_mm = pile_geo_from_mark(mark)
        if length_m is None or section_mm is None:
            continue
        indexes["piles"].append(GeoPrice(section_mm, length_m, mark, price))
    for mark, price in our_prices["bridge"].items():
        length_m, section_mm = parse_bridge_pile_geometry(mark)
        if length_m is None or section_mm is None:
            continue
        indexes["bridge"].append(GeoPrice(section_mm, length_m, mark, price))
    for mark, price in our_prices["steps"].items():
        key = step_key(mark)
        if not key:
            continue
        rest, num = key
        indexes["steps"].append(GeoPrice(rest, num, mark, price))
    for mark, price in our_prices["marches"].items():
        key = march_key(mark)
        if not key:
            continue
        prefix, num = key
        indexes["marches"].append(GeoPrice(prefix, num, mark, price))
    return indexes


def neighbor_for(group: str, name_1c: str, indexes: dict[str, list[GeoPrice]]) -> tuple[Optional[float], str]:
    if group == "piles":
        length_m, section_mm = pile_geo_from_1c(name_1c)
        if length_m is None or section_mm is None:
            return None, ""
        hit = nearest_neighbor(indexes["piles"], section_mm, length_m)
    elif group == "bridge":
        length_m, section_mm = bridge_geo_from_1c(name_1c)
        if length_m is None or section_mm is None:
            return None, ""
        hit = nearest_neighbor(indexes["bridge"], section_mm, length_m)
    elif group == "steps":
        key = step_key(name_1c)
        if not key:
            return None, ""
        hit = nearest_neighbor(indexes["steps"], key[0], key[1])
    elif group == "marches":
        key = march_key(name_1c)
        if not key:
            return None, ""
        hit = nearest_neighbor(indexes["marches"], key[0], key[1])
    else:
        return None, ""
    if not hit:
        return None, ""
    return fmt_price(hit.price), hit.mark


def plate_length_dm(name: str) -> Optional[float]:
    match = PLATE_LEN_RE.match(name)
    if not match:
        return None
    return float(match.group(1).replace(",", "."))


def row_from_our(
    group_key: str,
    group_title: str,
    db_key: str,
    db_label: str,
    src: str,
    matched: MatchResult,
) -> QueueRow:
    if matched.status == "ok" and matched.item:
        note = matched.note
        return QueueRow(
            group_key=group_key,
            group_title=group_title,
            state=STATE_READY,
            action=ACTION_READY,
            our_label=db_label,
            our_source=src,
            name_1c=matched.item.name,
            guid=matched.item.guid,
            guid_u=matched.u_item.guid if matched.u_item else "",
            name_u=matched.u_item.name if matched.u_item else "",
            price_1c=matched.item.price,
            unit=matched.item.unit,
            note=note,
        )
    if matched.status == "ambig":
        if matched.item:
            # плиты: каждая строка prays_plity — свой GUID, состояние ⚠️
            return QueueRow(
                group_key=group_key,
                group_title=group_title,
                state=STATE_DUP,
                action=ACTION_DUP,
                our_label=db_label,
                our_source=src,
                name_1c=matched.item.name,
                guid=matched.item.guid,
                price_1c=matched.item.price,
                unit=matched.item.unit,
                note=matched.note or ACTION_DUP,
            )
        guids = " | ".join(sorted({c.guid for c in matched.candidates}))
        names = " | ".join(dict.fromkeys(c.name for c in matched.candidates))
        return QueueRow(
            group_key=group_key,
            group_title=group_title,
            state=STATE_DUP,
            action=ACTION_DUP,
            our_label=db_label,
            our_source=src,
            name_1c=names,
            guid=guids,
            note=matched.note or ACTION_DUP,
        )
    action = ACTION_CREATE
    suggested = db_label
    guid_u = ""
    name_u = ""
    if group_key == "piles":
        suggested = suggested_pile_1c_name(db_label)
        if matched.status == "u_only":
            action = ACTION_CREATE_PLAIN
            if matched.u_item:
                guid_u = matched.u_item.guid
                name_u = matched.u_item.name
    elif group_key == "bridge":
        suggested = f"Сваи мостовые {db_label}"
    elif group_key in ("marches", "steps", "fbs"):
        suggested = db_key if group_key != "fbs" else f"Блоки {db_label}"
    return QueueRow(
        group_key=group_key,
        group_title=group_title,
        state=STATE_CREATE,
        action=action,
        our_label=db_label,
        our_source=src,
        guid_u=guid_u,
        name_u=name_u,
        note=matched.note,
        suggested_1c_name=suggested,
    )


def build_rows(
    c1c,
    name_idx,
    db,
    neighbor_idx: dict[str, list[GeoPrice]],
) -> dict[str, list[QueueRow]]:
    groups = [
        ("plates", "Плиты", True),
        ("piles", "Сваи", True),
        ("bridge", "Сваи мостовые", True),
        ("fbs", "Блоки ФБС", True),
        ("marches", "Марши", True),
        ("steps", "Ступени", True),
        ("piles_comp", "Сваи составные", False),
        ("pads", "Площадки", False),
        ("peremy", "Перемычки", False),
        ("progony", "Прогоны", False),
        ("fund", "Фундаменты", False),
        ("lm_other", "Прочее ЛМ", False),
        ("materials", "Материалы 1С", False),
    ]
    plates_db_guids = {g for g, _n, _s in db["plates"]}
    by_group: dict[str, list[QueueRow]] = {key: [] for key, _t, _h in groups}
    claimed: dict[str, set[str]] = {key: set() for key in c1c}

    for key, title, has_db in groups:
        if not has_db:
            continue
        for db_key, db_label, src in db[key]:
            matched = match_catalog_row(key, db_key, db_label, name_idx, c1c)
            row = row_from_our(key, title, db_key, db_label, src, matched)
            by_group[key].append(row)
            if matched.item:
                claimed[key].add(matched.item.guid)
            if matched.u_item:
                claimed[key].add(matched.u_item.guid)
            if matched.status == "ambig":
                for cand in matched.candidates:
                    claimed[key].add(cand.guid)
            # аналог ступеней (status none + candidates): карточку 1С не claim'им —
            # она пойдёт в 💰 как отдельное изделие.

    # 1C-only
    for key, title, has_db in groups:
        items = c1c[key]
        extra_exclude = plates_db_guids if key == "materials" else set()
        for guid, item in items.items():
            if guid in claimed[key] or guid in extra_exclude:
                continue
            if key == "plates" and not item.name.startswith("Плиты"):
                continue
            if key in ("piles", "bridge", "fbs", "marches", "steps"):
                if key == "piles" and is_reinforced_pile_name(item.name):
                    # «у» без нашей марки — цена общая с plain, отдельной строки 💰 нет
                    continue
                neigh_price, neigh_label = neighbor_for(key, item.name, neighbor_idx)
                by_group[key].append(
                    QueueRow(
                        group_key=key,
                        group_title=title,
                        state=STATE_PRICE,
                        action=ACTION_PRICE,
                        name_1c=item.name,
                        guid=item.guid,
                        price_1c=item.price,
                        unit=item.unit,
                        neighbor_price=neigh_price,
                        neighbor_label=neigh_label,
                        note="цена вводится за штуку; для свай — одна на марку (все классы)",
                    )
                )
            elif key == "plates":
                length_dm = plate_length_dm(item.name)
                short = length_dm in SHORT_PLATE_LENGTHS if length_dm is not None else False
                by_group[key].append(
                    QueueRow(
                        group_key=key,
                        group_title=title,
                        state=STATE_PLATE_ONLY,
                        action=ACTION_PLATE_SHORT if short else ACTION_PLATE_MATRIX,
                        name_1c=item.name,
                        guid=item.guid,
                        price_1c=item.price,
                        unit=item.unit,
                        note=(
                            f"длина {length_dm} дм — вне матрицы, открытый вопрос"
                            if short
                            else "SKU 1С покрывается матрицей длина×нагрузка"
                        ),
                    )
                )
            elif key == "materials":
                dup_hits = name_idx[key].get(norm(item.name), [])
                note = ""
                if len({h.guid for h in dup_hits}) > 1:
                    note = "дубль наименования в 1С — не изделие прайса"
                by_group[key].append(
                    QueueRow(
                        group_key=key,
                        group_title=title,
                        state=STATE_MAT,
                        action=ACTION_NONE,
                        name_1c=item.name,
                        guid=item.guid,
                        price_1c=item.price,
                        unit=item.unit,
                        note=note,
                    )
                )
            else:
                dup_hits = name_idx[key].get(norm(item.name), [])
                note = ""
                if len({h.guid for h in dup_hits}) > 1:
                    note = "дубль наименования в 1С — группа не ведётся"
                by_group[key].append(
                    QueueRow(
                        group_key=key,
                        group_title=title,
                        state=STATE_OUT,
                        action=ACTION_NONE,
                        name_1c=item.name,
                        guid=item.guid,
                        price_1c=item.price,
                        unit=item.unit,
                        note=note,
                    )
                )
    return by_group


def sort_rows(rows: list[QueueRow]) -> list[QueueRow]:
    return sorted(
        rows,
        key=lambda r: (
            STATE_SORT.get(r.state, 9),
            r.our_label or "",
            r.name_1c or "",
        ),
    )


def style_header(ws: Worksheet, headers: list[str], widths: list[int]) -> None:
    ws.append(headers)
    for col, width in enumerate(widths, start=1):
        cell = ws.cell(row=1, column=col)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"


def paint_state(ws: Worksheet, row_idx: int, state: str) -> None:
    fill = STATE_FILLS.get(state)
    if fill:
        ws.cell(row=row_idx, column=1).fill = fill


def add_group_sheet(wb: Workbook, title: str, rows: list[QueueRow], *, piles: bool) -> None:
    headers = [
        "Состояние",
        "Действие",
        "Марка/наименование у нас",
        "Источник у нас",
        "Наименование 1С",
        "GUID 1С",
    ]
    widths = [28, 42, 34, 20, 46, 38]
    if piles:
        headers += ["GUID «у»-варианта", "Наименование «у»-варианта"]
        widths += [38, 34]
    headers += [
        "Цена 1С (Продажная)",
        "Ед.",
        "Цена за шт (ввод)",
        "как у ближайшего соседа",
        "Сосед (марка)",
        "Завести в 1С как",
        "Примечание",
    ]
    widths += [18, 8, 18, 22, 22, 36, 56]
    ws = wb.create_sheet(title[:31])
    style_header(ws, headers, widths)
    for row in sort_rows(rows):
        values = [
            row.state,
            row.action,
            row.our_label or None,
            row.our_source or None,
            row.name_1c or None,
            row.guid or None,
        ]
        if piles:
            values += [row.guid_u or None, row.name_u or None]
        values += [
            fmt_price(row.price_1c),
            row.unit or None,
            None,  # ручной ввод
            row.neighbor_price,
            row.neighbor_label or None,
            row.suggested_1c_name or None,
            row.note or None,
        ]
        ws.append(values)
        paint_state(ws, ws.max_row, row.state)
        price_col = 11 if not piles else 13
        ws.cell(row=ws.max_row, column=price_col).number_format = "#,##0.00"
    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col}{max(len(rows) + 1, 1)}"


def add_simple_sheet(wb: Workbook, title: str, rows: list[QueueRow]) -> None:
    headers = [
        "Состояние",
        "Действие",
        "GUID 1С",
        "Наименование 1С",
        "Цена 1С (Продажная)",
        "Ед.",
        "Примечание",
    ]
    widths = [22, 18, 38, 52, 18, 8, 50]
    ws = wb.create_sheet(title[:31])
    style_header(ws, headers, widths)
    for row in sort_rows(rows):
        ws.append(
            [
                row.state,
                row.action,
                row.guid or None,
                row.name_1c or None,
                fmt_price(row.price_1c),
                row.unit or None,
                row.note or None,
            ]
        )
        paint_state(ws, ws.max_row, row.state)
    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col}{max(len(rows) + 1, 1)}"


def add_tasks_sheet(wb: Workbook, by_group: dict[str, list[QueueRow]]) -> tuple[int, int]:
    create_rows = [
        row
        for key in ("piles", "bridge", "steps", "fbs", "marches")
        for row in by_group[key]
        if row.state == STATE_CREATE
    ]
    price_rows = [
        row
        for key in ("piles", "bridge", "fbs", "marches", "steps")
        for row in by_group[key]
        if row.state == STATE_PRICE
    ]
    create_rows = sort_rows(create_rows)
    price_rows = sort_rows(price_rows)

    ws = wb.create_sheet("Задачи")
    ws["A1"] = "🏭 завести в 1С"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].fill = SECTION_CREATE_FILL
    create_headers = [
        "Группа",
        "Марка у нас",
        "Завести в 1С как",
        "Действие",
        "GUID «у» уже есть",
        "Наименование «у»",
        "Примечание",
    ]
    create_widths = [18, 28, 36, 24, 38, 36, 60]
    for col, (header, width) in enumerate(zip(create_headers, create_widths), start=1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        ws.column_dimensions[get_column_letter(col)].width = width
    for row in create_rows:
        ws.append(
            [
                row.group_title,
                row.our_label,
                row.suggested_1c_name or row.our_label,
                row.action,
                row.guid_u or None,
                row.name_u or None,
                row.note or None,
            ]
        )
        ws.cell(row=ws.max_row, column=4).fill = STATE_FILLS[STATE_CREATE]

    gap = ws.max_row + 2
    ws.cell(row=gap, column=1, value="💰 ввести цену")
    ws.cell(row=gap, column=1).font = Font(bold=True, size=14)
    ws.cell(row=gap, column=1).fill = SECTION_PRICE_FILL
    price_headers = [
        "Группа",
        "Наименование 1С",
        "GUID 1С",
        "Цена за шт (ввод)",
        "как у ближайшего соседа",
        "Сосед (марка)",
        "Цена 1С (Продажная)",
        "Ед.",
        "Примечание",
    ]
    price_widths = [18, 52, 38, 18, 22, 22, 18, 8, 56]
    hdr_row = gap + 1
    for col, (header, width) in enumerate(zip(price_headers, price_widths), start=1):
        cell = ws.cell(row=hdr_row, column=col, value=header)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        ws.column_dimensions[get_column_letter(col)].width = max(
            ws.column_dimensions[get_column_letter(col)].width or 0, width
        )
    for row in price_rows:
        ws.append(
            [
                row.group_title,
                row.name_1c,
                row.guid,
                None,
                row.neighbor_price,
                row.neighbor_label or None,
                fmt_price(row.price_1c),
                row.unit or None,
                row.note or None,
            ]
        )
        ws.cell(row=ws.max_row, column=4).number_format = "#,##0.00"
        ws.cell(row=ws.max_row, column=5).number_format = "#,##0.00"
        ws.cell(row=ws.max_row, column=1).fill = STATE_FILLS[STATE_PRICE]
    ws.freeze_panes = "A3"
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="left")
    return len(create_rows), len(price_rows)


_KIND_TITLES = {
    "pile": "Сваи",
    "bridge_pile": "Сваи мостовые",
    "fbs": "Блоки ФБС",
    "stair_flight": "Марши",
    "stair_step": "Ступени",
    "plate": "Плиты",
}


def add_live_tasks_sheet(wb: Workbook, conn: sqlite3.Connection) -> tuple[int, int, int]:
    """Лист «Задачи»: живые очереди 🏭/💰/⚠️ из collect_guid_queue (тот же состав, что API)."""
    from core.guid_queue import collect_guid_queue

    snap = collect_guid_queue(conn)
    ws = wb.create_sheet("Задачи")
    ws["A1"] = STATE_CREATE
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].fill = SECTION_CREATE_FILL
    create_headers = ["Группа", "Марка у нас", "Поле", "Действие"]
    for col, (header, width) in enumerate(zip(create_headers, [18, 28, 16, 60]), start=1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        ws.column_dimensions[get_column_letter(col)].width = width
    for item in snap.to_create_1c:
        ws.append(
            [
                _KIND_TITLES.get(item.product_kind, item.product_kind),
                item.mark,
                item.field,
                item.hint,
            ]
        )
        ws.cell(row=ws.max_row, column=4).fill = STATE_FILLS[STATE_CREATE]

    gap = ws.max_row + 2
    ws.cell(row=gap, column=1, value=STATE_PRICE)
    ws.cell(row=gap, column=1).font = Font(bold=True, size=14)
    ws.cell(row=gap, column=1).fill = SECTION_PRICE_FILL
    price_headers = ["Группа", "Марка", "GUID 1С", "Наименование 1С"]
    hdr_row = gap + 1
    for col, (header, width) in enumerate(zip(price_headers, [18, 28, 38, 52]), start=1):
        cell = ws.cell(row=hdr_row, column=col, value=header)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        ws.column_dimensions[get_column_letter(col)].width = max(
            ws.column_dimensions[get_column_letter(col)].width or 0, width
        )
    for item in snap.to_price:
        ws.append(
            [
                _KIND_TITLES.get(item.product_kind, item.product_kind),
                item.mark,
                item.guid,
                item.name,
            ]
        )
        ws.cell(row=ws.max_row, column=1).fill = STATE_FILLS[STATE_PRICE]

    gap = ws.max_row + 2
    ws.cell(row=gap, column=1, value=STATE_DUP)
    ws.cell(row=gap, column=1).font = Font(bold=True, size=14)
    ws.cell(row=gap, column=1).fill = STATE_FILLS[STATE_DUP]
    dup_headers = ["Scope", "Ключ", "Группа", "Кандидаты GUID"]
    hdr_row = gap + 1
    for col, (header, width) in enumerate(zip(dup_headers, [14, 28, 18, 60]), start=1):
        cell = ws.cell(row=hdr_row, column=col, value=header)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        ws.column_dimensions[get_column_letter(col)].width = max(
            ws.column_dimensions[get_column_letter(col)].width or 0, width
        )
    for item in snap.duplicates:
        packed = ", ".join(guid for guid, _name, _price in item.candidates)
        ws.append(
            [
                item.scope,
                item.key,
                _KIND_TITLES.get(item.product_kind, item.product_kind),
                packed,
            ]
        )
        ws.cell(row=ws.max_row, column=1).fill = STATE_FILLS[STATE_DUP]

    ws.freeze_panes = "A3"
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="left")
    return len(snap.to_create_1c), len(snap.to_price), len(snap.duplicates)


def count_states(rows: Iterable[QueueRow]) -> dict[str, int]:
    out: dict[str, int] = defaultdict(int)
    for row in rows:
        out[row.state] += 1
    return out


def add_summary_sheet(
    wb: Workbook,
    by_group: dict[str, list[QueueRow]],
    create_n: int,
    price_n: int,
    dup_names_1c: int,
    notes: list[str],
) -> None:
    ws = wb.create_sheet("Сводка", 0)
    headers = [
        "Группа",
        "✅ готово",
        "⚠️ дубль",
        "🏭 завести в 1С",
        "💰 ввести цену",
        "⏸️ вне контура",
        "🚫 материалы",
        "только в 1С (плиты)",
        "Всего строк",
    ]
    widths = [20, 14, 12, 18, 16, 16, 14, 20, 14]
    style_header(ws, headers, widths)
    order = [
        "plates",
        "piles",
        "bridge",
        "fbs",
        "marches",
        "steps",
        "piles_comp",
        "peremy",
        "progony",
        "fund",
        "pads",
        "lm_other",
        "materials",
    ]
    titles = {
        "plates": "Плиты",
        "piles": "Сваи",
        "bridge": "Сваи мостовые",
        "fbs": "Блоки ФБС",
        "marches": "Марши",
        "steps": "Ступени",
        "piles_comp": "Сваи составные",
        "peremy": "Перемычки",
        "progony": "Прогоны",
        "fund": "Фундаменты",
        "pads": "Площадки",
        "lm_other": "Прочее ЛМ",
        "materials": "Материалы 1С",
    }
    totals = defaultdict(int)
    for key in order:
        rows = by_group[key]
        counts = count_states(rows)
        line = [
            titles[key],
            counts.get(STATE_READY, 0),
            counts.get(STATE_DUP, 0),
            counts.get(STATE_CREATE, 0),
            counts.get(STATE_PRICE, 0),
            counts.get(STATE_OUT, 0),
            counts.get(STATE_MAT, 0),
            counts.get(STATE_PLATE_ONLY, 0),
            len(rows),
        ]
        ws.append(line)
        for state, col in (
            (STATE_READY, 2),
            (STATE_DUP, 3),
            (STATE_CREATE, 4),
            (STATE_PRICE, 5),
            (STATE_OUT, 6),
            (STATE_MAT, 7),
        ):
            if line[col - 1]:
                ws.cell(row=ws.max_row, column=col).fill = STATE_FILLS[state]
        for i, state in enumerate(
            (STATE_READY, STATE_DUP, STATE_CREATE, STATE_PRICE, STATE_OUT, STATE_MAT, STATE_PLATE_ONLY),
            start=2,
        ):
            totals[state] += counts.get(state, 0)
        totals["rows"] += len(rows)
    ws.append(
        [
            "ИТОГО",
            totals[STATE_READY],
            totals[STATE_DUP],
            totals[STATE_CREATE],
            totals[STATE_PRICE],
            totals[STATE_OUT],
            totals[STATE_MAT],
            totals[STATE_PLATE_ONLY],
            totals["rows"],
        ]
    )
    for col in range(1, 10):
        ws.cell(row=ws.max_row, column=col).font = Font(bold=True)

    ws.append([])
    ws.append(["Лист «Задачи»", f"🏭 {create_n} строк", f"💰 {price_n} строк"])
    ws.append(["Дубли имён в выгрузке 1С (все группы)", dup_names_1c, "см. «Дубли GUID в 1С.xlsx»"])
    ws.append([])
    ws.append(["Примечания:"])
    for text in notes:
        ws.append([text])


def global_dup_name_count(name_idx) -> int:
    seen: dict[str, set[str]] = defaultdict(set)
    for group_idx in name_idx.values():
        for nkey, hits in group_idx.items():
            for hit in hits:
                seen[nkey].add(hit.guid)
    return sum(1 for guids in seen.values() if len(guids) > 1)


def build_notes(by_group: dict[str, list[QueueRow]], dup_names_1c: int) -> list[str]:
    piles_create = [r for r in by_group["piles"] if r.state == STATE_CREATE]
    u_only = [r for r in piles_create if r.action == ACTION_CREATE_PLAIN]
    neither = [r for r in piles_create if r.action == ACTION_CREATE]
    ready_non_plate = sum(
        1
        for key in ("piles", "bridge", "fbs", "marches", "steps")
        for r in by_group[key]
        if r.state == STATE_READY
    )
    return [
        "Плиты не мигрируют: GUID уже в prays_plity. «Только в 1С» (~229) не в «Задачах».",
        "Сваи «у»: отдельная карточка 1С, цена общая. ✅ если есть plain GUID; нет «у» ≠ 🏭.",
        f"Сваи 🏭: {len(u_only)} марок «завести только plain» (в 1С уже есть «у») "
        f"и {len(neither)} марок «завести» (нет ни plain, ни «у»).",
        "Дубли не разрешаются скриптом. ⚠️ видно на листах изделий, в «Задачи» не попадают.",
        f"Имён-дублей во всех выгрузках 1С: {dup_names_1c} (как в «Дубли GUID в 1С.xlsx»).",
        f"Не-плиты ✅ готово: {ready_non_plate} (спека 14.09 считала 271, включая 22 «у»-only как GUID).",
        "💰 подсказка «как у ближайшего соседа»: то же сечение, ближайшая длина; цена B25 (или первый класс).",
        "Авто-ступени цены между классами бетона не считаем (шаг 79–599 руб, правила нет).",
        "Источник 1С: выгрузки «Прайс-лист» из папки GUID. Прайс программы: pb.db, не pile_catalog.",
        "Лист «Плиты» без строк ✅ — их десятки тысяч в prays_plity; на листе дубли и «только в 1С».",
    ]


def print_report(
    by_group: dict[str, list[QueueRow]],
    create_n: int,
    price_n: int,
    dup_names_1c: int,
    out_path: Path,
) -> None:
    titles = {
        "plates": "Плиты",
        "piles": "Сваи",
        "bridge": "Сваи мостовые",
        "fbs": "Блоки ФБС",
        "marches": "Марши",
        "steps": "Ступени",
        "piles_comp": "Сваи составные",
        "peremy": "Перемычки",
        "progony": "Прогоны",
        "fund": "Фундаменты",
        "pads": "Площадки",
        "lm_other": "Прочее ЛМ",
        "materials": "Материалы 1С",
    }
    print(f"Сохранено: {out_path}")
    print()
    header = (
        f"{'Группа':<18}{'✅':>8}{'⚠️':>8}{'🏭':>8}{'💰':>8}"
        f"{'⏸️':>8}{'🚫':>8}{'плит1С':>8}"
    )
    print(header)
    totals = defaultdict(int)
    for key, title in titles.items():
        counts = count_states(by_group[key])
        print(
            f"{title:<18}"
            f"{counts.get(STATE_READY, 0):>8}"
            f"{counts.get(STATE_DUP, 0):>8}"
            f"{counts.get(STATE_CREATE, 0):>8}"
            f"{counts.get(STATE_PRICE, 0):>8}"
            f"{counts.get(STATE_OUT, 0):>8}"
            f"{counts.get(STATE_MAT, 0):>8}"
            f"{counts.get(STATE_PLATE_ONLY, 0):>8}"
        )
        for state, n in counts.items():
            totals[state] += n
    print(
        f"{'ИТОГО':<18}"
        f"{totals[STATE_READY]:>8}"
        f"{totals[STATE_DUP]:>8}"
        f"{totals[STATE_CREATE]:>8}"
        f"{totals[STATE_PRICE]:>8}"
        f"{totals[STATE_OUT]:>8}"
        f"{totals[STATE_MAT]:>8}"
        f"{totals[STATE_PLATE_ONLY]:>8}"
    )
    ready_non_plate = sum(
        1
        for key in ("piles", "bridge", "fbs", "marches", "steps")
        for r in by_group[key]
        if r.state == STATE_READY
    )
    u_only = sum(1 for r in by_group["piles"] if r.action == ACTION_CREATE_PLAIN)
    create_plain = sum(
        1 for r in by_group["piles"] if r.state == STATE_CREATE and r.action == ACTION_CREATE
    )
    print()
    print(f"✅ не-плит: {ready_non_plate}")
    print(f"⚠️ имён-дублей во всех выгрузках 1С: {dup_names_1c}")
    print(f"🏭 сваи: {count_states(by_group['piles']).get(STATE_CREATE, 0)} "
          f"(завести только plain {u_only}, завести {create_plain})")
    print(f"Лист «Задачи»: 🏭 {create_n}   💰 {price_n}")
    print()
    print("Спека 14.09.2026: ✅ 57251 (не-плит 271) · ⚠️ 58 · 🏭 224 · 💰 ~98 · ⏸️ 232")
    print("Расхождения:")
    print(
        f"  ✅ {totals[STATE_READY]} (не-плит {ready_non_plate}) vs 57251 (271): "
        "спека считала все строки prays_plity как ✅ и 22 сваи с только «у» как с GUID."
    )
    print(
        f"  ⚠️ строк на листах {totals[STATE_DUP]} / имён в 1С {dup_names_1c} vs 58 имён: "
        "58 — уникальные дубли имён во всех файлах 1С; на листах прайса — по строкам."
    )
    print(
        f"  🏭 {totals[STATE_CREATE]} vs 224 (сваи 123 / мостовые 71 / ступени 30): "
        "123 в спеке = старый лист «без GUID» (121 нет + 2 дубля). "
        "Дубли ушли в ⚠️; 22 «у»-only добавлены в 🏭 (решение 15.09). "
        "Мостовые/ступени: −1 дубль из 🏭 в ⚠️; ступени-аналоги остаются 🏭."
    )
    print(
        f"  💰 {totals[STATE_PRICE]} vs ~98: из «только в 1С» исключены «у»-карточки свай "
        "и дубли, которые теперь ⚠️, а не очередь цены."
    )
    out_pause = totals[STATE_OUT]
    named_pause = (
        count_states(by_group["piles_comp"]).get(STATE_OUT, 0)
        + count_states(by_group["peremy"]).get(STATE_OUT, 0)
        + count_states(by_group["progony"]).get(STATE_OUT, 0)
        + count_states(by_group["fund"]).get(STATE_OUT, 0)
        + count_states(by_group["pads"]).get(STATE_OUT, 0)
    )
    print(
        f"  ⏸️ {out_pause} (из них 5 групп спеки {named_pause}, ожидалось 232; "
        f"плюс «Прочее ЛМ» {count_states(by_group['lm_other']).get(STATE_OUT, 0)})."
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Собрать Excel-очередь синхронизации GUID и цен программа ↔ 1С."
    )
    parser.add_argument(
        "--guid-dir",
        default=str(DEFAULT_GUID_DIR),
        help="Папка выгрузок 1С (по умолчанию: банк знаний/Новая папка/GUID)",
    )
    parser.add_argument(
        "--db",
        default=str(PRICE_DB_PATH),
        help=f"Путь к pb.db (по умолчанию: {PRICE_DB_PATH})",
    )
    parser.add_argument(
        "--out",
        default=str(DEFAULT_OUT),
        help="Куда писать xlsx (исходное «Сопоставление…» не трогаем)",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    guid_dir = Path(args.guid_dir)
    db_path = Path(args.db)
    out_path = Path(args.out)
    if not guid_dir.is_dir():
        print(f"❌ Нет папки выгрузок 1С: {guid_dir}")
        return 1
    if not db_path.is_file():
        print(f"❌ Нет pb.db: {db_path}")
        return 1
    for fname in PRICE_FILE_NAMES.values():
        path = guid_dir / fname
        if not path.is_file():
            print(f"❌ Нет файла выгрузки: {path}")
            return 1

    print("Загрузка выгрузок 1С и pb.db…")
    c1c, name_idx = load_1c(guid_dir)
    db = load_db(db_path)
    our_prices = load_our_unit_prices(db_path)
    neighbor_idx = build_neighbor_indexes(our_prices)
    print("Сверка…")
    by_group = build_rows(c1c, name_idx, db, neighbor_idx)
    dup_names_1c = global_dup_name_count(name_idx)

    print("Запись Excel…")
    wb = Workbook()
    wb.remove(wb.active)
    live_conn = sqlite3.connect(str(db_path))
    try:
        create_n, price_n, _dup_n = add_live_tasks_sheet(wb, live_conn)
        live_conn.commit()
    finally:
        live_conn.close()
    notes = build_notes(by_group, dup_names_1c)
    add_summary_sheet(wb, by_group, create_n, price_n, dup_names_1c, notes)
    # Лист «Плиты»: только ⚠️ и «только в 1С». ✅ (~57k в prays_plity) — в сводке.
    plate_queue = [row for row in by_group["plates"] if row.state != STATE_READY]
    add_group_sheet(wb, "Плиты", plate_queue, piles=False)
    add_group_sheet(wb, "Сваи", by_group["piles"], piles=True)
    add_group_sheet(wb, "Сваи мостовые", by_group["bridge"], piles=False)
    add_group_sheet(wb, "Блоки ФБС", by_group["fbs"], piles=False)
    add_group_sheet(wb, "Марши", by_group["marches"], piles=False)
    add_group_sheet(wb, "Ступени", by_group["steps"], piles=False)
    add_simple_sheet(wb, "Сваи составные", by_group["piles_comp"])
    add_simple_sheet(wb, "Перемычки", by_group["peremy"])
    add_simple_sheet(wb, "Прогоны", by_group["progony"])
    add_simple_sheet(wb, "Фундаменты", by_group["fund"])
    add_simple_sheet(wb, "Площадки", by_group["pads"])
    add_simple_sheet(wb, "Прочее ЛМ", by_group["lm_other"])
    add_simple_sheet(wb, "Материалы 1С", by_group["materials"])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    print_report(by_group, create_n, price_n, dup_names_1c, out_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
