# -*- coding: utf-8 -*-
"""Сетка пула 1/5/10/20/40/60 на четырёх окнах КЗ 2026.

Режим укладки — «сильные первыми». Для каждой дорожки карты и для каждой
дорожки перераскладки: переармирование, перебетонирование, остаток ширины
(≥200 мм) и обрезок (<200 мм), плюс незалитый хвост дорожки.
"""
from __future__ import annotations

import contextlib
import io
import json
import logging
import sys
import time
import traceback
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.build_kz_overrun_xlsx_2026 import (  # noqa: E402
    CACHE,
    DELTA_B40_B30,
    GRADE_RANK,
    VOL_M_PER_M_12,
    WIRE_PRICE_M,
    grade_of,
    load_wires,
    make_lookup,
)

OUT = Path("/home/roman/Рабочий стол/с производства/Пул_перерасход_КЗ_2026.xlsx")
JSON_OUT = ROOT / "data" / "pool_experiment_2026.json"

TRACK_M = 1.2
MIN_USEFUL_MM = 200
WINDOWS = (
    ("зима", 0),
    ("весна", 116),
    ("лето", 219),
    ("осень", 348),
)
SIZES = (1, 5, 10, 20, 40, 60)


def series_load(load: float) -> int:
    lc = int(round(float(load)))
    return 13 if lc == 12 else lc


def pieces_of(track: dict, lookup) -> list[dict]:
    """Плиты дорожки, для которых есть норма проволок."""
    out = []
    for l_dm, width_dm, load, qty in track["plates"]:
        lc = series_load(load)
        ld = int(round(l_dm))
        wires = lookup(ld, lc)
        if wires is None:
            continue
        width_mm = int(round(float(width_dm) * 100))
        length_m = float(l_dm) / 10.0
        width_label = int(width_mm / 100) if abs(width_mm / 100 - round(width_mm / 100)) < 0.05 else round(width_mm / 100, 1)
        for _ in range(int(qty)):
            out.append({
                "length": length_m,
                "width_mm": width_mm,
                "load": lc,
                "wires": float(wires),
                "grade": grade_of(ld, lc),
                "ld": ld,
                "name": f"ПБ {ld}-{width_label}-{lc}п",
            })
    return out


def pack_slabs(pieces: list[dict]) -> list[dict]:
    """Плиты одной длины стоят рядом в плите 1200 мм. Остаток — незаполненная ширина."""
    slabs: list[dict] = []
    for p in sorted(pieces, key=lambda x: -x["length"]):
        placed = False
        for slab in slabs:
            if abs(slab["length"] - p["length"]) <= 0.03 and slab["used_mm"] + p["width_mm"] <= 1200:
                slab["pieces"].append(p)
                slab["used_mm"] += p["width_mm"]
                placed = True
                break
        if not placed:
            slabs.append({
                "length": p["length"],
                "used_mm": p["width_mm"],
                "pieces": [p],
            })
    return slabs


def composition_of(pieces: list[dict]) -> str:
    counts: dict[str, int] = defaultdict(int)
    order: list[str] = []
    for p in pieces:
        name = p["name"]
        if name not in counts:
            order.append(name)
        counts[name] += 1
    return "; ".join(f"{name}×{counts[name]}" for name in order)


def money_and_waste(pieces: list[dict], track_len: float | None) -> dict:
    empty = {
        "qty": 0, "occupied_m": 0.0, "plate_m": 0.0, "wire_rub": 0.0, "concrete_rub": 0.0,
        "over_bc_m": 0.0, "bc_pct": 0.0, "delta_avg": 0.0, "delta_max": 0.0, "over_wire_m": 0.0,
        "strip_m2": 0.0, "scrap_m2": 0.0, "end_gap_m2": 0.0,
        "max_wires": 0.0, "grades": "", "composition": "",
    }
    if not pieces:
        return empty
    slabs = pack_slabs(pieces)
    occupied = sum(s["length"] for s in slabs)
    plate_m = sum(p["length"] for p in pieces)
    tmax = max(p["wires"] for p in pieces)
    gmax = max(GRADE_RANK[p["grade"]] for p in pieces)
    wire_rub = concrete_rub = strip_m2 = scrap_m2 = 0.0
    over_bc_m = over_wire_m = delta_sum = 0.0
    over_qty = 0
    for p in pieces:
        if GRADE_RANK[p["grade"]] < gmax:
            over_bc_m += p["length"]
            vol = VOL_M_PER_M_12 * (p["width_mm"] / 1200.0) * p["length"]
            concrete_rub += vol * DELTA_B40_B30
        d = tmax - p["wires"]
        if d > 0.01:
            over_wire_m += d * p["length"]
            wire_rub += d * p["length"] * WIRE_PRICE_M
            delta_sum += d
            over_qty += 1
    for slab in slabs:
        gap_mm = max(0, 1200 - slab["used_mm"])
        area = (gap_mm / 1000.0) * slab["length"]
        if gap_mm >= MIN_USEFUL_MM:
            strip_m2 += area
        elif gap_mm > 5:
            scrap_m2 += area
    poured = track_len if track_len and track_len > 0 else occupied
    end_gap_m2 = max(0.0, poured - occupied) * TRACK_M
    return {
        "qty": len(pieces),
        "occupied_m": round(occupied, 2),
        "plate_m": plate_m,
        "wire_rub": wire_rub,
        "concrete_rub": concrete_rub,
        "over_bc_m": over_bc_m,
        "bc_pct": over_bc_m / plate_m if plate_m else 0.0,
        "delta_avg": delta_sum / over_qty if over_qty else 0.0,
        "delta_max": max(tmax - p["wires"] for p in pieces),
        "over_wire_m": over_wire_m,
        "strip_m2": strip_m2,
        "scrap_m2": scrap_m2,
        "end_gap_m2": end_gap_m2,
        "max_wires": tmax,
        "grades": ",".join(sorted({p["grade"] for p in pieces})),
        "composition": composition_of(pieces),
    }


def per_m(value: float, meters: float) -> float:
    return value / meters if meters else 0.0


def signature(pieces: list[dict]) -> Counter:
    return Counter((round(p["length"], 2), p["width_mm"], p["load"]) for p in pieces)


def layout_pieces(track: dict, lookup) -> list[dict]:
    """Плиты дорожки после раскладки — те же поля, что у карт."""
    out = []
    extra_scrap = 0.0
    for it in track.get("items") or []:
        length = float(it.get("length") or 0)
        width_m = float(it.get("width") or TRACK_M)
        width_mm = int(round(width_m * 1000))
        load = series_load(it.get("load_code") or 8)
        ld = int(round(length * 10))
        wires = it.get("reinforcement")
        if wires is None:
            wires = lookup(ld, load)
        if wires is None or length <= 0:
            continue
        mode = it.get("mode")
        if mode == "transverse":
            rem = float(it.get("remainder") or 0)
            if rem > 0.05:
                extra_scrap += rem * width_m
            length = float(it.get("target_length") or length)
            ld = int(round(length * 10))
        width_label = int(width_mm / 100) if abs(width_mm / 100 - round(width_mm / 100)) < 0.05 else round(width_mm / 100, 1)
        out.append({
            "length": length,
            "width_mm": width_mm,
            "load": load,
            "wires": float(wires),
            "grade": grade_of(ld, load),
            "ld": ld,
            "name": it.get("plate_name") or f"ПБ {ld}-{width_label}-{load}п",
        })
    return out, extra_scrap


def relayout(pieces: list[dict]) -> tuple[list[dict], str]:
    """Перераскладка пула. Возвращает дорожки программы и текст ошибки."""
    from core.domain.plate_order import PlateOrder
    from core.optimization import optimize_with_cascading_longitudinal_cuts
    from core.optimization.layout_runtime_snapshot import (
        build_layout_runtime_snapshot_from_plate_order_context,
    )
    from core.optimization.result_contract import is_optimization_success
    from core.plate_order_context import PlateOrderContext
    from core.visualization import LayoutIntegrityError, TrackLayoutInvariantError, split_sequence_into_tracks
    from viz_modules.layout_sequence.builder import build_layout_sequence

    orders_2d = []
    grouped: dict[tuple, int] = defaultdict(int)
    for p in pieces:
        grouped[(round(p["length"], 3), p["width_mm"], p["load"])] += 1
    for (length, width_mm, load), qty in grouped.items():
        orders_2d.append({
            "length": length,
            "width": width_mm,
            "qty": qty,
            "load_code": load,
            "plate_name": f"ПБ {int(round(length * 10))}-{width_mm / 100:.2g}-{load}п",
        })

    logging.disable(logging.CRITICAL)
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            core_order = PlateOrder.from_orders_2d(orders_2d)
            plate_ctx = PlateOrderContext.fresh_empty()
            plate_ctx.hydrate_from_order(core_order)
            with plate_ctx.bound():
                optimization_result = optimize_with_cascading_longitudinal_cuts(orders_2d=orders_2d)
            if not is_optimization_success(optimization_result) or not optimization_result.get("total_plates"):
                return [], f"optimize:{optimization_result.get('_opt_error_code') or optimization_result.get('_opt_status')}"
            all_loads = sorted({int(float(item.get("load_code", 8))) for item in orders_2d})
            plate_ctx.load_optimization_snapshot(
                optimization_result=optimization_result,
                plan_by_load={"all": optimization_result},
                load_to_reinforcement_map={load: ["all"] for load in all_loads},
            )
            with plate_ctx.bound():
                layout_rt = build_layout_runtime_snapshot_from_plate_order_context(
                    plate_ctx,
                    layout_reinforcement_order="desc",
                )
                seq = build_layout_sequence(runtime=layout_rt)
                try:
                    tracks = split_sequence_into_tracks(seq, strict_layout_integrity=False) or []
                except (LayoutIntegrityError, TrackLayoutInvariantError) as exc:
                    return [], f"split:{exc}"
        return tracks, ""
    except Exception as exc:
        return [], f"{type(exc).__name__}: {exc}"
    finally:
        logging.disable(logging.NOTSET)


def track_row(window: str, source: str, pool: int | None, idx: int, date: str, sheet: str, stats: dict, poured_m: float) -> dict:
    meters = poured_m or stats["occupied_m"]
    return {
        "window": window,
        "source": source,
        "pool": pool,
        "idx": idx,
        "date": date,
        "sheet": sheet,
        "meters": round(meters, 2),
        "qty": stats["qty"],
        "composition": stats["composition"],
        "over_bc_m": round(stats["over_bc_m"], 1),
        "bc_pct": stats["bc_pct"],
        "concrete_rub": round(stats["concrete_rub"]),
        "delta_avg": round(stats["delta_avg"], 1),
        "delta_max": int(stats["delta_max"]) if stats["qty"] else 0,
        "over_wire_m": round(stats["over_wire_m"]),
        "wire_rub": round(stats["wire_rub"]),
        "total_rub": round(stats["wire_rub"] + stats["concrete_rub"]),
        "max_wires": stats["max_wires"],
        "grades": stats["grades"],
        "wire_per_m": round(per_m(stats["wire_rub"], meters), 2),
        "concrete_per_m": round(per_m(stats["concrete_rub"], meters), 2),
        "strip_m2": round(stats["strip_m2"], 2),
        "scrap_m2": round(stats["scrap_m2"], 2),
        "end_gap_m2": round(stats["end_gap_m2"], 2),
        "unused_m2": round(stats["strip_m2"] + stats["scrap_m2"] + stats["end_gap_m2"], 2),
    }


def aggregate(rows: list[dict], match: bool, note: str, elapsed: float) -> dict:
    meters = sum(r["meters"] for r in rows) or 0.0
    wire = sum(r["wire_rub"] for r in rows)
    concrete = sum(r["concrete_rub"] for r in rows)
    worst = max(rows, key=lambda r: r["wire_per_m"] + r["concrete_per_m"]) if rows else None
    return {
        "tracks": len(rows),
        "meters": round(meters, 1),
        "wire_rub": wire,
        "concrete_rub": concrete,
        "wire_per_m": round(per_m(wire, meters), 2),
        "concrete_per_m": round(per_m(concrete, meters), 2),
        "worst_wire_per_m": worst["wire_per_m"] if worst else 0,
        "worst_concrete_per_m": worst["concrete_per_m"] if worst else 0,
        "strip_m2": round(sum(r["strip_m2"] for r in rows), 2),
        "scrap_m2": round(sum(r["scrap_m2"] for r in rows), 2),
        "end_gap_m2": round(sum(r["end_gap_m2"] for r in rows), 2),
        "unused_m2": round(sum(r["unused_m2"] for r in rows), 2),
        "plates_match": match,
        "note": note,
        "seconds": round(elapsed, 1),
    }


def save_book(summary: list[dict], tracks: list[dict]) -> None:
    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    hdr_font = Font(color="FFFFFF", bold=True, size=10)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def header(ws, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=1, column=c)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
            cell.border = border

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Сводка пулов"
    h = ["Окно", "Источник", "Пул, дорожек", "Дорожек получилось", "Метраж, м",
         "Проволока, руб", "Бетон, руб", "Проволока, руб/м", "Бетон, руб/м",
         "Худшая дорожка, проволока руб/м", "Худшая дорожка, бетон руб/м",
         "Остаток ≥200 мм, м²", "Обрезок <200 мм, м²", "Хвост дорожки, м²", "Неиспользовано, м²",
         "Набор плит совпал", "Секунды", "Примечание"]
    ws.append(h)
    header(ws, len(h))
    for s in summary:
        ws.append([
            s["window"], s["source"], s["pool"], s["tracks"], s["meters"],
            s["wire_rub"], s["concrete_rub"], s["wire_per_m"], s["concrete_per_m"],
            s["worst_wire_per_m"], s["worst_concrete_per_m"],
            s["strip_m2"], s["scrap_m2"], s["end_gap_m2"], s["unused_m2"],
            "да" if s["plates_match"] else "нет", s["seconds"], s["note"],
        ])
    for col in range(1, len(h) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["R"].width = 42
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:R{ws.max_row}"

    ws2 = wb.create_sheet("Дорожки")
    h2 = ["Окно", "Источник", "Пул, дорожек", "№ в пуле", "Дата", "Дорожка",
          "Метраж, м", "Плит, шт", "Состав дорожки",
          "Перебетонир.: метраж, м", "% метража перебетонир.", "Цена перебетонир., руб",
          "Среднее переармир., проволок", "Макс переармир., проволок",
          "Лишняя проволока, м", "Цена переармир., руб", "Итого перерасход, руб",
          "Остаток ≥200 мм, м²", "Обрезок <200 мм, м²", "Хвост дорожки, м²", "Неиспользовано, м²"]
    ws2.append(h2)
    header(ws2, len(h2))
    pct = "0.0%"
    for r in tracks:
        ws2.append([
            r["window"], r["source"], r["pool"] if r["pool"] is not None else "карта",
            r["idx"], r["date"], r["sheet"], r["meters"], r["qty"], r["composition"],
            r["over_bc_m"], r["bc_pct"], r["concrete_rub"],
            r["delta_avg"], r["delta_max"], r["over_wire_m"], r["wire_rub"], r["total_rub"],
            r["strip_m2"], r["scrap_m2"], r["end_gap_m2"], r["unused_m2"],
        ])
    for row in ws2.iter_rows(min_row=2, min_col=11, max_col=11):
        for cell in row:
            cell.number_format = pct
    for col in range(1, len(h2) + 1):
        ws2.column_dimensions[get_column_letter(col)].width = 14
    ws2.freeze_panes = "A2"
    ws2.column_dimensions["I"].width = 55
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(h2))}{max(ws2.max_row, 1)}"

    ws3 = wb.create_sheet("Методика")
    for line in (
        "Режим укладки: сильные первыми (жадное склеивание выключено).",
        "Пул N — эта дорожка и N−1 следующие соседние, одним пулом через оптимизатор и раскладку «сильные первыми».",
        "Пул 5 — сама дорожка и 4 соседние. Размеры: 1, 5, 10, 20, 40, 60.",
        "Каждая получившаяся дорожка — отдельная строка в том же виде, что лист «Дорожки» отчёта КЗ.",
        "Проволока: (макс дорожки − норма плиты) × длина × 17.06 руб/м.",
        "Бетон: плита марки ниже максимума дорожки, объём × 131 руб/м³.",
        "Остаток: плиты одной длины стоят рядом в 1200 мм; зазор ≥200 мм × длина.",
        "Обрезок: тот же зазор, но уже 200 мм.",
        "Хвост: метраж заливки минус занятая длина, × 1.2 м.",
        "Карта — ноль: те же плиты в составе КЗ, без перераскладки.",
    ):
        ws3.append([line])
    ws3.column_dimensions["A"].width = 100

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)


def main() -> None:
    wires = load_wires()
    lookup = make_lookup(wires)
    tracks = json.load(open(CACHE))
    summary: list[dict] = []
    detail: list[dict] = []

    for name, start in WINDOWS:
        window = tracks[start:start + 60]
        print(f"\n=== {name} {window[0]['date']} … {window[-1]['date']} ===", flush=True)
        kz_rows = []
        kz_pieces = []
        for i, t in enumerate(window, 1):
            pcs = pieces_of(t, lookup)
            kz_pieces.extend(pcs)
            stats = money_and_waste(pcs, t.get("total_len"))
            row = track_row(name, "карта", None, i, t["date"], t["sheet"], stats, t.get("total_len") or stats["occupied_m"])
            kz_rows.append(row)
            detail.append(row)
        summary.append({"window": name, "source": "карта", "pool": 60, **aggregate(kz_rows, True, "все 60 дорожек окна", 0)})
        print(
            f"  карта 60: проволока {summary[-1]['wire_per_m']} руб/м, "
            f"бетон {summary[-1]['concrete_per_m']} руб/м, "
            f"неиспользовано {summary[-1]['unused_m2']} м²",
            flush=True,
        )

        for n in SIZES:
            chunk = window[:n]
            pcs = [p for t in chunk for p in pieces_of(t, lookup)]
            want = signature(pcs)
            t0 = time.perf_counter()
            laid, err = relayout(pcs)
            elapsed = time.perf_counter() - t0
            rows = []
            got_pieces: list[dict] = []
            extra_note = err
            if laid:
                for i, tr in enumerate(laid, 1):
                    lp, extra_scrap = layout_pieces(tr, lookup)
                    got_pieces.extend(lp)
                    stats = money_and_waste(lp, tr.get("length"))
                    stats["scrap_m2"] += extra_scrap
                    rows.append(track_row(
                        name, "сильные первыми", n, i, "", tr.get("label") or "",
                        stats, tr.get("length") or stats["occupied_m"],
                    ))
                detail.extend(rows)
                match = signature(got_pieces) == want
                if not match:
                    extra_note = (extra_note + "; " if extra_note else "") + "набор плит отличается от карт"
            else:
                match = False
            agg = aggregate(rows, match, extra_note, elapsed)
            summary.append({"window": name, "source": "сильные первыми", "pool": n, **agg})
            print(
                f"  пул {n:2d}: дорожек {agg['tracks']:3d}  "
                f"проволока {agg['wire_per_m']:8}  бетон {agg['concrete_per_m']:8}  "
                f"остаток {agg['strip_m2']:8}  обрезок {agg['scrap_m2']:8}  "
                f"хвост {agg['end_gap_m2']:8}  совпал={agg['plates_match']}  {elapsed:.0f}с"
                + (f"  {extra_note}" if extra_note else ""),
                flush=True,
            )
            JSON_OUT.write_text(json.dumps({"summary": summary, "tracks": detail}, ensure_ascii=False), encoding="utf-8")
            save_book(summary, detail)

    print(f"\n[OK] {OUT}", flush=True)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        raise
