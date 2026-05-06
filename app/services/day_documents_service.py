"""Генерация производственных документов по конкретному дню.

Повторяет логику из `bot/handlers/production_day_view.py`:
- схема (PDF) через `core.visualization.visualize_plan`;
- детальная разбивка (XLSX) — побочный файл той же функции;
- формовка (XLSX по дорожке) через `core.formovka_excel.create_formovka_files_for_tracks`;
  файлы формовки упаковываются в ZIP.

Сервис работает во временных папках, обязанность очистки лежит на вызывающем коде
(через FastAPI `BackgroundTasks`). `visualize_plan` использует глобали из
`core.optimization` / `core.config_and_data`, поэтому параллельные вызовы
сериализуются через `asyncio.Lock`.
"""
from __future__ import annotations

import asyncio
import copy
import logging
import os
import shutil
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder as AppPlateOrder
import core.optimization as optimization
from core.optimization.result_contract import is_optimization_success
from core.config_and_data import PlateOrder
from core.formovka_excel import create_formovka_files_for_tracks
from core.visualization import visualize_plan

from app.services.day_view_service import (
    _aggregate_plates_for_track,
    _build_smart_lookup,
)
from bot.handlers import plan_manager

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FORMOVKA_TEMPLATE_PATH = PROJECT_ROOT / "банк знаний" / "!КЗ ПБ Шаблон.xlsx"

# visualize_plan правит глобальные переменные — сериализуем вызовы,
# чтобы не словить гонку при параллельных запросах.
_visualize_lock = asyncio.Lock()


class DayDocumentsError(RuntimeError):
    """Ошибка генерации документа для дня."""


def _load_day_bundle(target_date: str) -> dict:
    multi = plan_manager.get_tracks_for_date_from_all_plans(target_date)
    if not multi or not multi.get("tracks"):
        raise DayDocumentsError(f"На дату {target_date} нет дорожек ни в одном плане")
    return multi


def _restore_optimization_globals(orders_2d: list, optimization_result: dict) -> None:
    """Ставит глобальные переменные оптимизации, нужные `visualize_plan`."""
    app_order = AppPlateOrder.from_orders_2d(orders_2d)
    if orders_2d:
        load_codes = sorted({p.get("load_code", 8) for p in orders_2d})
        load_map = {code: ["all"] for code in load_codes}
    else:
        load_map = {8: ["all"]}
    context = OptimizationContext(
        order=app_order,
        optimization_result=optimization_result,
        plan_by_load={"all": optimization_result} if is_optimization_success(optimization_result) else {},
        load_to_reinforcement_map=load_map,
    )
    optimization.OPT_CASCADING_PLAN = context.optimization_result
    optimization.OPT_CASCADING_PLAN_BY_LOAD = context.plan_by_load
    optimization.LOAD_TO_REINFORCEMENT_MAP = context.load_to_reinforcement_map
    PlateOrder.from_dict(app_order.to_dict()).apply_to_globals()


def _run_visualize(existing_tracks: list, output_dir: Path) -> tuple[str | None, str | None]:
    return visualize_plan(
        output_dir=str(output_dir),
        tracks_per_file=None,
        start_track_index=0,
        use_production_pricing=True,
        existing_tracks=existing_tracks,
    )


def _find_breakdown(output_dir: Path) -> Path | None:
    if not output_dir.exists():
        return None
    candidates = [
        p
        for p in output_dir.iterdir()
        if p.is_file() and p.name.startswith("Детальная_разбивка_") and p.suffix == ".xlsx"
    ]
    if not candidates:
        return None
    candidates.sort(key=lambda p: p.stat().st_ctime, reverse=True)
    return candidates[0]


def _build_formovka_tracks_data(multi: dict) -> list[dict[str, Any]]:
    """Готовит data для `create_formovka_files_for_tracks` на основе тех же
    агрегированных плит, что и в `DayViewDetailResponse` (fuzzy lookup)."""
    lookup = _build_smart_lookup(
        multi.get("plate_lookup_exact", {}),
        multi.get("plate_lookup_by_length", {}),
    )

    tracks_data: list[dict[str, Any]] = []
    for index, track in enumerate(multi.get("tracks") or [], start=1):
        plates = _aggregate_plates_for_track(track, lookup)
        if not plates:
            continue
        tracks_data.append(
            {
                "track_number": index,
                "max_reinforcement": float(track.get("max_reinforcement") or 0),
                "plates_info": [
                    {
                        "length": p["length_m"],
                        "width": p["width_mm"],
                        "qty": p["qty"],
                        "reinforcement": p["reinforcement"],
                        "kp_date": p["kp_date"],
                        "customer": p["customer"],
                        "kp_id": p.get("kp_id"),
                        "plate_name": p["plate_name"],
                    }
                    for p in plates
                ],
            }
        )
    return tracks_data


def _make_tmp_dir(prefix: str) -> Path:
    path = Path(tempfile.mkdtemp(prefix=prefix))
    return path


def _cleanup_dir(path: Path) -> None:
    try:
        shutil.rmtree(path, ignore_errors=True)
    except Exception:
        logger.exception("[day-docs] Не удалось удалить временную папку %s", path)


async def generate_day_schema(target_date: str) -> tuple[Path, Path]:
    """Возвращает (pdf_path, cleanup_dir). Вызывающий код отвечает за удаление dir."""
    multi = _load_day_bundle(target_date)
    tmp_dir = _make_tmp_dir(prefix=f"day_schema_{target_date}_")
    try:
        async with _visualize_lock:
            _restore_optimization_globals(
                multi.get("orders_2d", []),
                multi.get("optimization_result", {}),
            )
            result = await asyncio.to_thread(
                _run_visualize, copy.deepcopy(multi["tracks"]), tmp_dir
            )
        if not isinstance(result, tuple) or len(result) < 2:
            raise DayDocumentsError("visualize_plan не вернул PDF")
        _png_path, pdf_path = result
        if not pdf_path or not os.path.exists(pdf_path):
            raise DayDocumentsError("PDF со схемой не был создан")
        return Path(pdf_path), tmp_dir
    except Exception:
        _cleanup_dir(tmp_dir)
        raise


async def generate_day_breakdown(target_date: str) -> tuple[Path, Path]:
    """Возвращает (xlsx_path, cleanup_dir)."""
    multi = _load_day_bundle(target_date)
    tmp_dir = _make_tmp_dir(prefix=f"day_breakdown_{target_date}_")
    try:
        async with _visualize_lock:
            _restore_optimization_globals(
                multi.get("orders_2d", []),
                multi.get("optimization_result", {}),
            )
            await asyncio.to_thread(
                _run_visualize, copy.deepcopy(multi["tracks"]), tmp_dir
            )
        breakdown = _find_breakdown(tmp_dir)
        if breakdown is None:
            raise DayDocumentsError("Файл детальной разбивки не найден")
        return breakdown, tmp_dir
    except Exception:
        _cleanup_dir(tmp_dir)
        raise


async def generate_day_formovka(target_date: str) -> tuple[Path, Path]:
    """Возвращает (zip_path, cleanup_dir)."""
    multi = _load_day_bundle(target_date)
    tracks_data = _build_formovka_tracks_data(multi)
    if not tracks_data:
        raise DayDocumentsError("Нет данных для формовки: ни одной плиты в дорожках")

    tmp_dir = _make_tmp_dir(prefix=f"day_formovka_{target_date}_")
    try:
        date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        template_path = str(FORMOVKA_TEMPLATE_PATH) if FORMOVKA_TEMPLATE_PATH.exists() else None

        def _run_formovka() -> list[str]:
            return create_formovka_files_for_tracks(
                tracks_data,
                str(tmp_dir),
                start_track_number=1,
                template_path=template_path,
                date_str=date_str,
            )

        formovka_files: list[str] = await asyncio.to_thread(_run_formovka)
        existing = [Path(p) for p in formovka_files if p and os.path.exists(p)]
        if not existing:
            raise DayDocumentsError("Файлы формовки не были созданы")

        zip_path = tmp_dir / f"Формовка_{target_date}.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file_path in existing:
                zf.write(file_path, arcname=file_path.name)
        return zip_path, tmp_dir
    except Exception:
        _cleanup_dir(tmp_dir)
        raise


def make_cleanup_callback(cleanup_dir: Path):
    """Возвращает callable для `BackgroundTasks.add_task`, удаляющий временную папку."""

    def _cleanup() -> None:
        _cleanup_dir(cleanup_dir)

    return _cleanup
