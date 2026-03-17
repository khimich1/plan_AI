"""
Сквозной аудит плит по этапам производственного конвейера.

Использование:
    audit = PlateAudit(orders_2d)
    audit.checkpoint("demand_2d", demand_2d_dict)         # {(L,W,lc): qty}
    audit.checkpoint("solver", primary_cuts_list)         # list of cut dicts
    audit.checkpoint("post_correction", primary_cuts)
    audit.checkpoint("layout_sequence", sequence_list)    # list of item dicts
    audit.checkpoint("tracks", tracks_list)               # list of track dicts
    audit.checkpoint("final", tracks_list)
    logger.info(audit.summary())
"""

from __future__ import annotations

import logging
from collections import Counter
from typing import Any

from .config_and_data import canonical_plate_key

logger = logging.getLogger(__name__)


class PlateAudit:
    """
    Сквозной счётчик плит по именованным этапам конвейера.

    Каждый checkpoint сохраняет количество плит по canonical_plate_key.
    После добавления нескольких этапов можно получить diff и summary.
    """

    def __init__(self, orders_2d: list[dict]) -> None:
        self.checkpoints: dict[str, Counter] = {}
        self._stage_order: list[str] = []
        self.checkpoint("input", orders_2d)

    # ------------------------------------------------------------------
    # Запись этапов
    # ------------------------------------------------------------------

    def checkpoint(self, stage: str, data: Any) -> None:
        """
        Записать количество плит на данном этапе.

        data может быть:
        - list[dict] с orders_2d (поля length, width, load_code, qty)
        - dict {(L, W, lc): qty}  — demand_2d или похожее
        - list[dict] с primary_cuts / plate_assignments (поля length/lengths, width, load_code)
        - list[dict] с track дорожками (поле items -> list of track items)
        - list[dict] с sequence items напрямую
        """
        counts: Counter = Counter()
        if isinstance(data, dict):
            for key, qty in data.items():
                if isinstance(key, tuple) and len(key) == 3:
                    ck = canonical_plate_key(*key)
                    counts[ck] += qty
        elif isinstance(data, list):
            for item in data:
                if not isinstance(item, dict):
                    continue
                # Дорожка (track): содержит поле 'items'
                if 'items' in item:
                    for sub in (item.get('items') or []):
                        self._count_track_item(sub, counts)
                # Sequence-группа: содержит поле 'sequence'
                elif 'sequence' in item:
                    for sub in (item.get('sequence') or []):
                        self._count_track_item(sub, counts)
                # Одиночный cut/assignment/order
                else:
                    self._count_cut_or_order(item, counts)

        self.checkpoints[stage] = counts
        if stage not in self._stage_order:
            self._stage_order.append(stage)
        logger.debug(
            "[AUDIT] Checkpoint %-20s: %d ключей, %d плит",
            stage, len(counts), sum(counts.values()),
        )

    # ------------------------------------------------------------------
    # Анализ
    # ------------------------------------------------------------------

    def lost(self, stage_from: str, stage_to: str) -> Counter:
        """
        Плиты, которые были на stage_from, но исчезли на stage_to.
        Возвращает Counter {canonical_key: недостача}.
        """
        before = self.checkpoints.get(stage_from, Counter())
        after = self.checkpoints.get(stage_to, Counter())
        diff: Counter = Counter()
        for key, qty in before.items():
            have = after.get(key, 0)
            if have < qty:
                diff[key] = qty - have
        return diff

    def summary(self) -> str:
        """Текстовый отчёт по всем зафиксированным этапам."""
        lines = ["[PlateAudit] Сводка по этапам конвейера:"]
        for stage in self._stage_order:
            counts = self.checkpoints[stage]
            total = sum(counts.values())
            lines.append(f"  {stage:<22}: {total:>4} плит ({len(counts)} ключей)")

        # Показываем потери между соседними этапами
        stages = self._stage_order
        loss_lines = []
        for i in range(1, len(stages)):
            diff = self.lost(stages[i - 1], stages[i])
            if diff:
                lost_total = sum(diff.values())
                loss_lines.append(
                    f"  ПОТЕРЯ {stages[i-1]} -> {stages[i]}: {lost_total} плит"
                )
                for key, qty in diff.most_common(5):
                    loss_lines.append(f"    - {key}: -{qty}")

        if loss_lines:
            lines.append("[PlateAudit] ОБНАРУЖЕНЫ ПОТЕРИ:")
            lines.extend(loss_lines)
        else:
            lines.append("[PlateAudit] Потерь не обнаружено.")

        return "\n".join(lines)

    def has_losses(self, stage_from: str = "input", stage_to: str = "final") -> bool:
        """True если хотя бы одна плита потерялась между двумя этапами."""
        return bool(self.lost(stage_from, stage_to))

    # ------------------------------------------------------------------
    # Вспомогательные методы разбора форматов
    # ------------------------------------------------------------------

    def _count_cut_or_order(self, item: dict, counts: Counter) -> None:
        """Разобрать один cut/order/assignment и добавить в counts."""
        qty = item.get('qty', 1)
        width = item.get('width') or item.get('demand_width')
        if width is None:
            return

        # Формат primary_cut: поле 'lengths' (множественное) содержит список длин
        lengths_raw = item.get('lengths')
        if lengths_raw:
            for L in lengths_raw:
                ck = canonical_plate_key(L, width, item.get('load_code', 8))
                counts[ck] += 1
            return

        # Формат orders_2d / plate_assignment: одиночное поле 'length'
        if 'length' in item:
            ck = canonical_plate_key(item['length'], width, item.get('load_code', 8))
            counts[ck] += int(qty)

    def _count_track_item(self, item: dict, counts: Counter) -> None:
        """Разобрать один элемент трека."""
        if not isinstance(item, dict):
            return
        load_code = item.get('load_code', 8)
        mode = item.get('mode', 'solid')
        length = item.get('length', 0) or 0
        width_raw = item.get('width', item.get('main_w', 1.2))

        # Конвертируем ширину из метров в мм, если нужно
        width_mm = _to_width_mm_audit(width_raw)

        if mode == 'transverse':
            target_length = item.get('target_length') or length
            if target_length > 0:
                ck = canonical_plate_key(target_length, width_mm, load_code)
                counts[ck] += 1
            remainder = item.get('remainder', 0) or 0
            if remainder > 0.1:
                ck = canonical_plate_key(remainder, width_mm, load_code)
                counts[ck] += 1
        else:
            if length > 0:
                ck = canonical_plate_key(length, width_mm, load_code)
                counts[ck] += 1

        # Вторичные резы внутри элемента
        for sec in (item.get('secondary_cuts') or []):
            if isinstance(sec, dict):
                sec_w = _to_width_mm_audit(sec.get('width', 0))
                sec_len = sec.get('target_length') or sec.get('length') or length
                if sec_w > 0 and sec_len > 0:
                    ck = canonical_plate_key(sec_len, sec_w, load_code)
                    counts[ck] += 1


def _to_width_mm_audit(w: Any, default_m: float = 1.2) -> int:
    """Конвертирует ширину к мм: если < 20 — считаем метры × 1000, иначе уже мм."""
    try:
        f = float(w)
    except (TypeError, ValueError):
        f = default_m
    return round(f * 1000) if f < 20 else round(f)
