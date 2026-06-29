"""Track layout: split sequence into tracks and integrity checks."""
from __future__ import annotations

import logging
from collections import Counter

from core.config_and_data import format_reinforcement_from_load_code
from core.domain.plate_order import normalize_load_code

__all__ = [
    "LayoutIntegrityError",
    "TrackLayoutInvariantError",
    "split_sequence_into_tracks",
    "validate_track_integrity",
]

logger = logging.getLogger(__name__)


class LayoutIntegrityError(RuntimeError):
    """Ошибка целостности раскладки: часть плит потеряна или задублирована."""


class TrackLayoutInvariantError(RuntimeError):
    """Нельзя выполнить правило начала дорожки с целой плиты без изменения плана/заказа."""


_TRACK_LEN_EPS = 1e-6


def _starter_solid_tiers(
    items: list[dict],
    anchor_idx: int,
    *,
    scope: str = "both",
) -> list[tuple[str, list[int]]]:
    """
    Ярусы индексов целых плит относительно anchor_idx.

    scope='both': при пустой дорожке (до/после якоря).
    scope='after_only': только j > anchor при переносе сплита со следующей дорожкой,
        чтобы не взять целую уже из закрываемой дорожки (одна layout_uid дважды).
    """
    n = len(items)
    tiers_ns_after = [
        j
        for j in range(anchor_idx + 1, n)
        if items[j].get("mode") == "solid" and not items[j].get("is_separator", False)
    ]
    tiers_s_after = [
        j
        for j in range(anchor_idx + 1, n)
        if items[j].get("mode") == "solid" and items[j].get("is_separator", False)
    ]
    if scope == "after_only":
        return [
            ("after_ns", tiers_ns_after),
            ("after_sep", tiers_s_after),
        ]
    tiers_ns_before = [
        j
        for j in range(anchor_idx)
        if items[j].get("mode") == "solid" and not items[j].get("is_separator", False)
    ]
    tiers_s_before = [
        j
        for j in range(anchor_idx)
        if items[j].get("mode") == "solid" and items[j].get("is_separator", False)
    ]
    return [
        ("after_ns", tiers_ns_after),
        ("before_ns", tiers_ns_before),
        ("after_sep", tiers_s_after),
        ("before_sep", tiers_s_before),
    ]


def _pick_track_starter_solid_index(
    items: list[dict],
    anchor_idx: int,
    *,
    exclude_indices: frozenset[int],
    neighbor_reinf: float,
    max_track_length: float,
    pref_reinf: bool,
    max_reinf_cap_strict: float,
    paired_partner_length: float | None,
    allow_reinf_relaxation: bool,
    group_label_for_log: str,
    tier_scope: str = "both",
) -> int | None:
    """
    Возвращает индекс целой для старта дорожки или None если нельзя без нарушения ограничений.
    paired_partner_length: если задан (длина сплита вторым на дорожке), целая + партнёр ≤ max_track_length.
    """

    def _length_ok(idx: int) -> bool:
        if paired_partner_length is None:
            return True
        return (
            float(items[idx]["length"]) + float(paired_partner_length)
            <= max_track_length + _TRACK_LEN_EPS
        )

    def _tier_pick(
        apply_reinf_cap: bool,
    ) -> tuple[int | None, str | None]:
        max_cap = float(max_reinf_cap_strict or 0.0)
        for tier_name, tier in _starter_solid_tiers(items, anchor_idx, scope=tier_scope):
            eligible: list[int] = []
            for j in tier:
                if j in exclude_indices:
                    continue
                if items[j].get("mode") != "solid":
                    continue
                if not _length_ok(j):
                    continue
                if apply_reinf_cap and max_cap > 0:
                    cand_reinf = float(items[j].get("reinforcement", 0) or 0)
                    if cand_reinf > max_cap:
                        continue
                eligible.append(j)
            if not eligible:
                continue
            if pref_reinf:
                pick_idx = min(
                    eligible,
                    key=lambda jj, _nr=float(neighbor_reinf): (
                        abs(float(items[jj].get("reinforcement", 0) or 0) - _nr),
                        1 if items[jj].get("is_separator") else 0,
                        jj,
                    ),
                )
            else:
                pick_idx = min(eligible)
            return pick_idx, tier_name
        return None, None

    pick, tier_logged = _tier_pick(True)
    if pick is not None:
        logger.info(
            "[SPLIT_TRACKS] Стартовая целая (%s phase=strict): индекс %s",
            tier_logged,
            pick,
        )
        return pick

    if allow_reinf_relaxation:
        pick2, tier2 = _tier_pick(False)
        if pick2 is not None:
            logger.warning(
                "[SPLIT_TRACKS] Стартовая целая (%s): допущение по армированию "
                "(LAYOUT_TRACK_START_REINF_RELAXATION) для группы %s индекс %s",
                tier2,
                group_label_for_log,
                pick2,
            )
            return pick2

    return None


def _assert_track_starts_with_solid(tracks: list[dict]) -> None:
    """Пост-условие: первая плита дорожки — целая (mode solid). Пустые дорожки не ожидаются."""
    for ti, track in enumerate(tracks or []):
        if not isinstance(track, dict):
            continue
        its = track.get("items") or []
        if not its:
            continue
        first = its[0]
        if first.get("mode") != "solid":
            raise TrackLayoutInvariantError(
                f"Дорожка #{ti + 1}: первый элемент не целая плита (mode={first.get('mode')!r})"
            )


def _count_solids_remaining(items: list[dict]) -> int:
    return sum(1 for it in items if it.get("mode") == "solid")


def _iter_sequence_items(sequence: list) -> list[dict]:
    """Возвращает flat-список root items из grouped/flat sequence."""
    if not isinstance(sequence, list):
        return []
    if sequence and isinstance(sequence[0], dict) and isinstance(sequence[0].get("sequence"), list):
        items: list[dict] = []
        for group in sequence:
            if not isinstance(group, dict):
                continue
            for item in group.get('sequence') or []:
                if isinstance(item, dict):
                    items.append(item)
        return items
    return [item for item in sequence if isinstance(item, dict)]


def _ensure_layout_uid(items: list[dict], *, prefix: str) -> None:
    """Проставляет стабильный uid у элементов, если он ещё не задан."""
    for idx, item in enumerate(items):
        if item.get("layout_uid"):
            continue
        unit_id = item.get("unit_id")
        if unit_id:
            item["layout_uid"] = str(unit_id)
            continue
        # fallback uid нужен только для контроля целостности split.
        item["layout_uid"] = f"{prefix}:{idx}"


def validate_track_integrity(
    input_sequence: list,
    tracks: list[dict],
    *,
    strict: bool = False,
) -> dict:
    """Сверяет identity входной последовательности и выхода split по layout_uid."""
    in_items = _iter_sequence_items(input_sequence)
    _ensure_layout_uid(in_items, prefix="seq")
    input_ids = [str(item.get("layout_uid")) for item in in_items]

    out_items: list[dict] = []
    for track in tracks or []:
        if not isinstance(track, dict):
            continue
        for item in track.get("items") or []:
            if isinstance(item, dict):
                out_items.append(item)
    _ensure_layout_uid(out_items, prefix="track")
    output_ids = [str(item.get("layout_uid")) for item in out_items]

    input_counter = Counter(input_ids)
    output_counter = Counter(output_ids)
    missing = dict(input_counter - output_counter)
    duplicated = dict(output_counter - input_counter)

    report = {
        "input_total": len(input_ids),
        "output_total": len(output_ids),
        "missing": missing,
        "duplicated": duplicated,
        "ok": not missing and not duplicated,
    }
    if strict and not report["ok"]:
        raise LayoutIntegrityError(
            "Нарушена целостность split_sequence_into_tracks: "
            f"input={report['input_total']}, output={report['output_total']}, "
            f"missing={len(missing)}, duplicated={len(duplicated)}"
        )
    return report


def split_sequence_into_tracks(
    sequence: list,
    max_track_length: float = 101.0,
    strict_layout_integrity: bool = False,
    *,
    track_reinf_preference: bool | None = None,
    track_start_reinf_relaxation: bool | None = None,
    require_solid_track_start: bool = True,
) -> list:
    """
    Разбивает последовательность плит на дорожки с соблюдением правил завода.
    
    ПРАВИЛА:
    1. Каждая дорожка ДОЛЖНА начинаться с ЦЕЛОЙ плиты (без реза)
    2. Дорожка не должна превышать max_track_length (101м)
    3. При превышении ищет целую плиту для завершения
    
    Args:
        sequence: Результат build_layout_sequence() - может быть:
                  - list[dict] с группами по нагрузкам [{'load_code', 'sequence', 'label'}, ...]
                  - list[dict] с плитами [{'length', 'mode', ...}, ...]
        max_track_length: Максимальная длина дорожки (101м)
        track_reinf_preference: Если True — среди допустимых целых выбирать ближайшую по армированию
            к соседней плите с резом; None — из Settings (LAYOUT_TRACK_REINF_PREFERENCE).
        track_start_reinf_relaxation: Если True и строгое армирование не даёт целую для старта —
            разрешить кандидата с армированием выше порога дорожки; None — из Settings.
        require_solid_track_start: Если True — перед возвратом проверять пост-условие «первая плита solid».

    Returns:
        list[dict]: Список дорожек [{'items': [...], 'length': float, 'load_code': int, 'label': str, 'max_reinforcement': float}, ...]
    Raises:
        TrackLayoutInvariantError: нельзя выбрать целую без потери элементов или смены плана.
    """
    def _label_from_track_items(items, fallback):
        """Подпись дорожки по фактическим нагрузкам плит в ней (чтобы при группе 'all' не было везде 8п)."""
        load_codes = set()
        for item in items:
            lc = item.get('load_code')
            if lc is not None:
                try:
                    n = normalize_load_code(lc)
                    if n is not None:
                        load_codes.add(n)
                except Exception:
                    pass
        if not load_codes:
            return fallback
        disp = [format_reinforcement_from_load_code(lc) for lc in sorted(load_codes)]
        return 'Нагрузка ' + ', '.join(disp) if len(disp) > 1 else f'Нагрузка {disp[0]}'

    tracks = []
    _sequence_items_for_uid = _iter_sequence_items(sequence)
    _ensure_layout_uid(_sequence_items_for_uid, prefix="seq")

    if track_reinf_preference is None:
        try:
            from core.config.settings import get_settings

            track_reinf_preference = get_settings().layout_track_reinf_preference
        except Exception:
            track_reinf_preference = False
    _pref_reinf = bool(track_reinf_preference)

    if track_start_reinf_relaxation is None:
        try:
            from core.config.settings import get_settings

            track_start_reinf_relaxation = get_settings().layout_track_start_reinf_relaxation
        except Exception:
            track_start_reinf_relaxation = True
    _relax_reinf_start = bool(track_start_reinf_relaxation)
    
    # Проверяем формат данных (с группировкой по нагрузке или без)
    _first = sequence[0] if isinstance(sequence, list) and sequence else None
    if (
        isinstance(_first, dict)
        and "load_code" in _first
        and isinstance(_first.get("sequence"), list)
    ):
        logger.info(f"[SPLIT_TRACKS] Обнаружена группировка по нагрузкам. Групп: {len(sequence)}")
        input_count = sum(len(g['sequence']) for g in sequence)
        # Каждая группа нагрузки = отдельные дорожки
        for group in sequence:
            load_code = group['load_code']
            items = list(group['sequence'])  # Копия для возможной модификации
            group_label = group.get('label', f'Нагрузка {load_code}п')
            
            logger.info(f"[SPLIT_TRACKS] Группа '{group_label}': {len(items)} плит")
            
            # Разбиваем группу на дорожки по 96-101м
            # ЖЁСТКОЕ ПРАВИЛО: дорожка НИКОГДА не превышает max_track_length!
            current_track = []
            current_track_length = 0.0
            max_reinforcement_in_track = 0.0  # Макс. армирование в текущей дорожке
            
            i = 0
            while i < len(items):
                item = items[i]
                item_length = item['length']
                is_solid = (item.get('mode') == 'solid')
                item_reinforcement = item.get('reinforcement', 0) or 0
                
                # 🔥 ПРАВИЛО: Если дорожка пустая и текущая плита НЕ целая - 
                # сначала найти целую плиту для НАЧАЛА дорожки!
                # РАСШИРЕННЫЙ ПОИСК: ищем по ВСЕМУ списку с приоритетами
                if not current_track and not is_solid:
                    neighbor_reinf = float(item_reinforcement or 0)
                    found_solid_idx = _pick_track_starter_solid_index(
                        items,
                        i,
                        exclude_indices=frozenset({i}),
                        neighbor_reinf=neighbor_reinf,
                        max_track_length=max_track_length,
                        pref_reinf=_pref_reinf,
                        max_reinf_cap_strict=0.0,
                        paired_partner_length=float(item_length),
                        allow_reinf_relaxation=_relax_reinf_start,
                        group_label_for_log=str(group_label),
                    )
                    if found_solid_idx is None:
                        if _relax_reinf_start:
                            msg = (
                                f"Группа «{group_label}»: нет целой плиты для начала дорожки "
                                f"(элементов={len(items)}, целых={_count_solids_remaining(items)}, "
                                f"текущий mode={item.get('mode')!r}, длина={item_length} м; "
                                f"вместе со сплитом до {max_track_length} м нет подходящей целой). "
                                "Измените план или заказ."
                            )
                        else:
                            msg = (
                                f"Группа «{group_label}»: нет целой для старта дорожки при "
                                "LAYOUT_TRACK_START_REINF_RELAXATION=false. "
                                f"Целых в очереди: {_count_solids_remaining(items)}."
                            )
                        raise TrackLayoutInvariantError(msg)
                    if items[found_solid_idx].get("is_separator", False):
                        logger.warning(
                            "[SPLIT_TRACKS] Используем разделитель для начала дорожки (не идеально)"
                        )
                    solid_plate = items.pop(found_solid_idx)
                    is_sep = solid_plate.get('is_separator', False)
                    logger.info(
                        f"[SPLIT_TRACKS] Найдена целая плита для НАЧАЛА дорожки: {solid_plate['length']:.2f}м (разделитель={is_sep})"
                    )
                    current_track.append(solid_plate)
                    current_track_length += solid_plate['length']
                    solid_reinf = solid_plate.get('reinforcement', 0) or 0
                    max_reinforcement_in_track = max(max_reinforcement_in_track, solid_reinf)
                    if found_solid_idx < i:
                        i -= 1
                    continue
                
                # Проверяем: добавление плиты превысит максимум?
                will_exceed_max = (current_track_length + item_length > max_track_length and current_track)
                
                if will_exceed_max:
                    # Нужно закрыть дорожку - ЖЁСТКО не превышаем max_track_length!
                    remaining_space = max_track_length - current_track_length
                    
                    if is_solid and item_length <= remaining_space:
                        # Текущая плита целая и влезает - добавляем её
                        current_track.append(item)
                        current_track_length += item_length
                        max_reinforcement_in_track = max(max_reinforcement_in_track, item_reinforcement)
                        i += 1
                    elif not is_solid:
                        # Плита с резом не влезает: не теряем её — переносим в следующую дорожку.
                        # Ищем целую плиту: используем её как ПЕРВУЮ в новой дорожке, плиту с резом — второй.
                        split_plate = item
                        split_reinf = float(split_plate.get("reinforcement", 0) or 0)
                        cap = float(max_reinforcement_in_track or 0.0)
                        found_solid_idx = _pick_track_starter_solid_index(
                            items,
                            i,
                            exclude_indices=frozenset({i}),
                            neighbor_reinf=split_reinf,
                            max_track_length=max_track_length,
                            pref_reinf=_pref_reinf,
                            max_reinf_cap_strict=cap,
                            paired_partner_length=float(split_plate["length"]),
                            allow_reinf_relaxation=_relax_reinf_start,
                            group_label_for_log=str(group_label),
                            tier_scope="after_only",
                        )
                        if found_solid_idx is None:
                            if _relax_reinf_start:
                                raise TrackLayoutInvariantError(
                                    f"Группа «{group_label}»: нет целой для начала следующей дорожки "
                                    f"со сплитом {split_plate['length']:.2f} м (дорожка ≤{max_track_length} м, "
                                    f"порог арм. закрываемой дорожки={cap}). "
                                    f"Целых в очереди: {_count_solids_remaining(items)}. Измените план или заказ."
                                )
                            raise TrackLayoutInvariantError(
                                f"Группа «{group_label}»: включите LAYOUT_TRACK_START_REINF_RELAXATION "
                                "или ослабьте ограничения плана — нет целой для старта следующей дорожки."
                            )
                        if items[found_solid_idx].get("is_separator", False):
                            logger.warning(
                                "[SPLIT_TRACKS] Используем разделитель для начала следующей дорожки (не идеально)"
                            )

                        candidate = items.pop(found_solid_idx)
                        # Закрываем текущую дорожку БЕЗ добавления candidate (он пойдёт в новую).
                        logger.info(
                            f"[SPLIT_TRACKS] Плита с резом {split_plate['length']:.2f}м переносится в следующую дорожку, "
                            f"первой — целая {candidate['length']:.2f}м"
                        )
                        split_idx = i if found_solid_idx > i else i - 1
                        items.pop(split_idx)
                        if found_solid_idx < i:
                            _i_before = i
                            i = found_solid_idx
                        track_label = _label_from_track_items(current_track, group_label)
                        tracks.append({
                            'items': current_track,
                            'length': current_track_length,
                            'load_code': load_code,
                            'label': track_label,
                            'max_reinforcement': max_reinforcement_in_track
                        })
                        current_track = [candidate, split_plate]
                        current_track_length = candidate['length'] + split_plate['length']
                        max_reinforcement_in_track = max(max_reinforcement_in_track, candidate.get('reinforcement', 0) or 0, split_plate.get('reinforcement', 0) or 0)
                        continue
                    else:
                        # Целая плита, но не влезает
                        pass
                    
                    # Закрываем дорожку
                    logger.info(
                        f"[SPLIT_TRACKS] Закрываем дорожку на {current_track_length:.1f}м (макс. {max_track_length}м)"
                    )
                    track_label = _label_from_track_items(current_track, group_label)
                    tracks.append({
                        'items': current_track,
                        'length': current_track_length,
                        'load_code': load_code,
                        'label': track_label,
                        'max_reinforcement': max_reinforcement_in_track
                    })
                    current_track = []
                    current_track_length = 0.0
                    max_reinforcement_in_track = 0.0
                    
                    if is_solid and item_length <= remaining_space:
                        continue
                    continue
                
                # Плита влезает - добавляем
                current_track.append(item)
                current_track_length += item_length
                max_reinforcement_in_track = max(max_reinforcement_in_track, item_reinforcement)
                i += 1
            
            # Сохраняем последнюю дорожку группы
            if current_track:
                track_label = _label_from_track_items(current_track, group_label)
                tracks.append({
                    'items': current_track,
                    'length': current_track_length,
                    'load_code': load_code,
                    'label': track_label,
                    'max_reinforcement': max_reinforcement_in_track
                })
    
    else:
        # СТАРЫЙ ФОРМАТ (без группировки по нагрузке)
        logger.info("[SPLIT_TRACKS] Используем старый формат (без группировки)")
        input_count = len(sequence) if sequence else 0
        items = list(sequence)  # Копия для возможной модификации
        current_track = []
        current_track_length = 0.0
        max_reinforcement_in_track = 0.0
        
        i = 0
        while i < len(items):
            item = items[i]
            item_length = item['length']
            is_solid = (item.get('mode') == 'solid')
            item_reinforcement = item.get('reinforcement', 0) or 0
            
            # 🔥 ПРАВИЛО: Если дорожка пустая и текущая плита НЕ целая
            if not current_track and not is_solid:
                neighbor_reinf = float(item_reinforcement or 0)
                _flat_label = "flat"
                found_solid_idx = _pick_track_starter_solid_index(
                    items,
                    i,
                    exclude_indices=frozenset({i}),
                    neighbor_reinf=neighbor_reinf,
                    max_track_length=max_track_length,
                    pref_reinf=_pref_reinf,
                    max_reinf_cap_strict=0.0,
                    paired_partner_length=float(item_length),
                    allow_reinf_relaxation=_relax_reinf_start,
                    group_label_for_log=_flat_label,
                )
                if found_solid_idx is None:
                    if _relax_reinf_start:
                        msg = (
                            f"Последовательность без группировки: нет целой для начала дорожки "
                            f"(элементов={len(items)}, целых={_count_solids_remaining(items)}, "
                            f"mode={item.get('mode')!r}). Измените план или заказ."
                        )
                    else:
                        msg = (
                            "Последовательность без группировки: нет целой при "
                            "LAYOUT_TRACK_START_REINF_RELAXATION=false."
                        )
                    raise TrackLayoutInvariantError(msg)
                if items[found_solid_idx].get("is_separator", False):
                    logger.warning(
                        "[SPLIT_TRACKS] Используем разделитель для начала дорожки (не идеально)"
                    )
                solid_plate = items.pop(found_solid_idx)
                is_sep = solid_plate.get('is_separator', False)
                logger.info(
                    f"[SPLIT_TRACKS] Найдена целая плита для НАЧАЛА дорожки: {solid_plate['length']:.2f}м (разделитель={is_sep})"
                )
                current_track.append(solid_plate)
                current_track_length += solid_plate['length']
                solid_reinf = solid_plate.get('reinforcement', 0) or 0
                max_reinforcement_in_track = max(max_reinforcement_in_track, solid_reinf)
                if found_solid_idx < i:
                    i -= 1
                continue
            
            # Проверяем превышение максимума
            will_exceed_max = (current_track_length + item_length > max_track_length and current_track)
            
            if will_exceed_max:
                remaining_space = max_track_length - current_track_length
                
                if is_solid and item_length <= remaining_space:
                    current_track.append(item)
                    current_track_length += item_length
                    max_reinforcement_in_track = max(max_reinforcement_in_track, item_reinforcement)
                    i += 1
                elif not is_solid:
                    neighbor_reinf = float(item_reinforcement or 0)
                    eligible_ns = []
                    eligible_sep = []
                    for j in range(i + 1, len(items)):
                        candidate = items[j]
                        if candidate.get("mode") != "solid":
                            continue
                        cand_length = candidate["length"]
                        cand_reinf = candidate.get("reinforcement", 0) or 0
                        if cand_length > remaining_space:
                            continue
                        if max_reinforcement_in_track > 0 and cand_reinf > max_reinforcement_in_track:
                            continue
                        if not candidate.get("is_separator", False):
                            eligible_ns.append(j)
                        else:
                            eligible_sep.append(j)
                    found_solid_idx = None
                    if eligible_ns:
                        if _pref_reinf:
                            found_solid_idx = min(
                                eligible_ns,
                                key=lambda j, _nr=neighbor_reinf: (
                                    abs(float(items[j].get("reinforcement", 0) or 0) - _nr),
                                    j,
                                ),
                            )
                        else:
                            found_solid_idx = min(eligible_ns)
                    elif eligible_sep:
                        if _pref_reinf:
                            found_solid_idx = min(
                                eligible_sep,
                                key=lambda j, _nr=neighbor_reinf: (
                                    abs(float(items[j].get("reinforcement", 0) or 0) - _nr),
                                    j,
                                ),
                            )
                        else:
                            found_solid_idx = min(eligible_sep)
                        logger.warning(
                            "[SPLIT_TRACKS] Используем разделитель для завершения дорожки (не идеально)"
                        )

                    if found_solid_idx is not None:
                        candidate = items[found_solid_idx]
                        is_sep = candidate.get('is_separator', False)
                        logger.info(
                            f"[SPLIT_TRACKS] Найдена целая плита для завершения: {candidate['length']:.2f}м, "
                            f"арм.={candidate.get('reinforcement', 0):.1f} (разделитель={is_sep})"
                        )
                        current_track.append(candidate)
                        current_track_length += candidate['length']
                        items.pop(found_solid_idx)
                        if found_solid_idx < i:
                            i -= 1
                    else:
                        logger.warning("[SPLIT_TRACKS] Целой плиты для завершения не найдено, закрываем как есть")
                else:
                    pass
                
                # Закрываем дорожку
                logger.info(
                    f"[SPLIT_TRACKS] Закрываем дорожку на {current_track_length:.1f}м (макс. {max_track_length}м)"
                )
                tracks.append({
                    'items': current_track,
                    'length': current_track_length,
                    'max_reinforcement': max_reinforcement_in_track
                })
                current_track = []
                current_track_length = 0.0
                max_reinforcement_in_track = 0.0
                
                if is_solid and item_length <= remaining_space:
                    continue
                continue
            
            # Плита влезает - добавляем
            current_track.append(item)
            current_track_length += item_length
            max_reinforcement_in_track = max(max_reinforcement_in_track, item_reinforcement)
            i += 1
        
        if current_track:
            tracks.append({
                'items': current_track,
                'length': current_track_length,
                'max_reinforcement': max_reinforcement_in_track
            })
    
    output_count = sum(len(t['items']) for t in tracks)
    if input_count != output_count:
        logger.warning(
            "[SPLIT_TRACKS] Потеря плит при разбиении: было %s, в дорожках %s, разница %s",
            input_count, output_count, input_count - output_count
        )

    logger.info(f"[SPLIT_TRACKS] Плиты разбиты на {len(tracks)} дорожек")
    try:
        integrity_report = validate_track_integrity(
            sequence,
            tracks,
            strict=strict_layout_integrity,
        )
        if not integrity_report["ok"]:
            logger.error(
                "[SPLIT_TRACKS] Integrity mismatch: input=%s output=%s missing=%s duplicated=%s",
                integrity_report["input_total"],
                integrity_report["output_total"],
                len(integrity_report["missing"]),
                len(integrity_report["duplicated"]),
            )
    except LayoutIntegrityError:
        raise
    except Exception:
        logger.exception("[SPLIT_TRACKS] Ошибка проверки integrity split")
    if require_solid_track_start:
        _assert_track_starts_with_solid(tracks)
    return tracks


