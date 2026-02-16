#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль визуализации и работы с ценами:
- Загрузка прайса из XLSX
- Работа с базой цен SQLite
- Построение сметы
- Визуализация раскладки плит
"""
import os
import logging
from pathlib import Path
from collections import Counter
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.lines import Line2D

# Относительные импорты внутри core/
from . import config_and_data as cfg
from .optimization import optimize_cuts_pulp
from .price_db import init_schema, import_from_xlsx
from .exceptions import FileGenerationError

# Настройка логирования
logger = logging.getLogger(__name__)

# Импорты из новых модулей
from viz_modules.price_utils import load_price_table_from_xlsx
from viz_modules.procurement import (
    build_procurement_items,
    build_price_rows,
    build_component_breakdown,
    get_orders_from_opt_plan
)
from viz_modules.layout_sequence import build_layout_sequence
from viz_modules.visualization_drawing import (
    _draw_segment,
    _draw_split_plate,
    _draw_transverse_cut
)

try:
    import pandas as pd
except Exception:
    pd = None


# Реэкспорт для обратной совместимости
__all__ = ['visualize_plan', 'build_layout_sequence', 'split_sequence_into_tracks']


def split_sequence_into_tracks(
    sequence: list,
    max_track_length: float = 101.0,
    min_track_length: float = 96.0
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
        min_track_length: Минимальная длина для закрытия дорожки (96м)
    
    Returns:
        list[dict]: Список дорожек [{'items': [...], 'length': float, 'load_code': int, 'label': str, 'max_reinforcement': float}, ...]
    """
    def _label_from_track_items(items, fallback):
        """Подпись дорожки по фактическим нагрузкам плит в ней (чтобы при группе 'all' не было везде 8п)."""
        load_codes = set()
        for item in items:
            lc = item.get('load_code')
            if lc is not None:
                try:
                    n = cfg.normalize_load_code(lc)
                    if n is not None:
                        load_codes.add(n)
                except Exception:
                    pass
        if not load_codes:
            return fallback
        disp = [cfg.format_reinforcement_from_load_code(lc) for lc in sorted(load_codes)]
        return 'Нагрузка ' + ', '.join(disp) if len(disp) > 1 else f'Нагрузка {disp[0]}'

    tracks = []
    
    # Проверяем формат данных (с группировкой по нагрузке или без)
    if isinstance(sequence, list) and sequence and isinstance(sequence[0], dict) and 'load_code' in sequence[0]:
        logger.info(f"[SPLIT_TRACKS] Обнаружена группировка по нагрузкам. Групп: {len(sequence)}")
        
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
                    found_solid_idx = None
                    fallback_separator_idx = None  # Разделитель как запасной вариант
                    
                    # Приоритет 1: НЕ-разделители ПОСЛЕ текущей позиции
                    for j in range(i + 1, len(items)):
                        candidate = items[j]
                        if candidate.get('mode') == 'solid' and not candidate.get('is_separator', False):
                            found_solid_idx = j
                            logger.info(f"[SPLIT_TRACKS] Найдена НЕ-разделитель ПОСЛЕ: индекс {j}")
                            break
                    
                    # Приоритет 2: НЕ-разделители ДО текущей позиции (если не нашли выше)
                    if found_solid_idx is None:
                        for j in range(i):
                            candidate = items[j]
                            if candidate.get('mode') == 'solid' and not candidate.get('is_separator', False):
                                found_solid_idx = j
                                logger.info(f"[SPLIT_TRACKS] Найдена НЕ-разделитель ДО: индекс {j}")
                                break
                    
                    # Приоритет 3: Разделители ПОСЛЕ (если НЕ-разделителей нет)
                    if found_solid_idx is None:
                        for j in range(i + 1, len(items)):
                            candidate = items[j]
                            if candidate.get('mode') == 'solid' and candidate.get('is_separator', False):
                                fallback_separator_idx = j
                                logger.info(f"[SPLIT_TRACKS] Найден разделитель ПОСЛЕ: индекс {j}")
                                break
                    
                    # Приоритет 4: Разделители ДО (последний резерв)
                    if found_solid_idx is None and fallback_separator_idx is None:
                        for j in range(i):
                            candidate = items[j]
                            if candidate.get('mode') == 'solid' and candidate.get('is_separator', False):
                                fallback_separator_idx = j
                                logger.info(f"[SPLIT_TRACKS] Найден разделитель ДО: индекс {j}")
                                break
                    
                    # Если не нашли НЕ-разделитель, берём разделитель
                    if found_solid_idx is None and fallback_separator_idx is not None:
                        found_solid_idx = fallback_separator_idx
                        logger.warning("[SPLIT_TRACKS] Используем разделитель для начала дорожки (не идеально)")
                    
                    if found_solid_idx is not None:
                        # Нашли целую плиту - добавляем её ПЕРВОЙ!
                        solid_plate = items.pop(found_solid_idx)
                        is_sep = solid_plate.get('is_separator', False)
                        logger.info(
                            f"[SPLIT_TRACKS] Найдена целая плита для НАЧАЛА дорожки: {solid_plate['length']:.2f}м (разделитель={is_sep})"
                        )
                        current_track.append(solid_plate)
                        current_track_length += solid_plate['length']
                        solid_reinf = solid_plate.get('reinforcement', 0) or 0
                        max_reinforcement_in_track = max(max_reinforcement_in_track, solid_reinf)
                        # НЕ увеличиваем i - текущая плита с резом будет добавлена следующей итерацией
                        continue
                    else:
                        logger.warning("[SPLIT_TRACKS] ВНИМАНИЕ: целой плиты для начала дорожки не найдено!")
                        # #region agent log
                        try:
                            open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", 'a', encoding='utf-8').write(__import__('json').dumps({"hypothesisId": "H4", "location": "visualization.py:split_tracks_no_solid", "message": "no solid plate for track start", "data": {"group_label": group_label, "i": i, "item_length": item_length, "current_track_len": len(current_track)}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
                        except Exception:
                            pass
                        # #endregion
                
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
                        found_solid_idx = None
                        fallback_separator_idx = None
                        
                        for j in range(i + 1, len(items)):
                            candidate = items[j]
                            if candidate.get('mode') == 'solid':
                                cand_reinf = candidate.get('reinforcement', 0) or 0
                                if max_reinforcement_in_track == 0 or cand_reinf <= max_reinforcement_in_track:
                                    if not candidate.get('is_separator', False):
                                        found_solid_idx = j
                                        break
                                    elif fallback_separator_idx is None:
                                        fallback_separator_idx = j
                        
                        if found_solid_idx is None and fallback_separator_idx is not None:
                            found_solid_idx = fallback_separator_idx
                            logger.warning("[SPLIT_TRACKS] Используем разделитель для начала следующей дорожки (не идеально)")
                        
                        if found_solid_idx is not None:
                            candidate = items.pop(found_solid_idx)
                            # Закрываем текущую дорожку БЕЗ добавления candidate (он пойдёт в новую).
                            # После закрытия: новая дорожка = candidate + split_plate.
                            logger.info(
                                f"[SPLIT_TRACKS] Плита с резом {split_plate['length']:.2f}м переносится в следующую дорожку, "
                                f"первой — целая {candidate['length']:.2f}м"
                            )
                            # Удаляем split_plate из items (индекс сдвинулся, если found_solid_idx < i)
                            split_idx = i if found_solid_idx > i else i - 1
                            items.pop(split_idx)
                            if found_solid_idx < i:
                                i -= 2  # убрали два элемента до текущего; следующий к обработке теперь на i-2
                            # Закрываем дорожку ниже; затем current_track = [candidate, split_plate]
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
                            logger.warning("[SPLIT_TRACKS] Целой плиты для начала следующей дорожки не найдено, закрываем как есть и добавляем плиту с резом в новую")
                            # Закрываем как есть; новую дорожку начинаем с плиты с резом (fallback).
                            track_label = _label_from_track_items(current_track, group_label)
                            tracks.append({
                                'items': current_track,
                                'length': current_track_length,
                                'load_code': load_code,
                                'label': track_label,
                                'max_reinforcement': max_reinforcement_in_track
                            })
                            current_track = [split_plate]
                            current_track_length = split_plate['length']
                            max_reinforcement_in_track = max(max_reinforcement_in_track, split_plate.get('reinforcement', 0) or 0)
                            items.pop(i)
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
                found_solid_idx = None
                fallback_separator_idx = None
                
                # РАСШИРЕННЫЙ ПОИСК с 4 приоритетами
                for j in range(i + 1, len(items)):
                    candidate = items[j]
                    if candidate.get('mode') == 'solid' and not candidate.get('is_separator', False):
                        found_solid_idx = j
                        logger.info(f"[SPLIT_TRACKS] Найдена НЕ-разделитель ПОСЛЕ: индекс {j}")
                        break
                
                if found_solid_idx is None:
                    for j in range(i):
                        candidate = items[j]
                        if candidate.get('mode') == 'solid' and not candidate.get('is_separator', False):
                            found_solid_idx = j
                            logger.info(f"[SPLIT_TRACKS] Найдена НЕ-разделитель ДО: индекс {j}")
                            break
                
                if found_solid_idx is None:
                    for j in range(i + 1, len(items)):
                        candidate = items[j]
                        if candidate.get('mode') == 'solid' and candidate.get('is_separator', False):
                            fallback_separator_idx = j
                            logger.info(f"[SPLIT_TRACKS] Найден разделитель ПОСЛЕ: индекс {j}")
                            break
                
                if found_solid_idx is None and fallback_separator_idx is None:
                    for j in range(i):
                        candidate = items[j]
                        if candidate.get('mode') == 'solid' and candidate.get('is_separator', False):
                            fallback_separator_idx = j
                            logger.info(f"[SPLIT_TRACKS] Найден разделитель ДО: индекс {j}")
                            break
                
                if found_solid_idx is None and fallback_separator_idx is not None:
                    found_solid_idx = fallback_separator_idx
                    logger.warning("[SPLIT_TRACKS] Используем разделитель для начала дорожки (не идеально)")
                
                if found_solid_idx is not None:
                    solid_plate = items.pop(found_solid_idx)
                    is_sep = solid_plate.get('is_separator', False)
                    logger.info(
                        f"[SPLIT_TRACKS] Найдена целая плита для НАЧАЛА дорожки: {solid_plate['length']:.2f}м (разделитель={is_sep})"
                    )
                    current_track.append(solid_plate)
                    current_track_length += solid_plate['length']
                    solid_reinf = solid_plate.get('reinforcement', 0) or 0
                    max_reinforcement_in_track = max(max_reinforcement_in_track, solid_reinf)
                    continue
                else:
                    logger.warning("[SPLIT_TRACKS] ВНИМАНИЕ: целой плиты для начала дорожки не найдено!")
            
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
                    found_solid_idx = None
                    fallback_separator_idx = None
                    
                    for j in range(i + 1, len(items)):
                        candidate = items[j]
                        if candidate.get('mode') == 'solid':
                            cand_length = candidate['length']
                            cand_reinf = candidate.get('reinforcement', 0) or 0
                            
                            if cand_length <= remaining_space:
                                if max_reinforcement_in_track == 0 or cand_reinf <= max_reinforcement_in_track:
                                    if not candidate.get('is_separator', False):
                                        found_solid_idx = j
                                        break
                                    elif fallback_separator_idx is None:
                                        fallback_separator_idx = j
                    
                    if found_solid_idx is None and fallback_separator_idx is not None:
                        found_solid_idx = fallback_separator_idx
                        logger.warning("[SPLIT_TRACKS] Используем разделитель для завершения дорожки (не идеально)")
                    
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
    
    logger.info(f"[SPLIT_TRACKS] Плиты разбиты на {len(tracks)} дорожек")
    return tracks


def visualize_plan(output_dir: str = 'Визуализация_Раскладки', 
                    tracks_per_file: int = None, 
                    start_track_index: int = 0,
                    use_production_pricing: bool = False,
                    auto_import_price_to_db: bool = True,
                    existing_tracks: list = None):
    """
    Создаёт визуализацию раскладки плит и сохраняет файлы
    
    Args:
        output_dir: Директория для сохранения файлов
        tracks_per_file: Сколько дорожек поместить в один файл (None = все дорожки)
        start_track_index: С какой дорожки начинать (0 = с первой)
        use_production_pricing: Если True, использует расчет для планирования производства 
                                (базовая цена из raw_material_costs + переармирование)
        existing_tracks: Готовые дорожки из сохранённого плана (если переданы, 
                         build_layout_sequence НЕ вызывается - используем готовые данные)
    
    Raises:
        FileGenerationError: Если не удалось создать файлы или загрузить прайс
    """
    # Константы длины дорожки (определяем в начале, чтобы были доступны везде)
    MAX_TRACK_LENGTH = 101.0  # Максимальная длина дорожки (ЖЁСТКИЙ ЛИМИТ!)
    MIN_TRACK_LENGTH = 96.0   # Минимальная длина дорожки для закрытия
    
    logger.info(f"Начало генерации визуализации. Директория: {output_dir}")
    
    # Оптимизация резов (не критично, если не сработает)
    try:
        optimized = optimize_cuts_pulp({300: 4, 500: 3, 700: 2, 900: 2})
        logger.info(f"Оптимальные резы рассчитаны: {optimized}")
    except Exception as e:
        logger.warning(f"Ошибка при оптимизации резов (не критично): {e}")

    # Создаём директорию для результатов
    try:
        os.makedirs(output_dir, exist_ok=True)
        logger.debug(f"Директория создана/проверена: {output_dir}")
    except OSError as e:
        logger.error(f"Не удалось создать директорию {output_dir}: {e}")
        raise FileGenerationError(f"Не удалось создать папку для файлов: {e}")

    # Проверяем существование файла прайса (мягкая проверка)
    if not Path(cfg.PRICE_XLSX_PATH).exists():
        logger.warning(f"Файл прайса не найден по точному пути: {cfg.PRICE_XLSX_PATH}")
        logger.info("Попытка автоматического поиска файла прайса в папке 'банк знаний'...")
    
    # Загружаем прайс из Excel (функция имеет встроенный автопоиск)
    # Если не удалось загрузить из Excel - используем цены из БД
    try:
        price_table = load_price_table_from_xlsx(str(cfg.PRICE_XLSX_PATH))
        if not price_table:
            logger.warning("Прайс-лист из Excel пуст, будут использованы цены из БД")
            price_table = {}  # Пустой словарь - цены будут из БД
        else:
            logger.info(f"Прайс-лист успешно загружен из Excel ({len(price_table)} позиций)")
    except Exception as e:
        logger.warning(f"Не удалось загрузить прайс из Excel: {e}")
        logger.info("Будут использованы цены из базы данных")
        price_table = {}  # Пустой словарь - цены будут из БД
    
    # Импорт прайса в БД (не критично, если не сработает)
    if auto_import_price_to_db:
        try:
            init_schema(str(cfg.PRICE_DB_PATH))

            # Не импортируем каждый раз: импортируем только если БД отсутствует
            # или Excel новее базы (иначе это лишняя тяжёлая операция).
            xlsx_path = Path(cfg.PRICE_XLSX_PATH)
            db_path = Path(cfg.PRICE_DB_PATH)
            should_import = (not db_path.exists())
            if (not should_import) and xlsx_path.exists():
                try:
                    should_import = xlsx_path.stat().st_mtime > db_path.stat().st_mtime
                except Exception:
                    should_import = True

            if should_import:
                written = import_from_xlsx(str(cfg.PRICE_XLSX_PATH), str(cfg.PRICE_DB_PATH))
                if written:
                    logger.info(f'Прайс импортирован в БД: {written} строк')
            else:
                logger.debug("Импорт прайса в БД пропущен: база уже актуальна")
        except Exception as e:
            logger.warning(f"Не удалось импортировать прайс в БД (не критично): {e}")

    # Выбираем функции в зависимости от режима
    if use_production_pricing:
        from viz_modules.procurement import build_price_rows_production, build_component_breakdown_production
        price_rows, total_sum = build_price_rows_production(price_table)
        breakdown_tables = build_component_breakdown_production(price_table, price_rows)
        logger.info("Используется расчет для планирования производства (raw_material_costs + переармирование)")
    else:
        price_rows, total_sum = build_price_rows(price_table)
        breakdown_tables = build_component_breakdown(price_table, price_rows)
        logger.info("Используется расчет для коммерческого предложения")
    
    # Отладочная информация
    logger.debug(f'breakdown_tables count: {len(breakdown_tables) if breakdown_tables else 0}')
    if breakdown_tables:
        for i, bt in enumerate(breakdown_tables):
            logger.debug(f'breakdown_tables[{i}]: name={bt.get("name")}, rows={len(bt.get("rows", []))}')
    
    # ✅ НОВОЕ: Если переданы готовые дорожки - используем их напрямую!
    if existing_tracks:
        logger.info(f"[ВИЗУАЛИЗАЦИЯ] Используем готовые дорожки из плана: {len(existing_tracks)} дорожек")
        tracks = existing_tracks
        total_length = sum(t.get('length', 0) for t in tracks)
    
    # Стандартная логика: генерируем последовательность и разбиваем на дорожки (только если нет готовых дорожек)
    if not existing_tracks:
        seq = build_layout_sequence()
        tracks = split_sequence_into_tracks(seq, MAX_TRACK_LENGTH, MIN_TRACK_LENGTH)
        total_length = sum(t['length'] for t in tracks)
    
    num_tracks_total = len(tracks)
    logger.info(f"[ВИЗУАЛИЗАЦИЯ] Плиты разбиты на {num_tracks_total} дорожек")
    
    # ✅ НОВОЕ: Логируем плиты в дорожках
    logger.info(f"[TRACE] ===== ШАГ 6: ПЛИТЫ В ДОРОЖКАХ (tracks) =====")
    logger.info(f"[TRACE] Всего дорожек: {len(tracks)}")
    
    total_plates_in_tracks = 0
    for i, track in enumerate(tracks):
        track_plates = len(track['items'])
        total_plates_in_tracks += track_plates
        
        # Подсчитываем плиты с вторичными резами
        secondary_in_track = 0
        for item in track['items']:
            if item.get('secondary_cuts'):
                secondary_in_track += len(item['secondary_cuts'])
        
        logger.info(f"[TRACE]   Дорожка #{i+1}: {track_plates} основных плит + {secondary_in_track} вторичных = {track_plates + secondary_in_track} всего")
    
    logger.info(f"[TRACE] ИТОГО плит во всех дорожках: {total_plates_in_tracks}")
    
    # ✅ НОВОЕ: Фильтруем дорожки для текущего файла
    if tracks_per_file is not None:
        end_track_index = min(start_track_index + tracks_per_file, num_tracks_total)
        tracks = tracks[start_track_index:end_track_index]
        actual_start = start_track_index + 1
        actual_end = start_track_index + len(tracks)
        logger.info(
            f"[ВИЗУАЛИЗАЦИЯ] Файл содержит дорожки {actual_start}-{actual_end} (всего {len(tracks)} дорожек в файле)"
        )
    
    num_tracks = len(tracks)

    # === РАСЧЁТ МАКСИМАЛЬНОГО АРМИРОВАНИЯ ДЛЯ КАЖДОЙ ДОРОЖКИ ===
    for track in tracks:
        max_reinforcement = 0.0
        for item in track['items']:
            reinforcement = item.get('reinforcement', 0)
            # Исключаем fallback значения (999.0)
            if reinforcement and reinforcement < 999:
                max_reinforcement = max(max_reinforcement, reinforcement)
        track['max_reinforcement'] = max_reinforcement
        if max_reinforcement > 0:
            logger.info(f"[ВИЗУАЛИЗАЦИЯ] Дорожка: макс. армирование {max_reinforcement:.1f} кг/м²")
    
    # === ЗАПОЛНЯЕМ ГЛОБАЛЬНУЮ КАРТУ МАКСИМАЛЬНОГО АРМИРОВАНИЯ ДЛЯ КАЖДОЙ ПЛИТЫ ===
    # Это нужно для корректного расчёта переармирования в procurement.py
    for track in tracks:
        track_max_reinf = track.get('max_reinforcement', 0)
        for item in track['items']:
            length = item.get('length', 0)
            # Определяем ширину плиты
            if item.get('mode') == 'solid':
                width_mm = 1200
            elif item.get('mode') == 'split':
                width_mm = int(round(item.get('main_w', 1.2) * 1000))
            elif item.get('mode') == 'transverse':
                width_mm = int(round(item.get('width', 1.2) * 1000))
            else:
                width_mm = 1200
            
            # Сохраняем максимальное армирование дорожки для этой плиты
            key = (round(length, 3), width_mm)
            # Если плита уже есть в карте, берём максимум (она может быть в нескольких дорожках)
            if key in cfg.PLATE_MAX_REINFORCEMENT_MAP:
                cfg.PLATE_MAX_REINFORCEMENT_MAP[key] = max(cfg.PLATE_MAX_REINFORCEMENT_MAP[key], track_max_reinf)
            else:
                cfg.PLATE_MAX_REINFORCEMENT_MAP[key] = track_max_reinf
    
    logger.info(f"[ВИЗУАЛИЗАЦИЯ] Заполнена карта макс. армирования: {len(cfg.PLATE_MAX_REINFORCEMENT_MAP)} плит")
    
    # ✅ ПЕРЕСЧИТЫВАЕМ breakdown_tables после заполнения PLATE_MAX_REINFORCEMENT_MAP
    # Это нужно для корректного расчёта переармирования по дорожкам
    if use_production_pricing and cfg.PLATE_MAX_REINFORCEMENT_MAP:
        from viz_modules.procurement import build_component_breakdown_production
        breakdown_tables = build_component_breakdown_production(price_table, price_rows)
        logger.info("[ВИЗУАЛИЗАЦИЯ] Пересчитана детальная разбивка с учётом максимального армирования по дорожкам")

    # Убрали секцию детальной разбивки - она теперь в отдельном Excel файле
    # Увеличиваем высоту секции дорожек пропорционально их количеству
    track_section_height = 3.0 + (num_tracks - 1) * 2.5
    # ✅ УБРАНЫ таблицы с ценами и заказами - теперь только 2 секции
    num_sections = 2
    height_ratios = [track_section_height, 1.0]
    
    # Уменьшаем общую высоту окна (убрали таблицы)
    total_fig_height = 8 + (num_tracks - 1) * 5
    
    fig = plt.figure(figsize=(22, total_fig_height))
    gs = fig.add_gridspec(num_sections, 1, height_ratios=height_ratios)
    ax_track = fig.add_subplot(gs[0, 0])
    ax_strips = fig.add_subplot(gs[1, 0])
    # ✅ УБРАНЫ секции ax_table (таблица заказа) и ax_price (таблица с ценами)
    
    # ✅ НОВОЕ: Заголовок с правильными номерами дорожек (БЕЗ упоминания сметы)
    if num_tracks == 1:
        track_num = start_track_index + 1
        fig.suptitle(f'КЗ: Дорожка {track_num} (ширина 1.2 м) — раскладка, резы и ведомости', 
                     fontsize=16, fontweight='bold')
    else:
        first_track = start_track_index + 1
        last_track = start_track_index + num_tracks
        fig.suptitle(f'КЗ: Дорожки {first_track}-{last_track} (ширина 1.2 м, по {MAX_TRACK_LENGTH}м) — раскладка, резы и ведомости', 
                     fontsize=16, fontweight='bold')

    # Настройка осей для множественных дорожек
    track_height = cfg.TRACK_WIDTH_M  # 1.2 м
    track_spacing = 0.3  # Отступ между дорожками
    total_height = num_tracks * (track_height + track_spacing)
    
    # Рассчитываем максимальную длину дорожек для правильного xlim
    # (некоторые дорожки могут немного превышать 101м из-за правила завода)
    max_actual_length = max((t.get('length', 0) for t in tracks), default=MAX_TRACK_LENGTH)
    display_max_length = max(MAX_TRACK_LENGTH, max_actual_length + 1.0)
    
    ax_track.set_xlim(0, display_max_length)
    ax_track.set_ylim(0, total_height)
    ax_track.set_aspect('auto')
    ax_track.spines['top'].set_visible(False)
    ax_track.spines['right'].set_visible(False)
    
    # ✅ НОВОЕ: Метки по оси Y с правильными номерами дорожек
    y_ticks = []
    y_labels = []
    for i in range(num_tracks):
        y_pos = i * (track_height + track_spacing) + track_height / 2
        y_ticks.append(y_pos)
        
        # Номер дорожки с учётом start_track_index
        actual_track_num = start_track_index + i + 1
        track_label = f"Д{actual_track_num}"
        
        # НОВОЕ: Добавляем информацию о плане-источнике
        track_data = tracks[i]
        if 'source_plan_name' in track_data:
            plan_name = track_data['source_plan_name']
            # Обрезаем длинные названия для читаемости
            if len(plan_name) > 20:
                plan_name = plan_name[:17] + '...'
            track_label += f"\n({plan_name})"
        
        if 'load_code' in tracks[i]:
            load_label = tracks[i].get('label', f"Нагрузка {tracks[i]['load_code']}п")
            track_label += f"\n{load_label}"
        
        # ✅ НОВОЕ: Добавляем максимальное армирование
        max_reinf = tracks[i].get('max_reinforcement', 0)
        if max_reinf > 0:
            track_label += f"\nмакс. арм. {max_reinf:.1f}"
        
        if track_label == f"Д{actual_track_num}":
            # Если нет доп. информации, используем полное название
            y_labels.append(f'Дор.{actual_track_num}')
        else:
            y_labels.append(track_label)
    
    ax_track.set_yticks(y_ticks)
    ax_track.set_yticklabels(y_labels)
    
    ax_track.set_xlabel('Длина (м)')
    ax_track.set_xticks(range(0, int(MAX_TRACK_LENGTH) + 1, 5))
    ax_track.grid(axis='x', linestyle=':', linewidth=0.5, alpha=0.5)

    # Рисуем каждую дорожку
    for track_idx, track_data in enumerate(tracks):
        # Y-координата текущей дорожки
        y_base = track_idx * (track_height + track_spacing)
        
        # Рамка дорожки
        track_rect = patches.Rectangle(
            (0, y_base), 
            MAX_TRACK_LENGTH, 
            track_height, 
            linewidth=2, 
            edgecolor='black', 
            facecolor='none', 
            linestyle='--'
        )
        ax_track.add_patch(track_rect)
        
        # Рисуем плиты в этой дорожке
        x = 0.0
        for item in track_data['items']:
            if item.get('mode') == 'solid':
                _draw_segment(ax_track, x, item['length'], '#2ecc71', item['label'], 
                            y=y_base, height=track_height,
                            reinforcement=item.get('reinforcement'))
            elif item.get('mode') == 'transverse':
                # Плита с поперечным резом (по длине)
                _draw_transverse_cut(
                    ax_track, x, 
                    total_length=item['length'],
                    target_length=item['target_length'],
                    width=item['width'],
                    label_target=item['label_target'],
                    remainder_length=item['remainder'],
                    y_base=y_base,
                    reinforcement=item.get('reinforcement')
                )
            else:
                # Плиты с резами (первичными и возможными вторичными)
                _draw_split_plate(
                    ax_track, x, item['length'],
                    main_w=item['main_w'], rest_w=item['rest_w'],
                    label_main=item['label_main'], label_rest=item.get('label_rest'),
                    secondary_cuts=item.get('secondary_cuts'),
                    y_base=y_base,
                    reinforcement=item.get('reinforcement')
                )
            x += item['length']

    legend_patches = [
        patches.Patch(facecolor='#2ecc71', edgecolor='black', label='🟢 Основа (первичный рез)'),
        patches.Patch(facecolor='#3498db', edgecolor='black', label='🔵 Вторичный рез (из остатка)'),
        patches.Patch(facecolor='#95a5a6', edgecolor='gray', label='⬛ Отход'),
        patches.Patch(facecolor='#ecf0f1', edgecolor='black', label='⬜ Остаток (не использован)'),
        Line2D([0], [0], color='blue', linestyle='-', linewidth=2.5, label='━ Продольный рез (первичный)'),
        Line2D([0], [0], color='orange', linestyle='-', linewidth=2.0, label='━ Продольный рез (вторичный)'),
        Line2D([0], [0], color='red', linestyle='--', linewidth=2.5, label='┊ Поперечный рез (по длине)'),
    ]
    ax_track.legend(handles=legend_patches, loc='upper right', fontsize=9)

    ax_strips.set_xlim(0, 100)
    ax_strips.set_ylim(0, 1)
    ax_strips.axis('off')

    # Формируем сводку с учётом каскадной оптимизации
    from .optimization import OPT_CASCADING_PLAN
    txt = (
        f"Длина по плану: {total_length:.1f} м ({num_tracks} дорожек)  |  Продольных резов: {cfg.LONGITUDINAL_CUTS}  |  Подрезов по длине: {cfg.LENGTH_TRIMS}\n"
        f"Остатки лент 0.3: {cfg.UNUSED_STRIPS_0_3_M_TOTAL:.1f} пог.м  |  Обрезки 0.2: {cfg.SCRAP_STRIPS_0_2_M_TOTAL:.1f} пог.м (≈ {cfg.WASTE_AREA_M2:.2f} м²)"
    )
    
    # Добавляем информацию о каскадной оптимизации, если она была использована
    if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('total_plates', 0) > 0:
        txt += f"\n\nОПТИМИЗАЦИЯ: Плит потребуется {OPT_CASCADING_PLAN['total_plates']} шт (с каскадными резами)"
        txt += f" | Отходы: {OPT_CASCADING_PLAN.get('waste_width', 0)} мм"
    
    # Размещаем информацию о резах слева (не мешает дорожкам, т.к. выше)
    ax_strips.text(0.02, 0.6, txt, ha='left', va='center', fontsize=11,
                   bbox=dict(boxstyle='round,pad=0.5', facecolor='#f8f9fa', edgecolor='#bdc3c7'))

    # Детальный план резов убран из визуализации (по запросу пользователя)
    # Все данные о резах сохраняются в Excel файлах

    # ✅ УБРАНА таблица заказа/использования (ax_table)
    # Но данные всё равно формируем для CSV и Excel файлов
    
    # Формируем список заказа из реальных данных
    order_list = []
    
    plan_orders = get_orders_from_opt_plan()
    if plan_orders:
        plan_counter = Counter()
        for order in plan_orders:
            key = (round(float(order['length']), 3), int(order['width']))
            plan_counter[key] += order['qty']
        for (length, width_mm), qty in sorted(plan_counter.items(), key=lambda x: (x[0][0], x[0][1])):
            order_list.append(f"Заказ {length:.2f}м×{width_mm}мм: {qty} шт")
    else:
        # Собираем все плиты из заказа (legacy путь через cfg)
        all_orders = []
        for width_mm, plates_list in [
            (320, cfg.PLATES_0_32), (460, cfg.PLATES_0_46), (700, cfg.PLATES_0_70),
            (720, cfg.PLATES_0_72), (860, cfg.PLATES_0_86), (880, cfg.PLATES_0_88),
            (740, cfg.PLATES_0_74), (480, cfg.PLATES_0_48), (500, cfg.PLATES_0_50),
            (340, cfg.PLATES_0_34), (1080, cfg.PLATES_1_08)
        ]:
            if plates_list:
                length_counts = Counter(plates_list)
                for length, qty in sorted(length_counts.items(), key=lambda x: (-x[0], -x[1])):
                    all_orders.append({
                        'length': length,
                        'width': width_mm,
                        'qty': qty
                    })
        if all_orders:
            for order in all_orders:
                order_list.append(f"Заказ {order['length']:.1f}м×{order['width']}мм: {order['qty']} шт")
        else:
            order_list.append('Заказ не найден')
    
    # Формируем список использования из оптимизации
    used_list = []
    
    if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('total_plates', 0) > 0:
        # Итого плит
        used_list.append(f"Плит 1200мм потребуется: {OPT_CASCADING_PLAN['total_plates']} шт")
        
        # Первичные резы
        if OPT_CASCADING_PLAN.get('primary_cuts'):
            primary_info = []
            for cut in OPT_CASCADING_PLAN['primary_cuts']:
                primary_info.append(f"{cut['qty']}x({cut['width']}мм+{cut['rest']}мм)")
            if primary_info:
                used_list.append(f"Первичные резы: {'; '.join(primary_info)}")
        
        # Вторичные резы
        if OPT_CASCADING_PLAN.get('secondary_cuts'):
            secondary_info = []
            for cut in OPT_CASCADING_PLAN['secondary_cuts']:
                if cut.get('pieces', 1) > 1:
                    secondary_info.append(f"{cut['qty']}x{cut['source']}мм->{cut['pieces']}x{cut['cuts'][0]}мм")
                else:
                    secondary_info.append(f"{cut['qty']}x{cut['source']}мм->{cut['cuts'][0]}мм")
            if secondary_info:
                used_list.append(f"Вторичные резы: {'; '.join(secondary_info)}")
        
        # Поперечные резы
        if OPT_CASCADING_PLAN.get('transverse_cuts'):
            trans_count = len(OPT_CASCADING_PLAN['transverse_cuts'])
            used_list.append(f"Поперечных резов: {trans_count}")
        
        # Отходы
        if OPT_CASCADING_PLAN.get('waste_width', 0) > 0:
            used_list.append(f"Отходы: {OPT_CASCADING_PLAN['waste_width']} мм по ширине")
    else:
        # Старый формат, если оптимизация не использовалась
        used_list.append('1.2 без реза: 6.3x2; 3.8x2')
        used_list.append('1.5->1.2: 3.8x3; 2.9x1 (остаток 0.3)')
        used_list.append(f'Резы: продольных {cfg.LONGITUDINAL_CUTS}; подрезов {cfg.LENGTH_TRIMS}')

    rows = max(len(order_list), len(used_list))
    table_rows = []
    for i in range(rows):
        left = order_list[i] if i < len(order_list) else ''
        right = used_list[i] if i < len(used_list) else ''
        table_rows.append([left, right])

    col_labels = ['Список плит по заказу', 'Использовано (с учётом резов) / остатки / обрезки']

    # ✅ ПОЛНОСТЬЮ УБРАНЫ таблицы: ax_table (заказ/использование) и ax_price (цены)

    # Детальная разбивка теперь сохраняется в отдельный Excel файл, а не отображается на графике

    # ✅ НОВОЕ: Формируем суффикс имени файла с правильными номерами дорожек
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    if num_tracks == 1:
        track_num = start_track_index + 1
        file_suffix = f'Дорожка_{track_num}_{timestamp}'
    else:
        first_track = start_track_index + 1
        last_track = start_track_index + num_tracks
        file_suffix = f'Дорожки_{first_track}-{last_track}_{timestamp}'
    
    csv_path = os.path.join(output_dir, f'Ведомость_{file_suffix}.csv')
    with open(csv_path, 'w', encoding='utf-8') as f:
        f.write('Список плит по заказу;Использовано (с учётом резов) / остатки / обрезки\n')
        for left, right in table_rows:
            f.write(f'{left};{right}\n')

    # Инициализируем переменные путей (на случай ошибок)
    xlsx_path_v = None
    xlsx_path_p = None
    xlsx_path_breakdown = None
    
    if pd is not None:
        try:
            df_v = pd.DataFrame(table_rows, columns=col_labels)
            xlsx_path_v = os.path.join(output_dir, f'Ведомость_{file_suffix}.xlsx')
            df_v.to_excel(xlsx_path_v, index=False)
            logger.debug(f"Ведомость сохранена: {xlsx_path_v}")

            # Смета по дорожке (БЕЗ ЦЕН)
            # Определяем заголовки без столбцов цен
            price_headers = ['№', 'Наименование', 'Кол-во', 'Ед.', 'Неделя', 'Контрагент']
            # Обрезаем данные - оставляем только первые 6 столбцов
            price_rows_for_excel = [row[:6] for row in price_rows]
            df_p = pd.DataFrame(price_rows_for_excel, columns=price_headers)

            # ✅ Убрана итоговая строка с суммой
            # (теперь просто список плит без финансовой информации)

            xlsx_path_p = os.path.join(output_dir, f'Список_плит_{file_suffix}.xlsx')
            with pd.ExcelWriter(xlsx_path_p, engine='openpyxl') as writer:
                df_p.to_excel(writer, index=False, sheet_name='Список плит')
                df_v.to_excel(writer, index=False, sheet_name='Ведомость')
            logger.debug(f"Список плит сохранён: {xlsx_path_p}")
            
            # Сохраняем детальную разбивку компонентов в отдельный Excel файл (С ЦЕНАМИ)
            if breakdown_tables:
                breakdown_headers = ['Компонент', 'Расчёт', 'Сумма']  # ✅ 3 столбца
                all_breakdown_rows = []
                
                for breakdown in breakdown_tables:
                    # Заголовок с наименованием
                    all_breakdown_rows.append([breakdown['name'], '', ''])
                    # Строки таблицы (все 3 столбца)
                    for row in breakdown['rows']:
                        # Берём все 3 столбца
                        all_breakdown_rows.append(row if len(row) >= 3 else row + [''] * (3 - len(row)))
                    
                    # Пустая строка между таблицами
                    all_breakdown_rows.append(['', '', ''])
                
                # Удаляем последнюю пустую строку
                if all_breakdown_rows and all_breakdown_rows[-1] == ['', '', '']:
                    all_breakdown_rows.pop()
                
                df_breakdown = pd.DataFrame(all_breakdown_rows, columns=breakdown_headers)
                xlsx_path_breakdown = os.path.join(output_dir, f'Детальная_разбивка_{file_suffix}.xlsx')
                df_breakdown.to_excel(xlsx_path_breakdown, index=False)
                logger.debug(f"Детальная разбивка сохранена в Excel: {xlsx_path_breakdown}")
            else:
                logger.debug("breakdown_tables пустой - файл детальной разбивки не создан")
        except Exception as e:
            logger.exception(f"Ошибка при сохранении Excel файлов: {e}")
    else:
        logger.warning("pandas не установлен - Excel файлы не будут созданы")
        logger.warning("Установите: pip install pandas openpyxl")

    png_path = os.path.join(output_dir, f'Схема_{file_suffix}_КЗ.png')
    pdf_path = os.path.join(output_dir, f'Схема_{file_suffix}_КЗ.pdf')
    plt.tight_layout(rect=[0, 0, 1, 0.95])
    fig.savefig(png_path, dpi=300)
    fig.savefig(pdf_path)
    plt.close(fig)

    logger.info("Визуализация и файлы сохранены")
    logger.info(f"PNG: {png_path}")
    logger.info(f"PDF: {pdf_path}")
    logger.info(f"CSV: {csv_path}")
    if pd is not None:
        if xlsx_path_v:
            logger.info(f"XLSX (ведомость): {xlsx_path_v}")
        if xlsx_path_p:
            logger.info(f"XLSX (список плит): {xlsx_path_p}")
        if breakdown_tables and xlsx_path_breakdown:
            logger.info(f"XLSX (детальная разбивка): {xlsx_path_breakdown}")
    return png_path, pdf_path


if __name__ == '__main__':
    visualize_plan()
