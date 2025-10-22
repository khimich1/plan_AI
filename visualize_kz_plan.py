#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Визуализация КЗ-плана для одной дорожки 1.2 м:
- Укладка сегментов 1.2
- Преобразование 1.5 -> 1.2 + 0.3
- Получение 1.0 из 1.2 резом: 1.2 -> 1.0 + 0.2 (0.2 в обрезки)
- Показ суммарных резов и остатков/обрезков
- Расчёт стоимости: берём цены плит из XLSX `банк знаний/Новые цены для прайса с 19.08.24.xlsx`
  и цену реза из DOCX `банк знаний/Письмо Цены с 29.05.2024 цены на резы.docx` (если доступно)
Результат: PNG и PDF в папке "Визуализация_Раскладки".
Также выгружаются CSV/XLSX с ведомостью и сметой.
"""
import os
from datetime import datetime
import re
import math
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import sqlite3
from price_db import init_schema, import_from_xlsx, get_price

try:
    import pandas as pd  # для XLSX (опционально)
except Exception:
    pd = None

try:
    from docx import Document  # чтение цены резов из DOCX (опционально)
except Exception:
    Document = None

TRACK_LENGTH_M = 101.0
TRACK_WIDTH_M = 1.2

# Пути к прайсам (делаем абсолютными относительно файла скрипта)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRICE_XLSX_PATH = os.path.join(BASE_DIR, 'банк знаний', 'Новые цены для прайса с 19.08.24.xlsx')
CUTS_DOCX_PATH = os.path.join(BASE_DIR, 'банк знаний', 'Письмо Цены с 29.05.2024 цены на резы.docx')  # не используется в новой модели
PRICE_DB_PATH = os.path.join(BASE_DIR, 'pb.db')

# Стоимость резов
LONG_CUT_PRICE_PER_M = 460.0  # Продольный рез, руб/пог.м
TRANSVERSE_CUT_PRICE = 1200.0  # Поперечный (или скошенный) рез, руб/шт

# Данные из согласованного КЗ-плана
# 1) Плиты 1.2 м — без резов (новый заказ)
#  - ПБ 37,9-4,6-8п ×4 ×3 строки = 12 шт  → длина 3.79 м (ТРЕБУЕТ рез до 0.46, см. ниже)
#  - ПБ 58,6-10,8-8п ×5         = 5 шт   → длина 5.86 м (ТРЕБУЕТ рез до 1.08, см. ниже)
#  - ПБ 33,9-7,2-8п ×2         = 2 шт   → длина 3.39 м
# Для визуализации «как есть» оставляем только 3.39 как 1.2 м.
PLATES_1_2 = [3.39]*2

# Дополнительные целевые ширины, которые получаем продольным резом из 1.2 м
#  - 1.08 м: диапазон 1.02–1.08 (остаток 0.12–0.18)
#  - 0.46 м: диапазон 0.46–0.53 (остаток 0.67–0.74)
#  - 0.32 м: диапазон 0.26–0.32 (остаток 0.88–0.94)
#  - 0.72 м: диапазон 0.66–0.72 (остаток 0.48–0.54)
#  - 0.70 м: в рамках 0.66–0.72 (остаток ~0.50)
#  - 0.86 м: диапазон 0.86–0.92 (остаток 0.28–0.34)
PLATES_1_08 = []         # нет 1.08 в этом заказе
PLATES_0_46 = []         # нет 0.46 в этом заказе
PLATES_0_32 = [6.63]*4 + [7.83]*3
PLATES_0_72 = [5.63]*5
PLATES_0_70 = [4.65]*5
PLATES_0_86 = [6.75]*2 + [4.65]*5
# 2) Плиты 1.5 м — используем как 1.2 м (лента 0.3 образуется)
PLATES_1_5_TO_1_2 = []
# 3) Плиты 1.0 м — получаем из 1.2 (остаток 0.2 уходит в обрезки)
PLATES_1_0 = []

# Резы по плану: по одному на каждую плиту, получаемую резом
LONGITUDINAL_CUTS = (
    len(PLATES_1_5_TO_1_2) + len(PLATES_1_0) +
    len(PLATES_1_08) + len(PLATES_0_46) +
    len(PLATES_0_32) + len(PLATES_0_72) + len(PLATES_0_70) + len(PLATES_0_86)
)
LENGTH_TRIMS = 0

# Остатки и отходы
UNUSED_STRIPS_0_3_M_TOTAL = 0.0
SCRAP_STRIPS_0_2_M_TOTAL = 0.0
# Новые ленты/обрезки от резов 1.2 -> 1.08 (0.12 в обрезки) и 1.2 -> 0.46 (0.74 как используемая лента)
USABLE_STRIPS_0_74_M_TOTAL = round(sum(PLATES_0_46), 1)
USABLE_STRIPS_0_88_M_TOTAL = round(sum(PLATES_0_32), 1)
USABLE_STRIPS_0_48_M_TOTAL = round(sum(PLATES_0_72), 1)
USABLE_STRIPS_0_50_M_TOTAL = round(sum(PLATES_0_70), 1)
USABLE_STRIPS_0_34_M_TOTAL = round(sum(PLATES_0_86), 1)
SCRAP_STRIPS_0_12_M_TOTAL = round(sum(PLATES_1_08), 1)
WASTE_AREA_M2 = round(0.12 * SCRAP_STRIPS_0_12_M_TOTAL, 2)


# ---------- Утилиты для прайса ----------

def make_plate_name(length_m: float, width_m: float, reinforcement: str = '8п') -> str:
    """Формирует строку наименования в стиле прайса: 'Плиты ПБ 63-12-8п'.
    Для лент 0.3/0.2 записывает ширину как '0.3'/'0.2'."""
    length_dm = int(round(length_m * 10))
    if width_m < 0.5:
        width_str = '0.3' if abs(width_m - 0.3) < 1e-6 else ('0.2' if abs(width_m - 0.2) < 1e-6 else f'{width_m:.1f}'.replace('.', ','))
    else:
        width_dm = int(round(width_m * 10))
        width_str = str(width_dm)
    return f'Плиты ПБ {length_dm}-{width_str}-{reinforcement}'


def parse_name_to_sizes(name: str) -> tuple:
    """Достаёт (length_m, width_m) из строки прайса."""
    m = re.search(r'(\d+)-(\d+)', name.replace(',', '.'))
    if not m:
        return None, None
    return float(m.group(1)) / 10.0, float(m.group(2)) / 10.0


def load_price_table_from_xlsx(path: str):
    """Загружает таблицу цен вида: ключ length_dm -> {6:price,8:price,10:price,12:price}.
    Сравниваем только по длине. Ищем столбец 'Наименование' и ценовые колонки.
    Расширенные правила распознавания ценовых колонок:
      - '<x> нагрузка' (как раньше)
      - заголовки, где одновременно встречаются ('цен'|'руб'|'стоим') и цифра 6/8/10/12, например 'Цена 8п', 'Цена, руб (8)'
      - если явных колонок по нагрузкам нет, используем любой общий столбец цены ('цен'|'руб'|'стоим').
    """
    table = {}
    if pd is None:
        return table
    # Нормализуем путь: если файл не найден (возможны различия в диакритике "й"), пробуем подобрать кандидата в папке
    candidate_paths = []
    if os.path.exists(path):
        candidate_paths = [path]
    else:
        # пробуем несколько директорий: рядом с файлом, BASE_DIR и подкаталог 'банк знаний'
        search_dirs = []
        folder = os.path.dirname(path) if os.path.dirname(path) else BASE_DIR
        search_dirs.append(folder)
        search_dirs.append(BASE_DIR)
        search_dirs.append(os.path.join(BASE_DIR, 'банк знаний'))
        for d in search_dirs:
            if not os.path.isdir(d):
                continue
            for name in os.listdir(d):
                low = name.lower()
                if low.endswith('.xlsx') and ('нов' in low and 'цен' in low):
                    candidate_paths.append(os.path.join(d, name))
    if not candidate_paths:
        print('[ПРАЙС] Файл не найден. Искал около:', path)
        return table
    try:
        # Берём первый успешно распознанный файл
        chosen = None
        for p in candidate_paths:
            try:
                all_sheets = pd.read_excel(p, sheet_name=None)
                chosen = p
                break
            except Exception:
                continue
        if chosen is None:
            print('[ПРАЙС] Не удалось открыть ни один XLSX из кандидатов:', candidate_paths)
            return {}
        else:
            print('[ПРАЙС] Использую прайс-файл:', chosen)
        # Предпочтительно используем лист с датой, как на скрине: "24.06.2024"
        preferred_sheet = None
        for s in list(all_sheets.keys()):
            if '24.06.2024' in str(s):
                preferred_sheet = s
                break
        sheets_iter = [preferred_sheet] if preferred_sheet in all_sheets else list(all_sheets.keys())
        for sheet_name in sheets_iter:
            df = all_sheets[sheet_name]
            try:
                print(f"[ПРАЙС] Лист: {sheet_name} | колонки: {[str(c) for c in df.columns]}")
            except Exception:
                pass
            # Жёсткие имена колонок из вашего файла
            name_col = next((c for c in df.columns if str(c).strip().lower() == 'наименование'), None) or \
                       next((c for c in df.columns if 'наимен' in str(c).lower()), None)
            if name_col is None:
                try:
                    print('[ПРАЙС] Не найден столбец наименования на листе, пропускаю')
                except Exception:
                    pass
                continue
            # поищем столбцы цен по нагрузкам и/или общий столбец цены
            load_cols = {}
            # Жёсткая привязка к заголовкам: "6 нагрузка", "8 нагрузка", ...
            header_map = {6: None, 8: None, 10: None, 12: None}
            for c in df.columns:
                cl = str(c).strip().lower()
                if cl == '6 нагрузка':
                    header_map[6] = c
                elif cl == '8 нагрузка':
                    header_map[8] = c
                elif cl == '10 нагрузка':
                    header_map[10] = c
                elif cl == '12 нагрузка':
                    header_map[12] = c
            for k,v in header_map.items():
                if v is not None:
                    load_cols[k] = v
            # общий ценовой столбец (если он один для всех нагрузок)
            simple_price_col = next((c for c in df.columns if any(k in str(c).lower() for k in ['цен', 'руб', 'стоим'])), None)
            for c in df.columns:
                cl = str(c).lower()
                # вариант 1: "<число> нагрузка"
                m = re.search(r'(\d+)\s*нагруз', cl)
                if m:
                    load_cols[int(m.group(1))] = c
                    continue
                # вариант 2: заголовок с упоминанием цены и кода нагрузки (6/8/10/12)
                m2 = re.search(r'(?:цен|руб|стоим)[^\d]{0,10}(6|8|10|12)\b', cl)
                if not m2:
                    m2 = re.search(r'\b(6|8|10|12)[^\d]{0,10}(?:цен|руб|стоим)', cl)
                if m2:
                    try:
                        load_cols[int(m2.group(1))] = c
                    except Exception:
                        pass
            found_rows = 0
            for _, row in df.iterrows():
                name = str(row.get(name_col, '')).strip()
                if not name:
                    continue
                # В прайсе имена вида "ПБ 38-12" — достаём длину из этих двух чисел
                L, _ = parse_name_to_sizes(name)
                if L is None:
                    continue
                key = int(round(L*10))
                price_by_load = {}
                if load_cols:
                    for load_code, col in load_cols.items():
                        try:
                            val = row[col]
                            if pd.notna(val):
                                price_by_load[load_code] = float(str(val).replace(' ', '').replace(',', '.'))
                        except Exception:
                            pass
                elif simple_price_col is not None:
                    try:
                        val = row[simple_price_col]
                        if pd.notna(val):
                            price_val = float(str(val).replace(' ', '').replace(',', '.'))
                            # одинаковая цена для всех нагрузок, если нет отдельных столбцов
                            for load_code in [6, 8, 10, 12]:
                                price_by_load[load_code] = price_val
                    except Exception:
                        pass
                if price_by_load:
                    table[key] = price_by_load
                    found_rows += 1
            try:
                print(f"[ПРАЙС] Считано позиций на листе: {found_rows}")
            except Exception:
                pass
    except Exception:
        return {}
    return table


def sync_price_xlsx_to_db(xlsx_path: str = PRICE_XLSX_PATH, db_path: str = PRICE_DB_PATH,
                          sheet_hint: str = '24.06.2024') -> int:
    """Заливает прайс из XLSX в SQLite (`prices`):
    columns: length_dm INTEGER, load_code INTEGER, price REAL.
    Возвращает число записанных строк."""
    if pd is None:
        return 0
    # Загружаем словарь через уже отлаженную функцию
    price_table = load_price_table_from_xlsx(xlsx_path)
    if not price_table:
        return 0
    # Плоский список (length_dm, load, price)
    rows = []
    for length_dm, loads in price_table.items():
        for load_code, price in loads.items():
            rows.append((int(length_dm), int(load_code), float(price)))

    # Пишем в SQLite
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('CREATE TABLE IF NOT EXISTS prices (length_dm INTEGER, load_code INTEGER, price REAL, PRIMARY KEY(length_dm, load_code))')
        cur.executemany('INSERT OR REPLACE INTO prices (length_dm, load_code, price) VALUES (?,?,?)', rows)
        conn.commit()
        return len(rows)
    finally:
        conn.close()


def find_price_from_db(length_m: float, load_code: int = 8, db_path: str = PRICE_DB_PATH) -> float:
    """Ищет цену в БД с допуском ±1 дм, если нет точной длины."""
    length_dm = int(round(length_m * 10))
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        # гарантируем схему
        cur.execute('CREATE TABLE IF NOT EXISTS prices (length_dm INTEGER, load_code INTEGER, price REAL, PRIMARY KEY(length_dm, load_code))')
        cur.execute('SELECT price FROM prices WHERE length_dm=? AND load_code=?', (length_dm, load_code))
        row = cur.fetchone()
        if row:
            return float(row[0])
        # допуск ±1 дм
        cur.execute('SELECT price FROM prices WHERE ABS(length_dm-?)<=1 AND load_code=? ORDER BY ABS(length_dm-?) LIMIT 1', (length_dm, load_code, length_dm))
        row = cur.fetchone()
        return float(row[0]) if row else None
    finally:
        conn.close()


def find_price_for_plate(price_table: dict, length_m: float, load_code: int = 8) -> float:
    """Возвращает цену по длине и нагрузке (ширину игнорируем)."""
    key = int(round(length_m*10))
    if key in price_table and load_code in price_table[key]:
        return price_table[key][load_code]
    # ближайшая длина ±1 дм
    for Ldm, loads in price_table.items():
        if abs(Ldm - key) <= 1 and load_code in loads:
            return loads[load_code]
    return None


def approximate_weight_kg(length_m: float, width_m: float, thickness_m: float = 0.22) -> float:
    volume = length_m * width_m * thickness_m
    return round(volume * 2400, 1)


def optimize_cuts_pulp(orders: dict) -> list[dict]:
    """Оптимизирует раскрой стандартных плит 1200мм по ширине с помощью PuLP.

    Вход:
      orders: словарь {номинальная_ширина_мм: требуемое_количество}, например {300:4,500:3,700:2,900:2}

    Модель:
      - Есть стандартная плита шириной 1200мм. Один рез по выбранному типу даёт 2 куска: main и rest (остаток).
      - Для каждого типа реза i создаём целую переменную x_i — сколько плит резать этим способом (сколько исходных плат 1200мм).
      - Для каждого заказа w создаём переменную y_w — сколько плит нужной ширины будет выполнено.
      - Для распределения кусков вводим целочисленные переменные a_iw_main и a_iw_rest —
        сколько кусков типа main/rest от реза i пойдут на заказ ширины w.

    Ограничения:
      - y_w = sum_i (a_iw_main + a_iw_rest) и y_w >= orders[w]
      - Для каждого i: sum_w a_iw_main <= x_i  (main-кусков от i не больше, чем x_i)
      - Для каждого i: sum_w a_iw_rest <= x_i  (rest-кусков от i не больше, чем x_i)
      - a_iw_main создаются только если номинал w попадает в допустимый диапазон main у реза i (иначе переменная не создаётся)
      - a_iw_rest аналогично для диапазона rest

    Целевая функция:
      - минимизировать суммарное число резов sum_i x_i
      - плюс штраф 1000 за каждый неиспользованный остаток: для i это (x_i - sum_w a_iw_rest)
        Итого:  sum_i x_i + 1000 * sum_i (x_i - sum_w a_iw_rest)

    Возвращает список словарей с оптимальным количеством резов по каждому типу:
      [{"cut_id": "cut300", "qty": 2, "main_range": (260,320), "rest_range": (880,940)}, ...]

    Примечания:
      - Функция самодостаточна: при отсутствии pulp печатает предупреждение и возвращает пустой результат.
      - Ширины и диапазоны — в миллиметрах.
    """
    try:
        from pulp import LpProblem, LpMinimize, LpVariable, LpInteger, lpSum, LpStatusOptimal, value
    except Exception:
        print('[OPT] Модуль pulp не установлен. Пропускаю оптимизацию раскроя.')
        return []

    # Описание доступных типов резов (в мм)
    CUT_OPTIONS = [
        {"id": "cut300", "main": (260, 320), "rest": (880, 940)},
        {"id": "cut500", "main": (460, 530), "rest": (670, 740)},
        {"id": "cut700", "main": (660, 720), "rest": (480, 540)},
        {"id": "cut900", "main": (860, 920), "rest": (280, 340)},
    ]

    # Нормализация входа: ключи — int, значения — int >= 0
    orders_mm = {}
    for k, v in (orders or {}).items():
        try:
            w = int(k)
            q = int(v)
            if q > 0:
                orders_mm[w] = q
        except Exception:
            continue
    if not orders_mm:
        return []

    widths = sorted(orders_mm.keys())

    # Подготовим списки источников для каждого заказа w
    # Какие резы могут дать main-кусок для w, и какие — rest-кусок для w
    main_sources = {w: [] for w in widths}
    rest_sources = {w: [] for w in widths}
    for i, opt in enumerate(CUT_OPTIONS):
        m_lo, m_hi = opt["main"]
        r_lo, r_hi = opt["rest"]
        for w in widths:
            if m_lo <= w <= m_hi:
                main_sources[w].append(i)
            if r_lo <= w <= r_hi:
                rest_sources[w].append(i)

    # Если для какого-то w вообще нет источников — задача невыполнима, вернём пусто
    for w in widths:
        if not main_sources[w] and not rest_sources[w]:
            print(f"[OPT] Для ширины {w} мм нет допустимых резов. Задача может быть невыполнима.")

    # Модель
    prob = LpProblem('cut_optimization', LpMinimize)

    # Переменные x_i — количество исходных плит 1200мм, распиленных типом i
    x = {
        i: LpVariable(f"x_{CUT_OPTIONS[i]['id']}", lowBound=0, cat=LpInteger)
        for i in range(len(CUT_OPTIONS))
    }

    # Переменные распределения кусков по заказам: a_iw_main, a_iw_rest
    a_main = {i: {} for i in range(len(CUT_OPTIONS))}
    a_rest = {i: {} for i in range(len(CUT_OPTIONS))}
    for i, opt in enumerate(CUT_OPTIONS):
        # main
        for w in widths:
            if i in main_sources[w]:
                a_main[i][w] = LpVariable(f"a_main_{opt['id']}_{w}", lowBound=0, cat=LpInteger)
        # rest
        for w in widths:
            if i in rest_sources[w]:
                a_rest[i][w] = LpVariable(f"a_rest_{opt['id']}_{w}", lowBound=0, cat=LpInteger)

    # Переменные y_w — выполненное количество для каждого заказа
    y = {w: LpVariable(f"y_{w}", lowBound=0, cat=LpInteger) for w in widths}

    # Связи y_w с распределениями
    for w in widths:
        lhs = []
        for i in main_sources[w]:
            lhs.append(a_main[i][w])
        for i in rest_sources[w]:
            lhs.append(a_rest[i][w])
        if lhs:
            prob += (y[w] == lpSum(lhs)), f"def_y_{w}"
        else:
            # Нет источников — y_w фиксируем 0 (оставим, но требование всё равно зададим ниже)
            prob += (y[w] == 0), f"def_y_{w}"

    # Требуем обеспечить объёмы заказов
    for w in widths:
        prob += (y[w] >= orders_mm[w]), f"demand_{w}"

    # Ограничение по наличию кусков каждого типа реза: на каждый x_i есть не более x_i main и x_i rest
    for i, opt in enumerate(CUT_OPTIONS):
        if a_main[i]:
            prob += (lpSum(a_main[i].values()) <= x[i]), f"main_cap_{opt['id']}"
        else:
            # если переменных нет, то суммарно 0 <= x[i] всегда верно; ограничение опускаем
            pass
        if a_rest[i]:
            prob += (lpSum(a_rest[i].values()) <= x[i]), f"rest_cap_{opt['id']}"

    # Цель: минимизировать число резов + штраф за неиспользованный остаток
    # penalty = 1000 * sum_i (x_i - sum_w a_iw_rest)
    penalty_terms = []
    for i, opt in enumerate(CUT_OPTIONS):
        used_rest = lpSum(a_rest[i].values()) if a_rest[i] else 0
        penalty_terms.append(x[i] - used_rest)

    objective = lpSum(x.values()) + 1000 * lpSum(penalty_terms)
    prob += objective

    # Решаем
    prob.solve()

    # Проверка статуса
    try:
        status = prob.status
        # 1 — LpStatusOptimal в pulp
        if status != 1:
            print(f"[OPT] Решение не оптимально. Статус: {status}")
    except Exception:
        pass

    # Собираем результат
    result = []
    for i, opt in enumerate(CUT_OPTIONS):
        try:
            qty = int(round(x[i].value()))
        except Exception:
            qty = 0
        result.append({
            "cut_id": opt["id"],
            "qty": qty if qty > 0 else 0,
            "main_range": tuple(opt["main"]),
            "rest_range": tuple(opt["rest"]),
        })

    return result

def load_cut_price_from_docx(path: str) -> float:
    """Пытается извлечь цену продольного реза из DOCX. Возвращает 0, если не удалось."""
    if Document is None or not os.path.exists(path):
        return 0.0
    try:
        doc = Document(path)
        text = '\n'.join([p.text for p in doc.paragraphs])
        # Находим фразы про рез/вдоль/продольн и число рядом
        candidates = []
        for m in re.finditer(r'(рез|вдоль|продоль)[^\d]{0,20}(\d+[\s\u202f\,\.]?\d*)', text.lower()):
            try:
                val = float(m.group(2).replace(' ', '').replace('\u202f', '').replace(',', '.'))
                candidates.append(val)
            except Exception:
                pass
        if candidates:
            return float(max(candidates))
        # также ищем в таблицах
        for table in doc.tables:
            for row in table.rows:
                row_text = ' '.join(c.text for c in row.cells).lower()
                if any(k in row_text for k in ['рез', 'вдоль', 'продоль']):
                    nums = re.findall(r'\d+[\s\u202f\,\.]?\d*', row_text)
                    for s in nums:
                        try:
                            val = float(s.replace(' ', '').replace('\u202f', '').replace(',', '.'))
                            candidates.append(val)
                        except Exception:
                            pass
        return float(max(candidates)) if candidates else 0.0
    except Exception:
        return 0.0


def build_procurement_items():
    """Формирует реальные позиции закупки с учётом назначения реза.
    Возвращает список dict: {length, width, qty, cuts_per_plate, purpose}.
    """
    items = []
    # 1.2 без реза
    for L in PLATES_1_2:
        items.append({'length': round(L, 1), 'width': 1.2, 'qty': 1, 'long_cuts': 0, 'trans_cuts': 0, 'purpose': 'as_is'})
    # 1.5 -> 1.2 + 0.3: две позиции
    #   - ПБ L-12 (без реза)
    #   - Лента L-0.3 (как отдельная позиция, с продольным резом 1 шт)
    for L in PLATES_1_5_TO_1_2:
        items.append({'length': round(L, 1), 'width': 1.2, 'qty': 1, 'long_cuts': 0, 'trans_cuts': 0, 'purpose': 'to_1_2_main'})
        items.append({'length': round(L, 1), 'width': 0.3, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_2_strip'})
    # 1.2 -> 1.0 + 0.2: две позиции (вернули прежнее поведение)
    #   - ПБ L-10-8п — с продольным резом
    #   - Лента L-0.2-8п — как отдельная позиция, тарифицируется и учитывает 1 продольный рез
    for L in PLATES_1_0:
        items.append({'length': round(L, 1), 'width': 1.0, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_0_main'})
        items.append({'length': round(L, 1), 'width': 0.2, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_0_strip'})
    # агрегируем по (L,W,cuts)
    agg = {}
    for it in items:
        key = (it['length'], it['width'], it['long_cuts'], it['trans_cuts'])
        agg[key] = agg.get(key, 0) + it['qty']
    result = []
    for (L, W, long_cuts, trans_cuts), qty in sorted(agg.items(), key=lambda x: (x[0][1], x[0][0])):
        result.append({'length': L, 'width': W, 'qty': qty, 'long_cuts': long_cuts, 'trans_cuts': trans_cuts})
    return result


def build_price_rows(price_table: dict, reinforcement_code: int = 8):
    """Формирует строки сметы на базе реальных позиций закупки.
    Цена позиции = цена плиты (по длине/ширине и нагрузке) + цена резов на плиту.
    Возвращает (rows, total_sum)."""
    items = build_procurement_items()
    rows = []
    total = 0.0
    idx = 1
    for it in items:
        L, W, qty = it['length'], it['width'], it['qty']
        long_cuts, trans_cuts = it['long_cuts'], it['trans_cuts']
        name = make_plate_name(L, W)
        # 1) пытаемся взять из БД, если есть прайс-таблица
        # Используем БД (цены полностью перенесены)
        db_price = get_price(L, reinforcement_code, PRICE_DB_PATH)
        # 2) fallback — из XLSX-таблицы, если БД пустая
        base_price = db_price if db_price is not None else (find_price_for_plate(price_table, L, reinforcement_code) or 0.0)
        cuts_cost = long_cuts * (LONG_CUT_PRICE_PER_M * L) + trans_cuts * TRANSVERSE_CUT_PRICE
        unit_price = base_price + cuts_cost
        weight = approximate_weight_kg(L, W)
        row_sum = unit_price * qty
        total += row_sum
        rows.append([
            idx,
            name,
            qty,
            'шт',
            f'{weight:.0f}',
            f'{unit_price:,.2f}'.replace(',', ' ').replace('.', ','),
            f'{row_sum:,.2f}'.replace(',', ' ').replace('.', ',')
        ])
        idx += 1
    return rows, total


# ---------- Рендер ----------

def _draw_segment(ax, x0: float, length: float, color: str, label: str, y: float = 0.0, height: float = TRACK_WIDTH_M):
    rect = patches.Rectangle((x0, y), length, height, linewidth=1, edgecolor='black', facecolor=color, alpha=0.85)
    ax.add_patch(rect)
    ax.text(x0 + length/2, y + height/2, label, ha='center', va='center', fontsize=8, color='white', weight='bold')


def _draw_strip(ax, x0: float, length: float, width: float, color: str, label: str, y: float, hatch: str = None):
    rect = patches.Rectangle((x0, y), length, width, linewidth=0.8, edgecolor='black', facecolor=color, alpha=0.9, hatch=hatch)
    ax.add_patch(rect)
    ax.text(x0 + length/2, y + width/2, label, ha='center', va='center', fontsize=7, color='white', weight='bold')


def build_layout_sequence():
    """Формирует последовательность сегментов вдоль дорожки."""
    sequence = []

    for L in PLATES_1_2:
        sequence.append({
            'type': '1.2', 'length': L, 'strip': None, 'label': f'1.2×{L:.1f}'
        })

    for L in PLATES_1_5_TO_1_2:
        sequence.append({
            'type': '1.5->1.2', 'length': L, 'strip': {'width': 0.3, 'label': '0.3'}, 'label': f'1.5→1.2 {L:.1f}'
        })

    for L in PLATES_1_0:
        sequence.append({
            'type': '1.2->1.0', 'length': L, 'strip': {'width': 0.2, 'label': '0.2'}, 'label': f'1.2→1.0 {L:.1f}'
        })

    # Новые: 1.2 → 1.08 (остаток 0.12 в обрезки)
    for L in globals().get('PLATES_1_08', []):
        sequence.append({
            'type': '1.2->1.08', 'length': L, 'strip': {'width': 0.12, 'label': '0.12'}, 'label': f'1.2→1.08 {L:.1f}'
        })

    # Новые: 1.2 → 0.46 (остаток 0.74 — используемая лента)
    for L in globals().get('PLATES_0_46', []):
        sequence.append({
            'type': '1.2->0.46', 'length': L, 'strip': {'width': 0.74, 'label': '0.74'}, 'label': f'1.2→0.46 {L:.1f}'
        })

    # Новые: 1.2 → 0.32 (остаток 0.88 — используемая лента)
    for L in globals().get('PLATES_0_32', []):
        sequence.append({
            'type': '1.2->0.32', 'length': L, 'strip': {'width': 0.88, 'label': '0.88'}, 'label': f'1.2→0.32 {L:.1f}'
        })

    # Новые: 1.2 → 0.72 (остаток 0.48 — используемая лента)
    for L in globals().get('PLATES_0_72', []):
        sequence.append({
            'type': '1.2->0.72', 'length': L, 'strip': {'width': 0.48, 'label': '0.48'}, 'label': f'1.2→0.72 {L:.1f}'
        })

    # Новые: 1.2 → 0.70 (остаток ≈0.50 — используемая лента)
    for L in globals().get('PLATES_0_70', []):
        sequence.append({
            'type': '1.2->0.70', 'length': L, 'strip': {'width': 0.50, 'label': '0.50'}, 'label': f'1.2→0.70 {L:.1f}'
        })

    # Новые: 1.2 → 0.86 (остаток 0.34 — используемая лента)
    for L in globals().get('PLATES_0_86', []):
        sequence.append({
            'type': '1.2->0.86', 'length': L, 'strip': {'width': 0.34, 'label': '0.34'}, 'label': f'1.2→0.86 {L:.1f}'
        })

    return sequence


def visualize_plan(output_dir: str = 'Визуализация_Раскладки'):
    # Демонстрационный вызов оптимизации раскроя (не влияет на визуализацию)
    try:
        optimized = optimize_cuts_pulp({300: 4, 500: 3, 700: 2, 900: 2})
        print("Оптимальные резы:", optimized)
    except Exception as e:
        print("[OPT] Ошибка при оптимизации:", e)

    os.makedirs(output_dir, exist_ok=True)

    # Загружаем прайс
    price_table = load_price_table_from_xlsx(PRICE_XLSX_PATH)
    # Полная миграция цен в БД (однократно при запуске)
    try:
        init_schema(PRICE_DB_PATH)
        written = import_from_xlsx(PRICE_XLSX_PATH, PRICE_DB_PATH)
        if written:
            print(f'[ПРАЙС->БД] записано строк: {written}')
    except Exception:
        pass
    # Не читаем DOCX: используем константы LONG_CUT_PRICE_PER_M/TRANSVERSE_CUT_PRICE

    # Строим смету
    price_rows, total_sum = build_price_rows(price_table)

    seq = build_layout_sequence()
    total_length = sum(s['length'] for s in seq)

    # Настройка фигуры: 4 строки — дорожка, сводка, таблица ведомости, таблица сметы
    fig = plt.figure(figsize=(22, 14))
    gs = fig.add_gridspec(4, 1, height_ratios=[3.0, 1.0, 1.4, 1.8])
    ax_track = fig.add_subplot(gs[0, 0])
    ax_strips = fig.add_subplot(gs[1, 0])
    ax_table = fig.add_subplot(gs[2, 0])
    ax_price = fig.add_subplot(gs[3, 0])
    fig.suptitle('КЗ: Дорожка 1 (ширина 1.2 м) — раскладка, резы, ведомости и смета', fontsize=16, fontweight='bold')

    # Дорожка
    ax_track.set_xlim(0, max(total_length + 2, TRACK_LENGTH_M))
    ax_track.set_ylim(0, TRACK_WIDTH_M + 0.8)
    ax_track.set_aspect('auto')
    ax_track.spines['top'].set_visible(False)
    ax_track.spines['right'].set_visible(False)
    ax_track.set_yticks([0, 0.6, 1.2])
    ax_track.set_yticklabels(['0', '0.6', '1.2 м'])

    # Горизонтальная линейка и сетка по метрам
    ax_track.set_xlabel('Длина (м)')
    ax_track.set_xticks(range(0, int(max(total_length, TRACK_LENGTH_M)) + 1, 5))
    ax_track.grid(axis='x', linestyle=':', linewidth=0.5, alpha=0.5)

    # Рисуем границу дорожки
    track_rect = patches.Rectangle((0, 0), TRACK_LENGTH_M, TRACK_WIDTH_M, linewidth=2, edgecolor='black', facecolor='none', linestyle='--')
    ax_track.add_patch(track_rect)

    # Цвета типов
    colors = {
        '1.2': '#2ecc71',        # зелёный
        '1.5->1.2': '#e67e22',   # оранжевый
        '1.2->1.0': '#f1c40f',   # жёлтый
        '1.2->1.08': '#8e44ad',  # фиолетовый
        '1.2->0.46': '#1abc9c',  # бирюзовый
        '1.2->0.32': '#e84393',  # розовый
        '1.2->0.72': '#2d3436',  # тёмно-серый
        '1.2->0.70': '#00b894',  # зелёный оттенок
        '1.2->0.86': '#0984e3',  # синий оттенок
        'strip_0.3': '#e74c3c',  # красный
        'strip_0.2': '#3498db',  # синий
    }

    # Рисуем последовательность
    x = 0.0
    for item in seq:
        base_color = colors[item['type']]
        _draw_segment(ax_track, x, item['length'], base_color, item['label'])
        if item['strip']:
            strip_w = item['strip']['width']
            # Подбираем цвет/штриховку для новых лент
            if abs(strip_w - 0.3) < 1e-6:
                strip_color = colors['strip_0.3']; hatch = '//'
            elif abs(strip_w - 0.2) < 1e-6:
                strip_color = colors['strip_0.2']; hatch = 'xx'
            elif abs(strip_w - 0.12) < 1e-6:
                strip_color = '#9b59b6'; hatch = '...'
            elif abs(strip_w - 0.74) < 1e-6:
                strip_color = '#16a085'; hatch = '++'
            elif abs(strip_w - 0.88) < 1e-6:
                strip_color = '#e84393'; hatch = 'oo'
            elif abs(strip_w - 0.48) < 1e-6:
                strip_color = '#2d3436'; hatch = '__'
            elif abs(strip_w - 0.50) < 1e-6:
                strip_color = '#00b894'; hatch = '+++'
            elif abs(strip_w - 0.34) < 1e-6:
                strip_color = '#0984e3'; hatch = '//.'
            else:
                strip_color = '#7f8c8d'; hatch = '..'
            _draw_strip(ax_track, x, item['length'], strip_w, strip_color, item['strip']['label'], y=TRACK_WIDTH_M + 0.05, hatch=hatch)
        x += item['length']

    # Легенда
    legend_patches = [
        patches.Patch(facecolor=colors['1.2'], edgecolor='black', label='1.2 м (без реза)'),
        patches.Patch(facecolor=colors['1.5->1.2'], edgecolor='black', label='1.5 → 1.2 (лента 0.3)'),
        patches.Patch(facecolor=colors['1.2->1.0'], edgecolor='black', label='1.2 → 1.0 (лента 0.2 в обрезки)'),
        patches.Patch(facecolor=colors['strip_0.3'], edgecolor='black', label='Лента 0.3'),
        patches.Patch(facecolor=colors['strip_0.2'], edgecolor='black', label='Лента 0.2 (обрезки)'),
    ]
    ax_track.legend(handles=legend_patches, loc='upper right')

    # Панель сводки
    ax_strips.set_xlim(0, 100)
    ax_strips.set_ylim(0, 1)
    ax_strips.axis('off')

    txt = (
        f"Длина по плану: {total_length:.1f} м  |  Продольных резов: {LONGITUDINAL_CUTS}  |  Подрезов по длине: {LENGTH_TRIMS}\n"
        f"Остатки лент 0.3: {UNUSED_STRIPS_0_3_M_TOTAL:.1f} пог.м  |  Обрезки 0.2: {SCRAP_STRIPS_0_2_M_TOTAL:.1f} пог.м (≈ {WASTE_AREA_M2:.2f} м²)"
    )
    ax_strips.text(0.02, 0.6, txt, ha='left', va='center', fontsize=12,
                   bbox=dict(boxstyle='round,pad=0.6', facecolor='#f8f9fa', edgecolor='#bdc3c7'))

    leftovers = (
        "Остатки/обрезки:\n"
        "Ленты 0.3: 3×3.8 м; 1×2.9 м\n"
        f"Ленты 0.2 (обрезки): {', '.join(f'{L:.1f} м' for L in PLATES_1_0)}"
    )
    ax_strips.text(0.02, 0.15, leftovers, ha='left', va='center', fontsize=11,
                   bbox=dict(boxstyle='round,pad=0.5', facecolor='#eef7ff', edgecolor='#a3c9ff'))

    # Таблица: ведомость (заказ vs использовано/обрезки)
    ax_table.axis('off')

    order_list = [
        'Заказ 1.5: 3.8×3; 2.9×1',
        'Заказ 1.2: 6.3×2; 3.8×2',
        'Заказ 1.0: 6.3×1; 5.3×2; 3.8×1; 2.8×1',
    ]

    used_list = [
        '1.2 без реза: 6.3×2; 3.8×2',
        '1.5→1.2: 3.8×3; 2.9×1 (остаток 0.3)',
        '1.2→1.0: 6.3×1; 5.3×2; 3.8×1; 2.8×1 (остаток 0.2 → обрезки)',
        f'Резы: продольных {LONGITUDINAL_CUTS}; подрезов {LENGTH_TRIMS}; обрезки 0.2 ≈ {WASTE_AREA_M2:.2f} м²',
    ]

    rows = max(len(order_list), len(used_list))
    table_rows = []
    for i in range(rows):
        left = order_list[i] if i < len(order_list) else ''
        right = used_list[i] if i < len(used_list) else ''
        table_rows.append([left, right])

    col_labels = ['Список плит по заказу', 'Использовано (с учётом резов) / остатки / обрезки']

    table = ax_table.table(cellText=table_rows, colLabels=col_labels, loc='center', cellLoc='left', colLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    table.scale(1, 1.5)

    # Таблица: смета (как на фото)
    ax_price.axis('off')
    price_headers = ['№', 'Наименование', 'Кол-во', 'Ед.', 'Вес(кг)', 'Цена', 'Сумма']
    price_table = ax_price.table(cellText=price_rows, colLabels=price_headers, loc='center', cellLoc='center', colLoc='center')
    price_table.auto_set_font_size(False)
    price_table.set_fontsize(10)
    price_table.scale(1, 1.4)
    # Диагностика: если какие-то цены не найдены (0), подсказка в заголовке
    not_priced = any(row[5].strip().startswith('0') for row in price_rows)
    title = f'Итоговая стоимость: {total_sum:,.2f} ₽'.replace(',', ' ').replace('.', ',')
    if not_priced:
        title += ' (внимание: не найдены цены для некоторых позиций — проверьте прайс)'
    ax_price.set_title(title, fontsize=12, pad=10)

    # Экспорт таблиц
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    csv_path = os.path.join(output_dir, f'Ведомость_Дорожка_1_{timestamp}.csv')
    with open(csv_path, 'w', encoding='utf-8') as f:
        f.write('Список плит по заказу;Использовано (с учётом резов) / остатки / обрезки\n')
        for left, right in table_rows:
            f.write(f'{left};{right}\n')

    if pd is not None:
        try:
            df_v = pd.DataFrame(table_rows, columns=col_labels)
            xlsx_path_v = os.path.join(output_dir, f'Ведомость_Дорожка_1_{timestamp}.xlsx')
            df_v.to_excel(xlsx_path_v, index=False)

            df_p = pd.DataFrame(price_rows, columns=price_headers)
            xlsx_path_p = os.path.join(output_dir, f'Смета_Дорожка_1_{timestamp}.xlsx')
            with pd.ExcelWriter(xlsx_path_p, engine='openpyxl') as writer:
                df_p.to_excel(writer, index=False, sheet_name='Смета')
                df_v.to_excel(writer, index=False, sheet_name='Ведомость')
        except Exception:
            pass

    # Сохранение изображений
    png_path = os.path.join(output_dir, f'Схема_Дорожка_1_КЗ_{timestamp}.png')
    pdf_path = os.path.join(output_dir, f'Схема_Дорожка_1_КЗ_{timestamp}.pdf')
    plt.tight_layout(rect=[0, 0, 1, 0.95])
    fig.savefig(png_path, dpi=300)
    fig.savefig(pdf_path)
    plt.close(fig)

    print('[ГОТОВО] Визуализация и файлы сохранены:')
    print('  PNG:', png_path)
    print('  PDF:', pdf_path)
    print('  CSV:', csv_path)
    if pd is not None:
        print('  XLSX (ведомость):', os.path.join(output_dir, f'Ведомость_Дорожка_1_{timestamp}.xlsx'))
        print('  XLSX (смета):', os.path.join(output_dir, f'Смета_Дорожка_1_{timestamp}.xlsx'))
    return png_path, pdf_path


if __name__ == '__main__':
    visualize_plan()
