#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль работы с ценами:
- Загрузка прайса из XLSX
- Работа с базой цен SQLite
- Поиск цен по параметрам
"""
import os
import re
import sqlite3

import core.config_and_data as cfg

try:
    import pandas as pd
except Exception:
    pd = None

try:
    from docx import Document
except Exception:
    Document = None


def load_price_table_from_xlsx(path: str):
    """Загружает таблицу цен вида: ключ length_dm -> {6:price,8:price,10:price,12:price}."""
    table = {}
    if pd is None:
        return table
    
    candidate_paths = []
    if os.path.exists(path):
        candidate_paths = [path]
    else:
        search_dirs = [
            os.path.dirname(path) if os.path.dirname(path) else cfg.BASE_DIR,
            cfg.BASE_DIR,
            os.path.join(cfg.BASE_DIR, 'банк знаний')
        ]
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
            
            name_col = next((c for c in df.columns if str(c).strip().lower() == 'наименование'), None) or \
                       next((c for c in df.columns if 'наимен' in str(c).lower()), None)
            if name_col is None:
                continue
            
            load_cols = {}
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
            
            simple_price_col = next((c for c in df.columns if any(k in str(c).lower() for k in ['цен', 'руб', 'стоим'])), None)
            for c in df.columns:
                cl = str(c).lower()
                m = re.search(r'(\d+)\s*нагруз', cl)
                if m:
                    load_cols[int(m.group(1))] = c
                    continue
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
                L, _ = cfg.parse_name_to_sizes(name)
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


def sync_price_xlsx_to_db(xlsx_path: str = cfg.PRICE_XLSX_PATH, db_path: str = cfg.PRICE_DB_PATH,
                          sheet_hint: str = '24.06.2024') -> int:
    """Заливает прайс из XLSX в SQLite."""
    if pd is None:
        return 0
    price_table = load_price_table_from_xlsx(xlsx_path)
    if not price_table:
        return 0
    rows = []
    for length_dm, loads in price_table.items():
        for load_code, price in loads.items():
            rows.append((int(length_dm), int(load_code), float(price)))

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('CREATE TABLE IF NOT EXISTS prices (length_dm INTEGER, load_code INTEGER, price REAL, PRIMARY KEY(length_dm, load_code))')
        cur.executemany('INSERT OR REPLACE INTO prices (length_dm, load_code, price) VALUES (?,?,?)', rows)
        conn.commit()
        return len(rows)
    finally:
        conn.close()


def find_price_from_db(length_m: float, load_code: int = 8, db_path: str = cfg.PRICE_DB_PATH) -> float:
    """Ищет цену в БД с допуском ±1 дм."""
    length_dm = int(round(length_m * 10))
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('CREATE TABLE IF NOT EXISTS prices (length_dm INTEGER, load_code INTEGER, price REAL, PRIMARY KEY(length_dm, load_code))')
        cur.execute('SELECT price FROM prices WHERE length_dm=? AND load_code=?', (length_dm, load_code))
        row = cur.fetchone()
        if row:
            return float(row[0])
        cur.execute('SELECT price FROM prices WHERE ABS(length_dm-?)<=1 AND load_code=? ORDER BY ABS(length_dm-?) LIMIT 1', (length_dm, load_code, length_dm))
        row = cur.fetchone()
        return float(row[0]) if row else None
    finally:
        conn.close()


def find_price_for_plate(price_table: dict, length_m: float, load_code: int = 8) -> float:
    """Возвращает цену по длине и нагрузке."""
    key = int(round(length_m*10))
    if key in price_table and load_code in price_table[key]:
        return price_table[key][load_code]
    for Ldm, loads in price_table.items():
        if abs(Ldm - key) <= 1 and load_code in loads:
            return loads[load_code]
    return None


def load_cut_price_from_docx(path: str) -> float:
    """Пытается извлечь цену продольного реза из DOCX."""
    if Document is None or not os.path.exists(path):
        return 0.0
    try:
        doc = Document(path)
        text = '\n'.join([p.text for p in doc.paragraphs])
        candidates = []
        for m in re.finditer(r'(рез|вдоль|продоль)[^\d]{0,20}(\d+[\s\u202f\,\.]?\d*)', text.lower()):
            try:
                val = float(m.group(2).replace(' ', '').replace('\u202f', '').replace(',', '.'))
                candidates.append(val)
            except Exception:
                pass
        if candidates:
            return float(max(candidates))
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

