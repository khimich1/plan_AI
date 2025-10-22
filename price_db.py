import os
import sqlite3
from typing import Dict, Optional

try:
    import pandas as pd
except Exception:
    pd = None


DEFAULT_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pb.db')


def init_schema(db_path: str = DEFAULT_DB) -> None:
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            'CREATE TABLE IF NOT EXISTS prices (length_dm INTEGER, load_code INTEGER, price REAL, PRIMARY KEY(length_dm, load_code))'
        )
        conn.commit()
    finally:
        conn.close()


def import_from_xlsx(xlsx_path: str, db_path: str = DEFAULT_DB, preferred_sheet: Optional[str] = '24.06.2024') -> int:
    if pd is None:
        return 0
    if not os.path.exists(xlsx_path):
        return 0
    all_sheets = pd.read_excel(xlsx_path, sheet_name=None)
    if preferred_sheet in all_sheets:
        sheets = [preferred_sheet]
    else:
        sheets = list(all_sheets.keys())

    rows = []
    for s in sheets:
        df = all_sheets[s]
        # ожидаемые заголовки
        name_col = None
        for c in df.columns:
            if str(c).strip().lower() == 'наименование' or 'наимен' in str(c).strip().lower():
                name_col = c
                break
        if name_col is None:
            continue
        headers = {}
        for c in df.columns:
            cl = str(c).strip().lower()
            if cl == '6 нагрузка':
                headers[6] = c
            elif cl == '8 нагрузка':
                headers[8] = c
            elif cl == '10 нагрузка':
                headers[10] = c
            elif cl == '12 нагрузка':
                headers[12] = c
        # если в заголовках нет явных колонок, попробуем взять подписи из первой строки
        if not headers and len(df) > 0:
            first_row = df.iloc[0]
            for c in df.columns:
                val = str(first_row.get(c, '')).strip().lower()
                if 'нагруз' in val:
                    # переименуем колонку на явное имя, чтобы дальше сработал общий цикл
                    if '6' in val:
                        headers[6] = c
                    if '8' in val:
                        headers[8] = c
                    if '10' in val:
                        headers[10] = c
                    if '12' in val:
                        headers[12] = c
            # Если нашли подписи в первой строке, пропустим её при чтении данных
            data_df = df.iloc[1:].copy()
        else:
            data_df = df
        if not headers:
            continue
        for _, row in data_df.iterrows():
            name = str(row.get(name_col, '')).strip()
            m = None
            import re
            m = re.search(r'(\d+)\s*-\s*(\d+)', name)
            if not m:
                continue
            length_dm = int(m.group(1))
            for load_code, col in headers.items():
                val = row.get(col)
                if pd.notna(val):
                    try:
                        price = float(str(val).replace(' ', '').replace(',', '.'))
                        rows.append((length_dm, load_code, price))
                    except Exception:
                        pass

    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.executemany('INSERT OR REPLACE INTO prices (length_dm, load_code, price) VALUES (?,?,?)', rows)
        conn.commit()
        return len(rows)
    finally:
        conn.close()


def get_price(length_m: float, load_code: int = 8, db_path: str = DEFAULT_DB) -> Optional[float]:
    init_schema(db_path)
    length_dm = int(round(length_m * 10))
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('SELECT price FROM prices WHERE length_dm=? AND load_code=?', (length_dm, load_code))
        row = cur.fetchone()
        if row:
            return float(row[0])
        cur.execute('SELECT price FROM prices WHERE ABS(length_dm-?)<=1 AND load_code=? ORDER BY ABS(length_dm-?) LIMIT 1', (length_dm, load_code, length_dm))
        row = cur.fetchone()
        return float(row[0]) if row else None
    finally:
        conn.close()


