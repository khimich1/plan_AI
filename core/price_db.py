import math
import os
import re
import sqlite3
from typing import Dict, List, Optional, Tuple

try:
    import pandas as pd
except Exception:
    pd = None

PlatePriceRow = Tuple[int, int, float]


def length_m_to_price_length_dm(length_m: float) -> int:
    """
    Длина плиты в метрах → целый ключ length_dm в БД цен (дециметры, с потолком).

    Например 2.73 м → 28 дм, чтобы совпадать с прайсом, где длина в дм округляется вверх.
    """
    return int(math.ceil(round(float(length_m) * 10.0, 12)))


# Путь к базе данных в корне проекта (на уровень выше core/)
DEFAULT_DB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'pb.db')


def _connect(db_path: str) -> sqlite3.Connection:
    """
    Безопасное подключение к SQLite.

    - WAL уменьшает риск повреждения БД при сбоях
    - foreign_keys включает поддержку внешних ключей (где они используются)
    """
    conn = sqlite3.connect(db_path)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA foreign_keys = ON')
    return conn


def _column_has_plate_names(series) -> bool:
    for val in series.dropna().astype(str).head(15):
        if re.search(r"П[БК]\s*\d", val, re.IGNORECASE) or re.search(
            r"\d+(?:[,.]\d+)?\s*-\s*\d+", val
        ):
            return True
    return False


def _find_load_columns(df) -> Dict[int, object]:
    headers: Dict[int, object] = {}
    for col in df.columns:
        col_label = str(col).strip().lower()
        if col_label == "6 нагрузка":
            headers[6] = col
        elif col_label == "8 нагрузка":
            headers[8] = col
        elif col_label == "10 нагрузка":
            headers[10] = col
        elif col_label in {"12 нагрузка", "12,5 нагрузка", "12.5 нагрузка"}:
            headers[12] = col
        else:
            match = re.search(r"(\d+)\s*нагруз", col_label)
            if match:
                headers[int(match.group(1))] = col
    if headers:
        return headers

    if len(df) == 0:
        return headers

    first_row = df.iloc[0]
    for col in df.columns:
        val = str(first_row.get(col, "")).strip().lower()
        if "нагруз" not in val:
            continue
        for load_code in (6, 8, 10, 12):
            if str(load_code) in val or (load_code == 12 and ("12,5" in val or "12.5" in val)):
                headers[load_code] = col
    return headers


def _find_name_column(df):
    for col in df.columns:
        label = str(col).strip().lower()
        if label == "наименование" or "наимен" in label:
            if _column_has_plate_names(df[col]):
                return col
    for col in df.columns:
        if _column_has_plate_names(df[col]):
            return col
    return df.columns[0] if len(df.columns) else None


def _length_dm_from_plate_name(name: str) -> Optional[int]:
    match = re.search(r"(\d+(?:[,.]\d+)?)\s*-\s*\d+", str(name or ""))
    if not match:
        return None
    length_val = float(match.group(1).replace(",", "."))
    return int(round(length_val))


def _rows_from_price_dataframe(df) -> List[PlatePriceRow]:
    if df is None or df.empty:
        return []

    load_columns = _find_load_columns(df)
    name_col = _find_name_column(df)
    if name_col is None or not load_columns:
        return []

    skip_first_row = not any(
        str(col).strip().lower().endswith("нагрузка") or "нагруз" in str(col).strip().lower()
        for col in load_columns.values()
    )
    data_df = df.iloc[1:].copy() if skip_first_row else df

    rows: List[PlatePriceRow] = []
    for _, row in data_df.iterrows():
        name = str(row.get(name_col, "")).strip()
        length_dm = _length_dm_from_plate_name(name)
        if length_dm is None:
            continue
        for load_code, col in load_columns.items():
            val = row.get(col)
            if pd is not None and pd.notna(val):
                try:
                    price = float(str(val).replace(" ", "").replace(",", "."))
                    rows.append((length_dm, load_code, price))
                except (TypeError, ValueError):
                    pass
    return rows


def parse_plate_price_rows_from_xlsx(
    xlsx_path: str,
    preferred_sheet: Optional[str] = None,
) -> List[PlatePriceRow]:
    """Читает прайс плит из XLSX, поддерживает старый и новый формат листа."""
    if pd is None or not os.path.exists(xlsx_path):
        return []

    all_sheets = pd.read_excel(xlsx_path, sheet_name=None)
    if preferred_sheet in all_sheets:
        sheet_names = [preferred_sheet]
    else:
        sheet_names = list(all_sheets.keys())

    best_rows: List[PlatePriceRow] = []
    for sheet_name in sheet_names:
        for header_row in (0, 1):
            df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=header_row)
            rows = _rows_from_price_dataframe(df)
            if len(rows) > len(best_rows):
                best_rows = rows
    return best_rows


def init_schema(db_path: str = DEFAULT_DB) -> None:
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            'CREATE TABLE IF NOT EXISTS prices (length_dm INTEGER, load_code INTEGER, price REAL, PRIMARY KEY(length_dm, load_code))'
        )
        conn.commit()
    finally:
        conn.close()


def import_from_xlsx(
    xlsx_path: str,
    db_path: str = DEFAULT_DB,
    preferred_sheet: Optional[str] = None,
) -> int:
    rows = parse_plate_price_rows_from_xlsx(xlsx_path, preferred_sheet=preferred_sheet)
    if not rows:
        return 0

    init_schema(db_path)
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        cur.executemany('INSERT OR REPLACE INTO prices (length_dm, load_code, price) VALUES (?,?,?)', rows)
        conn.commit()
        return len(rows)
    finally:
        conn.close()


def get_price(length_m: float, load_code: float | int = 8, db_path: str = DEFAULT_DB) -> Optional[float]:
    """
    Получает цену плиты из базы данных.
    
    ВАЖНО: Для нагрузки 12.5 использует цену 12п (math.floor).
    12.5 кПа считается по цене 12 кПа, но отображается как 12,5п.
    """
    import math

    init_schema(db_path)
    length_dm = length_m_to_price_length_dm(length_m)
    
    # КЛЮЧЕВОЕ ИЗМЕНЕНИЕ: округляем нагрузку вниз (12.5 → 12)
    # В базе цен нет 12.5, используем цену 12
    load_code_for_db = int(math.floor(load_code)) if isinstance(load_code, (int, float)) else 8
    
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('SELECT price FROM prices WHERE length_dm=? AND load_code=?', (length_dm, load_code_for_db))
        row = cur.fetchone()
        if row:
            result = float(row[0])
        else:
            cur.execute('SELECT price FROM prices WHERE ABS(length_dm-?)<=1 AND load_code=? ORDER BY ABS(length_dm-?) LIMIT 1', (length_dm, load_code_for_db, length_dm))
            row = cur.fetchone()
            result = float(row[0]) if row else None
        return result
    finally:
        conn.close()


