"""Скрипт для проверки справочника prays_plity в pb.db.

Запуск: python check_prays_plity.py
"""
import os
import sqlite3

# Путь к pb.db (как в kp_db.py)
_PB_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pb.db')

def main():
    if not os.path.exists(_PB_DB_PATH):
        print(f"Файл pb.db не найден: {_PB_DB_PATH}")
        return

    conn = sqlite3.connect(_PB_DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # Запрос 1: плиты 61,8 / 45 / 37,9
    print("=== Плиты 61,8 / 45 / 37,9 ===")
    cur.execute('''
        SELECT "Уникальный идентификатор (Номенклатура)", "Товар"
        FROM prays_plity
        WHERE "Товар" LIKE '%61,8%' OR "Товар" LIKE '%45%' OR "Товар" LIKE '%37,9%'
        ORDER BY "Товар"
    ''')
    rows = cur.fetchall()
    for r in rows:
        print(f"  {r['Товар']} -> {r['Уникальный идентификатор (Номенклатура)']}")
    if not rows:
        print("  (нет записей)")

    # Запрос 2: плиты с шириной 5, 7, 9 (500/700/900 мм)
    print("\n=== Плиты с шириной -5-, -7-, -9- (в марке) ===")
    cur.execute('''
        SELECT "Уникальный идентификатор (Номенклатура)", "Товар"
        FROM prays_plity
        WHERE "Товар" LIKE '%-5-%' OR "Товар" LIKE '%-7-%' OR "Товар" LIKE '%-9-%'
        ORDER BY "Товар"
    ''')
    rows = cur.fetchall()
    for r in rows:
        print(f"  {r['Товар']} -> {r['Уникальный идентификатор (Номенклатура)']}")
    if not rows:
        print("  (нет записей)")

    # Запрос 3: варианты 5,0 / 7,0 / 9,0
    print("\n=== Плиты с шириной -5,0-, -7,0-, -9,0- ===")
    cur.execute('''
        SELECT "Уникальный идентификатор (Номенклатура)", "Товар"
        FROM prays_plity
        WHERE "Товар" LIKE '%-5,0-%' OR "Товар" LIKE '%-7,0-%' OR "Товар" LIKE '%-9,0-%'
        ORDER BY "Товар"
    ''')
    rows = cur.fetchall()
    for r in rows:
        print(f"  {r['Товар']} -> {r['Уникальный идентификатор (Номенклатура)']}")
    if not rows:
        print("  (нет записей)")

    # Запрос 5: точные совпадения для 45-7 / 45,0-7
    print("\n=== Точный поиск 45-7 / 45,0-7 / 45-7,0 ===")
    for q in ["Плиты ПБ 45-7-6п", "Плиты ПБ 45-7,0-6п", "Плиты ПБ 45,0-7-6п", "Плиты ПБ 45,0-7,0-6п"]:
        cur.execute('SELECT "Товар" FROM prays_plity WHERE "Товар" = ?', (q,))
        r = cur.fetchone()
        print(f"  {q!r}: {'найдено' if r else 'нет'}")

    # Запрос 6: общее количество записей
    cur.execute("SELECT COUNT(*) FROM prays_plity")
    total = cur.fetchone()[0]
    print(f"\nВсего записей в prays_plity: {total}")

    conn.close()

if __name__ == "__main__":
    main()
