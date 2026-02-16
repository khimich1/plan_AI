"""
Скрипт проверки: все ли плиты из БД попали в план.
Сравнивает kp_plates (qty>0) с плитами в JSON-плане.
"""
import json
import sqlite3
from pathlib import Path
from collections import defaultdict

PROJECT_ROOT = Path(__file__).parent.parent
DB_PATH = PROJECT_ROOT / "plita.db"
PLAN_PATH = PROJECT_ROOT / "bot" / "data" / "plans" / "plan_20260215_140537.json"


def collect_plan_plates(plan_path: Path) -> dict:
    """Собирает плиты из плана: (kp_id, plate_name, length, width) -> qty"""
    with open(plan_path, encoding="utf-8") as f:
        plan = json.load(f)
    
    result = defaultdict(int)
    for day_key, day_data in plan.get("days", {}).items():
        for track in day_data.get("tracks", []):
            for item in track.get("items", []):
                kp_id = item.get("kp_id")
                plate_name = item.get("plate_name") or item.get("label") or item.get("label_main")
                length = item.get("length")
                width = item.get("width") or item.get("main_w")
                if kp_id and plate_name:
                    key = (kp_id, plate_name, round(length, 2) if length else None, round(width, 3) if width else None)
                    result[key] += 1
    return dict(result)


def collect_db_plates(db_path: Path, kp_ids=None) -> dict:
    """Собирает плиты из БД: (kp_id, plate_name, length, width) -> qty"""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    
    if kp_ids:
        placeholders = ",".join("?" * len(kp_ids))
        cur.execute(f"""
            SELECT kp_id, plate_name, length_m, width_m, qty, status
            FROM kp_plates
            WHERE kp_id IN ({placeholders}) AND qty > 0
        """, kp_ids)
    else:
        cur.execute("""
            SELECT kp_id, plate_name, length_m, width_m, qty, status
            FROM kp_plates
            WHERE qty > 0
        """)
    
    result = defaultdict(int)
    for row in cur.fetchall():
        r = dict(row)
        key = (
            r["kp_id"],
            r["plate_name"],
            round(r["length_m"], 2) if r["length_m"] else None,
            round(r["width_m"], 3) if r["width_m"] else None,
        )
        result[key] += r["qty"]
    
    conn.close()
    return dict(result)


def main():
    plan_plates = collect_plan_plates(PLAN_PATH)
    kp_ids_in_plan = sorted(set(k[0] for k in plan_plates.keys()))
    
    print("=" * 60)
    print("PROVERKA: Vse li plity iz BD popali v plan?")
    print("=" * 60)
    print(f"\nКП в плане: {kp_ids_in_plan}")
    
    # Берём плиты из БД для этих КП (все с qty>0, без фильтра по status)
    db_plates_all = collect_db_plates(DB_PATH, kp_ids_in_plan)
    
    # Плиты со статусом "в производстве" (как при создании плана)
    # И "в плане" (уже добавлены в план - после создания плана они переходят в этот статус)
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    if kp_ids_in_plan:
        ph = ",".join("?" * len(kp_ids_in_plan))
        cur.execute(f"""
            SELECT kp_id, plate_name, length_m, width_m, qty, status
            FROM kp_plates
            WHERE kp_id IN ({ph}) AND qty > 0 AND status IN ('в производстве', 'в плане')
        """, kp_ids_in_plan)
    else:
        cur.execute("""
            SELECT kp_id, plate_name, length_m, width_m, qty, status
            FROM kp_plates
            WHERE qty > 0 AND status IN ('в производстве', 'в плане')
        """)
    
    db_plates_in_production = defaultdict(int)
    for row in cur.fetchall():
        kp_id, plate_name, length_m, width_m, qty, _status = row
        key = (
            kp_id,
            plate_name,
            round(length_m, 2) if length_m else None,
            round(width_m, 3) if width_m else None,
        )
        db_plates_in_production[key] += qty
    
    conn.close()
    db_plates_in_production = dict(db_plates_in_production)
    
    # Нормализуем ключи плана: width может быть в main_w (split), тогда в плане width = main_w
    # Упрощённое сравнение: (kp_id, plate_name, length) -> qty
    def normalize_plan():
        out = defaultdict(int)
        for (kp_id, plate_name, length, width), qty in plan_plates.items():
            key = (kp_id, plate_name, length)
            out[key] += qty
        return dict(out)
    
    def normalize_db(d):
        out = defaultdict(int)
        for (kp_id, plate_name, length, width), qty in d.items():
            key = (kp_id, plate_name, length)
            out[key] += qty
        return dict(out)
    
    plan_norm = normalize_plan()
    db_norm = normalize_db(db_plates_in_production)
    
    missing = []
    for key, qty_db in db_norm.items():
        qty_plan = plan_norm.get(key, 0)
        if qty_plan < qty_db:
            missing.append((*key, qty_db, qty_plan, qty_db - qty_plan))
    
    extra = []
    for key, qty_plan in plan_norm.items():
        qty_db = db_norm.get(key, 0)
        if qty_plan > qty_db:
            extra.append((*key, qty_db, qty_plan, qty_plan - qty_db))
    
    print(f"\n--- Плиты в БД (status IN ('в производстве','в плане'), qty>0) для КП {kp_ids_in_plan}: {sum(db_norm.values())} шт.")
    print(f"--- Плиты в плане: {sum(plan_norm.values())} шт.")
    
    if missing:
        print("\n[!] NE POPALI V PLAN (est v BD, no menshe ili net v plane):")
        for kp_id, plate_name, length, qty_db, qty_plan, diff in missing:
            print(f"   KP #{kp_id}: {plate_name} (dlina {length}m) - v BD: {qty_db}, v plane: {qty_plan}, ne hvataet: {diff}")
    else:
        print("\n[OK] Vse plity iz BD popali v plan po kolichestvu.")
    
    if extra:
        print("\n[!] V plane bolshe chem v BD (mozhet iz ostatkov):")
        for kp_id, plate_name, length, qty_db, qty_plan, diff in extra[:10]:
            print(f"   KP #{kp_id}: {plate_name} (dlina {length}m) - v BD: {qty_db}, v plane: {qty_plan}, lishnih: {diff}")
        if len(extra) > 10:
            print(f"   ... i esche {len(extra) - 10}")
    
    # Проверка: есть ли в БД плиты с другими статусами для этих КП?
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    ph = ",".join("?" * len(kp_ids_in_plan))
    cur.execute(f"""
        SELECT kp_id, plate_name, status, qty
        FROM kp_plates
        WHERE kp_id IN ({ph}) AND qty > 0 AND status != 'в производстве'
    """, kp_ids_in_plan)
    other_status = cur.fetchall()
    conn.close()
    
    if other_status:
        print("\n[INFO] Plity s drugim statusom - oni NE zagruzhayutsya v plan:")
        for row in other_status:
            print(f"   KP #{row[0]}: {row[1]} - status '{row[2]}', qty={row[3]}")


if __name__ == "__main__":
    main()
