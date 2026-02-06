#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json, sqlite3
from collections import defaultdict

log_file = r'c:\Users\Роман\Desktop\Шишов\.cursor\debug.log'
db_path = r'c:\Users\Роман\Desktop\Шишов\plita.db'

# 1. Из логов: что НЕ найдено
not_found = []
found = []
with open(log_file, 'r', encoding='utf-8') as f:
    for line in f:
        try:
            log = json.loads(line)
            if log.get('runId') != 'post-fix':
                continue
            if 'width-check' in log.get('message', ''):
                not_found.append(log['data'])
            elif log.get('location', '').endswith(':success'):
                found.append(log['data'])
        except:
            pass

print('='*70)
print('АНАЛИЗ ПОСЛЕ ФИКСА')
print('='*70)

# Подсчет успешных списаний
total_deducted = sum(f.get('deduct', 0) for f in found)
total_not_found = sum(nf.get('qty', 0) for nf in not_found)

print(f'\nУспешно списано: {len(found)} операций, SUM(deduct)={total_deducted}')
print(f'НЕ НАЙДЕНО: {len(not_found)} операций, SUM(qty)={total_not_found}')

# 2. Группируем НЕ НАЙДЕННЫЕ по имени
nf_by_name = defaultdict(int)
nf_details = defaultdict(list)
for nf in not_found:
    name = nf.get('plate_name', '?')
    qty = nf.get('qty', 0)
    nf_by_name[name] += qty
    nf_details[name].append({
        'kp_id': nf.get('kp_id'),
        'width': nf.get('width_m'),
        'length': nf.get('length_m'),
        'qty': qty
    })

print(f'\nНЕ НАЙДЕННЫЕ ПЛИТЫ (по имени):')
for name, qty in sorted(nf_by_name.items(), key=lambda x: -x[1]):
    details = nf_details[name]
    print(f'  {name}: {qty} шт')
    for d in details[:3]:
        print(f'    kp_id={d["kp_id"]}, length={d["length"]}, width={d["width"]}, qty={d["qty"]}')

# 3. Из БД: что осталось в kp_plates
conn = sqlite3.connect(db_path)
cur = conn.cursor()
remaining = cur.execute('SELECT kp_id, plate_name, length_m, width_m, load_class, qty FROM kp_plates WHERE qty > 0').fetchall()
total_remaining = sum(r[5] for r in remaining)

print(f'\nОСТАЛОСЬ В kp_plates: {total_remaining} шт ({len(remaining)} записей)')
for r in remaining:
    print(f'  kp_id={r[0]}, {r[1]}, length={r[2]}, width={r[3]}, load={r[4]}, qty={r[5]}')

# 4. completed_plates статистика
total_completed = cur.execute('SELECT COUNT(*), SUM(qty) FROM completed_plates').fetchone()
print(f'\nCOMPLETED_PLATES: {total_completed[0]} записей, SUM(qty)={total_completed[1]}')

conn.close()

# 5. Ключевой вопрос: 25,4-12-8п
print('\n' + '='*70)
print('АНАЛИЗ КЛЮЧЕВЫХ ПЛИТ:')
print('='*70)

for target in ['25,4-12-8', '43-12-8', '63,9-12-8']:
    target_found = [f for f in found if target in f.get('plate_name', '')]
    target_nf = [nf for nf in not_found if target in nf.get('plate_name', '')]
    found_qty = sum(f.get('deduct', 0) for f in target_found)
    nf_qty = sum(nf.get('qty', 0) for nf in target_nf)
    print(f'\n  {target}:')
    print(f'    Списано: {found_qty} шт ({len(target_found)} операций)')
    print(f'    НЕ найдено: {nf_qty} шт ({len(target_nf)} операций)')
    if target_nf:
        print(f'    Детали НЕ найденных:')
        for nf in target_nf:
            print(f'      kp_id={nf.get("kp_id")}, width={nf.get("width_m")}, qty={nf.get("qty")}')
            # Показываем rows_in_db
            rows = nf.get('rows_in_db_for_kp', [])
            same_length = [r for r in rows if abs(r.get('length_m', 0) - nf.get('length_m', 0)) < 0.02]
            if same_length:
                print(f'      В БД с такой же длиной: {same_length}')
            else:
                print(f'      В БД нет плит с длиной {nf.get("length_m")} для kp_id={nf.get("kp_id")}')
