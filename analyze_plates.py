import sqlite3

# Подключаемся к базе данных
conn = sqlite3.connect('pb.db')
cursor = conn.cursor()

# Все плиты из кода visualize_kz_plan.py с указанием ширины и правильных кодов нагрузки
plates_info = [
    (34, 12, 'ПБ 34-12-8п (3.39м)', 2, 1.2),  # PLATES_1_2 = [3.39]*2
    (66, 6, 'ПБ 66-6-8п (6.63м)', 4, 0.32),   # PLATES_0_32 = [6.63]*4 (используем нагрузку 6)
    (78, 6, 'ПБ 78-6-8п (7.83м)', 3, 0.32),   # PLATES_0_32 = [7.83]*3 (используем нагрузку 6)
    (56, 6, 'ПБ 56-6-8п (5.63м)', 5, 0.72),   # PLATES_0_72 = [5.63]*5 (используем нагрузку 6)
    (47, 6, 'ПБ 47-6-8п (4.65м)', 5, 0.70),   # PLATES_0_70 = [4.65]*5 (используем нагрузку 6)
    (68, 6, 'ПБ 68-6-8п (6.75м)', 2, 0.86),   # PLATES_0_86 = [6.75]*2 (используем нагрузку 6)
    (47, 6, 'ПБ 47-6-8п (4.65м)', 5, 0.86),   # PLATES_0_86 = [4.65]*5 (используем нагрузку 6)
]

print('=== ВСЕ ПЛИТЫ И ИХ СТОИМОСТЬ ===')
total_cost = 0
total_weight = 0
total_plates = 0

for length_dm, load_code, name, quantity, width_m in plates_info:
    cursor.execute('SELECT price FROM prices WHERE length_dm = ? AND load_code = ?', (length_dm, load_code))
    result = cursor.fetchone()
    
    if result:
        # Цена из БД - это цена за плиту шириной 1.2м
        price_per_unit_1_2m = result[0]
        
        # Корректируем цену пропорционально ширине
        width_factor = width_m / 1.2
        price_per_unit = price_per_unit_1_2m * width_factor
        
        total_price = price_per_unit * quantity
        
        # Примерный вес (кг/м² * площадь)
        area = length_dm * 10 * width_m  # длина * ширина в м²
        weight_per_unit = area * 250  # примерно 250 кг/м² для ПБ
        total_weight_for_plates = weight_per_unit * quantity
        
        total_cost += total_price
        total_weight += total_weight_for_plates
        total_plates += quantity
        
        print(f'{name} (ширина {width_m}м): {quantity} шт x {price_per_unit:,.0f} руб = {total_price:,.0f} руб')
        print(f'  (базовая цена за 1.2м: {price_per_unit_1_2m:,.0f} руб, коэффициент: {width_factor:.2f})')
    else:
        print(f'{name}: ЦЕНА НЕ НАЙДЕНА в базе данных')

print(f'\n=== ИТОГО ===')
print(f'Всего плит: {total_plates} шт')
print(f'Общая стоимость: {total_cost:,.0f} руб')
print(f'Общий вес: {total_weight:,.0f} кг')
print(f'Средняя цена за плиту: {total_cost/total_plates:,.0f} руб')

# Проверим количество резов
longitudinal_cuts = 4 + 3 + 5 + 5 + 2 + 5  # все плиты кроме PLATES_1_2
cuts_cost = longitudinal_cuts * 460  # 460 руб/пог.м за продольный рез

print(f'\n=== РЕЗЫ ===')
print(f'Продольных резов: {longitudinal_cuts} шт')
print(f'Стоимость резов: {cuts_cost:,.0f} руб')

print(f'\n=== ОБЩАЯ СТОИМОСТЬ ===')
print(f'Стоимость плит: {total_cost:,.0f} руб')
print(f'Стоимость резов: {cuts_cost:,.0f} руб')
print(f'ИТОГО: {total_cost + cuts_cost:,.0f} руб')

conn.close()
