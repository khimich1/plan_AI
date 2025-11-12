import pandas as pd

# Читаем новую смету
df = pd.read_excel('Визуализация_Раскладки/Смета_Дорожка_1_20251027_1346.xlsx', sheet_name='Смета')

print('=== НОВАЯ СМЕТА С ИСПРАВЛЕННЫМИ ЦЕНАМИ ===')
print(df.to_string(index=False))

# Подсчитываем общую стоимость
total_cost = df['Сумма'].str.replace(' ', '').str.replace(',', '.').astype(float).sum()
print(f'\n=== ИТОГО ===')
print(f'Общая стоимость: {total_cost:,.0f} руб')
print(f'Количество позиций: {len(df)}')
