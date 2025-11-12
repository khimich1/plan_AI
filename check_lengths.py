import sqlite3

conn = sqlite3.connect('pb.db')
cursor = conn.cursor()

print('=== ДОСТУПНЫЕ ДЛИНЫ В БД ===')
cursor.execute('SELECT DISTINCT length_dm FROM prices ORDER BY length_dm')
lengths = cursor.fetchall()
for row in lengths:
    print(f'{row[0]} дм ({row[0]/10:.1f}м)')

print('\n=== НУЖНЫЕ ДЛИНЫ ===')
needed_lengths = [34, 66, 78, 56, 47, 68]
for length_dm in needed_lengths:
    cursor.execute('SELECT COUNT(*) FROM prices WHERE length_dm = ?', (length_dm,))
    count = cursor.fetchone()[0]
    status = 'ЕСТЬ' if count > 0 else 'НЕТ'
    print(f'{length_dm} дм ({length_dm/10:.1f}м): {status}')

conn.close()

















