import pandas as pd
import sqlite3

# Настрой здесь, если нужно другое имя файла/листа
filename = 'Plan-proizvodstva-PB.xlsx'
sheet = 'ПП.ПлитыПБ'

# Чтение и подготовка данных
df = pd.read_excel(filename, sheet_name=sheet, header=2)
df = df.iloc[3:]  # удаляем первые 3 строки (0, 1, 2)
df = df.iloc[:, :15]
df.columns = df.columns.str.strip()
df.columns = df.columns.str.lower()

needed_columns = [
    'номенклатура пб к производству',
    'длина плиты, м',
    'осталось изготовить, факт',
    'контрагент заказчик',
    'армирование по серии',
    'кол-во по заказу',
    'неделя формовки'
]
df = df[needed_columns]
df['осталось изготовить, факт'] = df['осталось изготовить, факт'].fillna(0)
df['осталось изготовить, факт'] = df['осталось изготовить, факт'].abs()
df = df[df['осталось изготовить, факт'] != 0]

conn = sqlite3.connect('pb.db')
cur = conn.cursor()
cur.execute('DROP TABLE IF EXISTS plity')
cur.execute('''
CREATE TABLE plity (
    "номенклатура пб к производству" TEXT,
    "длина плиты, м" REAL,
    "осталось изготовить, факт" INTEGER,
    "контрагент заказчик" TEXT,
    "армирование по серии" TEXT,
    "кол-во по заказу" INTEGER,
    "неделя формовки" TEXT
)
''')
conn.commit()

for idx, row in df.iterrows():
    cur.execute('''
        INSERT INTO plity (
            "номенклатура пб к производству",
            "длина плиты, м",
            "осталось изготовить, факт",
            "контрагент заказчик",
            "армирование по серии",
            "кол-во по заказу",
            "неделя формовки"
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (
        row['номенклатура пб к производству'],
        row['длина плиты, м'],
        int(row['осталось изготовить, факт']),
        row['контрагент заказчик'],
        row['армирование по серии'],
        int(row['кол-во по заказу']) if not pd.isna(row['кол-во по заказу']) else None,
        row['неделя формовки'] if not pd.isna(row['неделя формовки']) else None
    ))
conn.commit()

# --- Таблица экземпляров плит ---
df['осталось изготовить, факт'] = df['осталось изготовить, факт'].astype(int)
expanded = df.loc[df.index.repeat(df['осталось изготовить, факт'])].reset_index(drop=True)
expanded = expanded.drop(columns=['осталось изготовить, факт'])
expanded.to_sql('plity_ex', conn, if_exists='replace', index=False)

conn.close()
print('База данных pb.db успешно создана и заполнена!')
