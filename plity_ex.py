# Например:
import pandas as pd
import sqlite3

conn = sqlite3.connect('pb.db')
df_ex = pd.read_sql('SELECT * FROM plity_ex', conn)
conn.close()

# Сортировка как выше!
df_ex.sort_values('неделя формовки', ascending=True).to_excel('экземпляры_по_неделям_возр.xlsx', index=False)
df_ex.sort_values('неделя формовки', ascending=False).to_excel('экземпляры_по_неделям_убыв.xlsx', index=False)

