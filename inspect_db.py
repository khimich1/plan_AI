import sqlite3
import json

def inspect_db():
    conn = sqlite3.connect('pb.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    # Get schema
    cursor.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='prays_plity';")
    schema = cursor.fetchone()[0]
    
    # Get columns
    cursor.execute("PRAGMA table_info(prays_plity)")
    columns = [dict(col) for col in cursor.fetchall()]
        
    # Get a few rows
    cursor.execute("SELECT * FROM prays_plity LIMIT 5")
    rows = [dict(row) for row in cursor.fetchall()]
    
    with open('inspect_db_out.json', 'w', encoding='utf-8') as f:
        json.dump({'schema': schema, 'columns': columns, 'rows': rows}, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    inspect_db()
