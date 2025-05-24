import sqlite3

def create_tables():
    conn = sqlite3.connect('database/converter.db')
    cursor = conn.cursor()

    # Таблица Excel файлов (конвертированные)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS xlsx_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT UNIQUE,
            added_at TEXT
        )
    ''')

    # Таблица сообщений, связанных с XLSX
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            xlsx_file_id INTEGER,
            serial_number INTEGER,
            time TEXT,
            message_type TEXT,
            task_number INTEGER,
            data_length INTEGER,
            data_blob TEXT,
            developer_note TEXT,
            FOREIGN KEY (xlsx_file_id) REFERENCES xlsx_files(id) ON DELETE CASCADE
        )
    ''')

    conn.commit()
    conn.close()
    print("База данных и таблицы успешно созданы.")
