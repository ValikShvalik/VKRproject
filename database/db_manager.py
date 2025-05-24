import sqlite3, os
from datetime import datetime
from Global import manual_widths

DB_PATH = 'database/converter.db'


def get_connection():
    return sqlite3.connect(DB_PATH)

def get_all_file_names():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT file_name FROM xlsx_files ORDER BY added_at DESC")
    rows = cursor.fetchall()
    conn.close()
    return [row[0] for row in rows]

def get_unique_file_name(base_name):
    conn = get_connection()
    cursor = conn.cursor()

    name, ext = os.path.splitext(base_name)
    i = 1
    unique_name = base_name

    cursor.execute("SELECT file_name FROM xlsx_files WHERE file_name = ?", (unique_name,))
    while cursor.fetchone():
        unique_name = f"{name}_{i}{ext}"
        cursor.execute("SELECT file_name FROM xlsx_files WHERE file_name = ?", (unique_name,))
        i += 1

    conn.close()
    return unique_name

def file_exists(file_name):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM xlsx_files WHERE file_name = ?", (file_name,))
    result = cursor.fetchone()
    conn.close()
    return result[0] if result else None

def insert_file(file_name):
    conn = get_connection()
    cursor = conn.cursor()
    now = datetime.now().isoformat()
    cursor.execute("INSERT INTO xlsx_files (file_name, added_at) VALUES (?, ?)", (file_name, now))
    file_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return file_id

def insert_messages(file_id, messages: list):
    conn = get_connection()
    cursor = conn.cursor()
    for msg in messages:
        cursor.execute('''
            INSERT INTO messages (
                xlsx_file_id, serial_number, time,
                message_type, task_number, data_length, data_blob, developer_note
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            file_id,
            msg['serial_number'],
            msg['time'],
            msg['message_type'],
            msg['task_number'],
            msg['data_length'],
            msg['data_blob'],
            msg.get('developer_note', '')
        ))
    conn.commit()
    conn.close()
