from PyQt5.QtCore import QThread, pyqtSignal
import os, shutil, sqlite3, openpyxl
from Global import manual_widths
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


DB_PATH = "database/converter.db"
EXPORT_DIR = "database/export"
FIELD_MAPPING = {
    "Порядковый номер": "serial_number",
    "Время": "time",
    "Номер задачи": "task_number",
    "Тип диагностического сообщения": "message_type",
    "Длина бинарных данных": "data_length",
    "Бинарные данные": "data_blob",
    "Текстовое сообщение разработчику": "developer_note"
}

def get_all_files_from_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT id, file_name, added_at FROM xlsx_files")
    rows = cursor.fetchall()
    conn.close()
    return [{"id": row[0], "file_name": row[1], "added_at": row[2]} for row in rows]

def get_db_size_mb():
    if os.path.exists(DB_PATH):
        size_bytes = os.path.getsize(DB_PATH)
        return size_bytes / (1024 * 1024)
    return 0

def delete_selected_files(file_ids: list[int]):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    for file_id in file_ids:
        cursor.execute("DELETE FROM xlsx_files WHERE id = ?", (file_id,))
    conn.commit()
    conn.close()


def decode_koi8r_if_needed(value):
    if isinstance(value, bytes):
        try:
            return value.decode('koi8-r')
        except UnicodeDecodeError:
            return value.decode('utf-8', errors='replace')
    elif isinstance(value, str):
        try:
            return value.encode('cp1252').decode('koi8-r')
        except (UnicodeEncodeError, UnicodeDecodeError):
            return value
    else:
        return value

def decode_key(key):
    """
    Декодирует ключ (название столбца) из байтов или некорректной строки.
    Если ключ - str, пытаемся привести к нормальному виду, иначе возвращаем как есть.
    """
    if isinstance(key, bytes):
        try:
            return key.decode('koi8-r')
        except UnicodeDecodeError:
            return key.decode('utf-8', errors='replace')
    elif isinstance(key, str):
        try:
            return key.encode('cp1252').decode('koi8-r')
        except (UnicodeEncodeError, UnicodeDecodeError):
            return key
    else:
        return key



class ExportSelectedFilesThread(QThread):
    finished = pyqtSignal(str)

    def __init__(self, file_ids):
        super().__init__()
        self.file_ids = file_ids

    def run(self):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        for file_id in self.file_ids:
            cursor.execute("SELECT file_name FROM xlsx_files WHERE id = ?", (file_id,))
            file_name_row = cursor.fetchone()
            if not file_name_row:
                continue
            file_name = file_name_row[0]
            cursor.execute("SELECT * FROM messages WHERE xlsx_file_id = ?", (file_id,))
            messages = cursor.fetchall()
            if messages:
                wb = Workbook()
                ws = wb.active
                ws.title = "Messages"
                header = ["Название файла", "Порядковый номер", "Время", "Номер задачи", "Тип диагностическго сообщения", 
                          "Длина бинарных данных", "Бинарные данные", "Сообщение разработчику"]
                ws.append(header)
                for msg in messages:
                    ws.append(msg[2:9])  # Пропускаем id и file_id
                for i, width in enumerate(manual_widths, 1):
                    ws.column_dimensions[get_column_letter(i)].width = width
                if not os.path.exists(EXPORT_DIR):
                    os.makedirs(EXPORT_DIR)
                export_path = os.path.join(EXPORT_DIR, f"{file_name}_exported.xlsx")
                wb.save(export_path)
        conn.close()
        self.finished.emit("Выбранные файлы экспортированы.")

class ExportEntireDatabaseThread(QThread):
    finished = pyqtSignal(str)

    def run(self):
        if not os.path.exists(EXPORT_DIR):
            os.makedirs(EXPORT_DIR)
        export_path = os.path.join(EXPORT_DIR, "full_database_export.sqlite")
        shutil.copyfile("database/converter.db", export_path)
        self.finished.emit("Вся база данных экспортирована.")


def load_excel_messages(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    messages = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        msg = dict(zip(headers, row))
        messages.append(msg)
    return messages

def load_db_messages(file_id, db_connection):
    cursor = db_connection.cursor()
    cursor.execute("SELECT * FROM messages WHERE xlsx_file_id = ?", (file_id,))
    
    # Декодируем названия столбцов
    raw_columns = [desc[0] for desc in cursor.description]
    columns = [decode_key(c) for c in raw_columns]
    
    results = []
    for row in cursor.fetchall():
        msg = {}
        for i, val in enumerate(row):
            key = columns[i]
            msg[key] = decode_koi8r_if_needed(val)
        results.append(msg)
    return results

def compare_messages(excel_msgs, db_msgs):
    diffs = []
    stats = {
        "total_rows": 0,
        "diff_rows": 0,
        "matched_rows": 0,
        "field_diff_counts": {}
    }

    for i, (excel_msg, db_msg) in enumerate(zip(excel_msgs, db_msgs)):
        diff_fields = []
        for excel_field in FIELD_MAPPING:
            db_field = FIELD_MAPPING[excel_field]
            excel_value = str(excel_msg.get(excel_field, "")).strip()
            db_value = str(db_msg.get(db_field, "")).strip()
            if excel_value != db_value:
                diff_fields.append(excel_field)
                stats["field_diff_counts"][excel_field] = stats["field_diff_counts"].get(excel_field, 0) + 1

        if diff_fields:
            diffs.append({
                "row": i + 1,
                "diff_fields": diff_fields,
                "db_values": {FIELD_MAPPING[k]: db_msg.get(FIELD_MAPPING[k], "<нет>") for k in diff_fields},
                "excel_values": {k: excel_msg.get(k, "<нет>") for k in diff_fields}
            })

    stats["total_rows"] = len(excel_msgs)
    stats["diff_rows"] = len(diffs)
    stats["matched_rows"] = stats["total_rows"] - stats["diff_rows"]
    return diffs, stats
