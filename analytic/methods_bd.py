from PyQt5.QtCore import QThread, pyqtSignal
import os, shutil, sqlite3
from Global import manual_widths
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

DB_PATH = "database/converter.db"
EXPORT_DIR = "database/export"

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
