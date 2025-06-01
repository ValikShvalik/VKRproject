import threading
from PyQt5.QtCore import pyqtSignal, QThread
from Convertation import parse_bin_file
from Sort_by_number_task import filter_rows_by_task, create_sorted_workbook, gain_task_number
import pandas as pd, time, openpyxl, os, tempfile
from database.db_manager import get_unique_file_name, insert_file, insert_messages
from Sort_by_diag_type import sort_by_diag_type_message



class Core_process(QThread):
    progress_updated = pyqtSignal(int)
    process_completed = pyqtSignal(object)

    def __init__(self, bin_file, parent=None):
        super().__init__(parent)
        self.bin_file = bin_file

    def run(self):
        self.progress_updated.emit(25)  # Начинаем процесс
        wb = parse_bin_file(self.bin_file)  # Конвертация bin файла
        if wb:
            self.progress_updated.emit(100)  # Завершаем процесс
            self.process_completed.emit(wb)  # Отправляем рабочую книгу обратно
        else:
            self.progress_updated.emit(0)  # Ошибка
            self.process_completed.emit(None)


class LoadXlsxThread(QThread):
    data_loaded = pyqtSignal(pd.DataFrame)
    progress_update = pyqtSignal(int)

    def __init__(self, wb):
        super().__init__()
        self.wb = wb

    
    def run(self):
        # Теперь мы будем использовать openpyxl для обработки данных
        try:
            self.progress_update.emit(25)
            all_data = []
            # Проходим по всем листам в Workbook
            for sheet in self.wb.worksheets:
                sheet_data = []
                for row in sheet.iter_rows(min_row=2, values_only=True):  # Пропускаем заголовок
                    sheet_data.append(row)
                
                # Преобразуем в DataFrame
                df = pd.DataFrame(sheet_data, columns=["Порядковый номер", "Время", "Номер задачи", 
                                                       "Тип диагностического сообщения", "Длина бинарных данных", 
                                                       "Бинарные данные", "Текстовое сообщение разработчику"])
                all_data.append(df)

            self.progress_update.emit(80)
            # Объединяем все данные в один DataFrame
            final_df = pd.concat(all_data, ignore_index=True)
            preview_df = final_df.head(200)
            self.progress_update.emit(95) 
            self.data_loaded.emit(preview_df)
            self.progress_update.emit(100)

        except Exception as e:
            self.progress_update.emit(0)  # Ошибка
            print(f"Ошибка загрузки данных: {e}")
            self.data_loaded.emit(pd.DataFrame())  # Пустая таблица при ошибке


class SaveFileThread(QThread):
    progress = pyqtSignal(int)
    file_saved = pyqtSignal(str)
    
    def __init__(self, workbook, save_path):
        super().__init__()
        self.workbook = workbook
        self.save_path = save_path
    
    def run(self):
        try:
            total_steps = 100  
            for i in range(total_steps):
                time.sleep(0.05)  
                self.progress.emit(int((i + 1) * 89 / total_steps)) 
            self.workbook.save(self.save_path)
            self.file_saved.emit(self.save_path)
            self.progress.emit(100) 
        except Exception as e:
            self.file_saved.emit(f"Error: {str(e)}")  


class TaskSearchThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(list)

    def __init__(self, xlsx_file):
        super().__init__()
        self.xlsx_file = xlsx_file

    def run(self):
        self.progress.emit(10)  # Начальный прогресс

        try:
            task_numbers = gain_task_number(self.xlsx_file)  # Запуск функции поиска номеров задач
        except Exception as e:
            task_numbers = []
            print(f"Ошибка при поиске номеров задач: {e}")

        self.progress.emit(100)  # Завершение
        self.finished.emit(task_numbers)

class SortMessageSearchThread(QThread):
    progress = pyqtSignal(int)
    search_done = pyqtSignal(list)

    def __init__(self, xlsx_file):
        super().__init__()
        self.xlsx_file = xlsx_file

    def run(self):
        self.progress.emit(17)
        all_unique_types = self.get_unique_diag_type(self.xlsx_file)
        self.search_done.emit(all_unique_types)

    def get_unique_diag_type(self, xlsx_file):
        import pandas as pd

        df = pd.read_excel(xlsx_file, sheet_name=None)
        self.progress.emit(25)
        all_unique_types = []

        for sheet_name, sheet_data in df.items():
            if isinstance(sheet_data, pd.DataFrame) and sheet_data.shape[1] > 3:
                # Получаем уникальные типы из 4-го столбца (индекса 3)
                unique_types = sheet_data.iloc[:, 3].dropna().unique().tolist()
                self.progress.emit(65)
                all_unique_types.extend(unique_types)

        all_unique_types = list(set(all_unique_types))  # Убираем дубликаты
        self.progress.emit(100)
        return all_unique_types

class SortMessageSortingThread(QThread):
    progress = pyqtSignal(int)
    sorting_done = pyqtSignal(object)

    def __init__(self, xlsx_file, selected_types):
        super().__init__()
        self.xlsx_file = xlsx_file
        self.selected_types = selected_types

    def run(self):
        self.progress.emit(25)
        sorted_workbook = sort_by_diag_type_message(self.xlsx_file, self.selected_types)
        self.progress.emit(54)
        self.sorting_done.emit(sorted_workbook)
        self.progress.emit(100)

class SortTaskThread(QThread):
    progress = pyqtSignal(int)  # Сигнал для обновления прогресса
    sorting_done = pyqtSignal(object)  # Сигнал для передачи отсортированного workbook

    def __init__(self, xlsx_file, selected_tasks):
        super().__init__()
        self.xlsx_file = xlsx_file
        self.selected_tasks = selected_tasks

    def run(self):
        self.progress.emit(10)
        new_header, data_by_task = filter_rows_by_task(self.xlsx_file, 3, self.selected_tasks) 
        self.progress.emit(25)
        sorted_workbook = create_sorted_workbook(new_header, data_by_task)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
            sorted_workbook.save(temp_file.name)

        self.progress.emit(80) 
        self.progress.emit(100)
        self.sorting_done.emit(sorted_workbook)

class LoadDoneXlsxFile(QThread):
    progress = pyqtSignal(int)
    loading_done = pyqtSignal(object)

    def __init__(self, xlsx_file, file_path):
        super().__init__()
        self.xlsx_file = xlsx_file
        self.file_path = file_path

    def run(self):
        self.progress.emit(25)
        workbook = openpyxl.load_workbook(self.file_path)
        total_steps = 75 
        time.sleep(0.05)  
        self.progress.emit(int((25 + 1) * 75 / total_steps)) 
        self.progress.emit(100)
        self.loading_done.emit(workbook)


class LoadReadyXlsx(QThread):
    progress_update = pyqtSignal(int)
    file_loaded = pyqtSignal(object)
    Ppath = pyqtSignal(object)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        self.progress_update.emit(63)

        try:
            # Загружаем workbook
            workbook = openpyxl.load_workbook(self.file_path)
            base_name = os.path.basename(self.file_path)

            # Получаем уникальное имя и записываем в БД
            unique_name = get_unique_file_name(base_name)
            file_id = insert_file(unique_name)

            # Чтение всех данных
            messages = []
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    messages.append({
                        'serial_number': row[0],
                        'time': row[1],
                        'task_number': row[2],
                        'message_type': row[3],
                        'data_length': row[4],
                        'data_blob': row[5],
                        'developer_note': row[6]
                    })

            # Запись сообщений в базу
            insert_messages(file_id, messages)
            print(f"Добавляем сообщения: {len(messages)} штук")


            self.progress_update.emit(100)
            self.file_loaded.emit(workbook)
            self.Ppath.emit(self.file_path)

        except Exception as e:
            print(f"Ошибка при загрузке XLSX: {e}")
            self.progress_update.emit(0)