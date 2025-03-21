import threading
from PyQt5.QtCore import pyqtSignal, QThread
from Convertation import parse_bin_file
import os
import pandas as pd, time

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