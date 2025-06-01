from PyQt5.QtWidgets import (QApplication, QLabel, QWidget, QFileDialog, QProgressBar, QTableWidget, QTableWidgetItem,
                             QHBoxLayout, QVBoxLayout, QPushButton, QListWidget, QListWidgetItem, QDialog, QGroupBox, QMessageBox, QMainWindow)
import pandas as pd
from PyQt5.QtCore import pyqtSignal
import sys, os, openpyxl
from Core_process import SaveFileThread, Core_process, LoadXlsxThread, LoadReadyXlsx, TaskSearchThread, SortMessageSearchThread, SortTaskThread, SortMessageSortingThread
from Sort_by_number_task import gain_task_number
from database.init_db import create_tables
from analytic.menu_bar import AppMenuBar, ExportDialog, DeleteDialog, SolveDialog, FilterSearcherDialog, AnalyticsLauncher


class fileConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.xlsx_file = None
        self.bin_file = None
        self.core_process = None
        self.load_thread = None
        self.save_thread = None

        self.initUI()
        create_tables()

        # Добавляем меню
        self.menu_bar = AppMenuBar(self)
        self.setMenuBar(self.menu_bar)
        self.menu_bar.request_show_files.connect(lambda: self.open_export_dialog(mode="export"))
        self.menu_bar.request_delete_file.connect(lambda: self.open_export_dialog(mode="delete"))
        self.menu_bar.request_compare_excel.connect(lambda: self.open_export_dialog(mode="solve"))
        self.menu_bar.request_filter_search.connect(lambda: self.open_export_dialog(mode="filter"))
        self.menu_bar.request_run_analytics.connect(lambda: self.open_export_dialog(mode="analytic"))

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout()
        top_layout = QHBoxLayout()

        # Левая секция
        left_layout = QVBoxLayout()
        self.file_label = QLabel("Вставьте BIN или XLSX file")
        self.bin_select_file = QPushButton("Вставьте BIN или XLSX file")
        self.bin_select_file.clicked.connect(self.select_bin_file)

        left_layout.addWidget(self.file_label)
        left_layout.addWidget(self.bin_select_file)

        self.btn_process = QPushButton("Обработать")
        self.btn_process.setEnabled(False)
        self.btn_process.clicked.connect(self.process_bin_file)
        left_layout.addWidget(self.btn_process)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        left_layout.addWidget(self.progress_bar)

        self.btn_download = QPushButton("Скачать xlsx file")
        self.btn_download.setEnabled(False)
        self.btn_download.clicked.connect(self.download_xlsx)
        left_layout.addWidget(self.btn_download)

        self.btn_download_sorted = QPushButton("Скачать отсортированный файл")
        self.btn_download_sorted.setEnabled(False)
        self.btn_download_sorted.clicked.connect(self.download_sorted_xlsx)
        left_layout.addWidget(self.btn_download_sorted)

        # Правая секция (сортировка)
        sort_group = QGroupBox("Сортировка")
        sort_layout = QVBoxLayout()

        self.btn_sort_task = QPushButton("По номеру задач")
        self.btn_sort_type = QPushButton("По типу сообщения")
        self.btn_sort_task.clicked.connect(self.open_sort_task_window)
        self.btn_sort_type.clicked.connect(self.open_sort_message_window)
        sort_layout.addWidget(self.btn_sort_task)
        sort_layout.addWidget(self.btn_sort_type)

        sort_group.setLayout(sort_layout)

        top_layout.addLayout(left_layout)
        top_layout.addStretch()
        top_layout.addWidget(sort_group)

        main_layout.addLayout(top_layout)

        # Таблица
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Порядковый номер", "Время", "Номер задачи",
            "Тип диагностического сообщения", "Длина бинарных данных",
            "Бинарные данные", "Текстовое сообщение разработчику"
        ])

        self.table.horizontalHeader().setStretchLastSection(True)
        for i in range(self.table.columnCount()):
            self.table.horizontalHeader().setSectionResizeMode(i, 1)
        main_layout.addWidget(self.table)

        central_widget.setLayout(main_layout)
        self.setGeometry(100, 100, 800, 600)
        self.setWindowTitle("Converter")

    def open_export_dialog(self, mode="export"):
        if (mode == "export"):
            dialog = ExportDialog(self)
        elif (mode == "delete"):
            dialog = DeleteDialog(self)
        elif (mode == "solve"):
            dialog = SolveDialog(self)
        elif (mode == "filter"):
            dialog = FilterSearcherDialog(self)
        elif (mode == "analytic"):
            if not hasattr(self, "analytics_launcher"):
                self.analytics_launcher = AnalyticsLauncher(self)
            self.analytics_launcher.open_analytics()
            return
        else: 
            QMessageBox.warning(self, "Неизвестный режим")
            return
        dialog.exec_()
    
    def update_progress(self, progress):
        self.progress_bar.setValue(progress)

    def process_bin_file(self):
        if not self.bin_file:
            self.file_label.setText("Ошибка: выберите файл перед обработкой")
            return

        self.progress_bar.setValue(25)

        # Создаем и запускаем поток
        self.core_process = Core_process(self.bin_file)
        self.core_process.progress_updated.connect(self.update_progress)
        self.core_process.process_completed.connect(self.on_process_completed)
        self.core_process.start()
        self.btn_process.setEnabled(False)

    def on_process_completed(self, wb):
        if wb:
            self.xlsx_file = os.path.join(os.getcwd(), "converted_file.xlsx")
            self.btn_download.setEnabled(True)
            self.process_workbook = wb
            self.load_xlsx_preview(wb)
        else:
            self.file_label.setText("Ошибка при конвертации BIN файла")


    def select_bin_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите файл", "", "Все файлы (*);;BIN файлы (*.bin);;XLSX файлы (*.xlsx)", options=options)
        if file_path.endswith(".bin"):
            self.file_label.setText(f"Выбран файл: {file_path}")
            self.bin_file = file_path
            self.btn_process.setEnabled(True)  # Включаем кнопку "Обработать" после выбора BIN файла
        elif file_path.endswith(".xlsx"):
            self.file_label.setText(f"Выбран файл: {file_path}")
            self.progress_bar.setValue(25)
            self.ready_data = LoadReadyXlsx(file_path)
            self.ready_data.progress_update.connect(self.update_progress)
            self.ready_data.file_loaded.connect(self.load_xlsx_preview)
            self.ready_data.Ppath.connect(self.handle_file_loaded)
            self.ready_data.start()
        elif os.path.splitext(file_path)[1]:
            file_path +=".bin"
            self.file_label.setText(f"Выбран файл: {file_path}")
            self.bin_file = file_path
            self.btn_process.setEnabled(True) 

        self.btn_process.setEnabled(False)

    def handle_file_loaded(self, path):
        self.xlsx_file = path
        self.file_label.setText(f"Выбран файл: {self.xlsx_file}") 

    def load_xlsx_preview(self, wb):
        self.load_thread = LoadXlsxThread(wb)
        self.load_thread.data_loaded.connect(self.update_table)  # Подключаем сигнал к методу обновления таблицы
        self.load_thread.progress_update.connect(self.update_progress)  # Подключаем сигнал для обновления прогресса
        self.load_thread.start()

    def update_table(self, df):
        if df.empty:
            self.file_label.setText("Ошибка при загрузке данных")
            return
        
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))

        for row in range(len(df)):
            for col in range(len(df.columns)):
                self.table.setItem(row, col, QTableWidgetItem(str(df.iloc[row, col])))


    def download_xlsx(self):
        if not self.xlsx_file:
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить файл", "converted_file.xlsx", "XLSX Files (*.xlsx)")
        if save_path:
            self.save_thread = SaveFileThread(self.process_workbook, save_path)      
            self.save_thread.progress.connect(self.update_progress)
            self.save_thread.start()

            self.btn_download.setEnabled(False)  
            self.progress_bar.setValue(0)  
            self.progress_bar.setVisible(True) 

        
    def open_sort_message_window(self):
        if not self.xlsx_file:
            self.message_error = QMessageBox.warning(self, "Ошибка", "Не найден xlsx файл")
            return

        # Проверяем, не открыто ли уже окно сортировки
        if hasattr(self, 'sort_window') and self.sort_window.isVisible():
            return  

        # Создаем окно сортировки
        self.sort_window = SortByDiagMessageType(self, self.xlsx_file)
        self.sort_window.sorting_aplied.connect(self.start_diag_type_sorting)
        self.sort_window.show()

    def start_diag_type_sorting(self, selected_types):
        self.progress_bar.setValue(1)
        self.sort_thread = SortMessageSortingThread(self.xlsx_file, selected_types)
        self.sort_thread.progress.connect(self.progress_bar.setValue)
        self.sort_thread.sorting_done.connect(self.apply_diag_type_sorting)
        self.sort_thread.start()

    def apply_diag_type_sorting(self, sorted_workbok):
        self.sorted_workbook  = sorted_workbok
        self.btn_download_sorted.setEnabled(True)

    def open_sort_task_window(self):
        if not self.xlsx_file:
            QMessageBox.warning(self, "Ошибка", "Не найден xlsx файл")
            return

        if hasattr(self, 'sort_task_window') and self.sort_task_window.isVisible():
            return

        self.sort_task_window = SortByTaskNumber(self, self.xlsx_file)
        self.sort_task_window.sorting_applied.connect(self.start_task_number_sorting)
        self.sort_task_window.show()

    def start_task_number_sorting(self, selected_tasks):
        self.progress_bar.setValue(0)  # Инициализация прогресс-бара
        self.sort_thread = SortTaskThread(self.xlsx_file, selected_tasks)
        self.sort_thread.progress.connect(self.progress_bar.setValue)
        self.sort_thread.sorting_done.connect(self.apply_task_number_sorting)
        self.sort_thread.start()


    def apply_task_number_sorting(self, sorted_workbook):
        self.sorted_workbook = sorted_workbook
        self.btn_download_sorted.setEnabled(True)


    def download_sorted_xlsx(self):
        if not self.sorted_workbook:
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить отсортированный файл", "Sorted.xlsx", "XLSX Files (*.xlsx)")
        if save_path:
            self.save_thread = SaveFileThread(self.sorted_workbook, save_path)
            self.save_thread.progress.connect(self.update_progress)
            self.save_thread.start()

            self.btn_download_sorted.setEnabled(False)  
            self.progress_bar.setValue(0)  
            self.progress_bar.setVisible(True) 

            
class SortByDiagMessageType(QDialog):
    sorting_aplied = pyqtSignal(list)  # Сигнал для применения сортировки

    def __init__(self, parent=None, xlsx_file=None):
        super().__init__(parent)
        self.setWindowTitle("Сортировка по типу сообщений")
        self.setGeometry(200, 200, 500, 380)
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)

        self.xlsx_file = xlsx_file  # Сохраняем путь к файлу Excel

        main_layout = QVBoxLayout()
        lists_layout = QHBoxLayout()

        left_layout = QVBoxLayout()
        self.label_all = QLabel("Типы сообщений")
        left_layout.addWidget(self.label_all)

        self.list_all_types = QListWidget()
        self.list_all_types.setSelectionMode(QListWidget.NoSelection)
        left_layout.addWidget(self.list_all_types)

        lists_layout.addLayout(left_layout)

        right_layout = QVBoxLayout()
        self.label_select = QLabel("Выберите тип/типы сообщений")
        right_layout.addWidget(self.label_select)

        self.list_select_types = QListWidget()
        self.list_select_types.setSelectionMode(QListWidget.MultiSelection)
        right_layout.addWidget(self.list_select_types)

        lists_layout.addLayout(right_layout)
        main_layout.addLayout(lists_layout)

        # Прогресс-бар для поиска
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setVisible(False)  # Скрыт по умолчанию
        main_layout.addWidget(self.progress_bar)

        # Кнопка поиска типов сообщений
        self.btn_search = QPushButton("Поиск типов сообщений")
        self.btn_search.clicked.connect(self.start_message_type_search)
        main_layout.addWidget(self.btn_search)

        # Кнопка выполнения сортировки
        self.btn_sort = QPushButton("Выполнить сортировку")
        self.btn_sort.setEnabled(False)
        self.btn_sort.clicked.connect(self.apply_sorting)
        main_layout.addWidget(self.btn_sort)
        main_layout.addWidget(self.progress_bar)

        self.setLayout(main_layout)

        self.list_select_types.itemSelectionChanged.connect(self.check_selection)

    def check_selection(self):
        selected_items = self.list_select_types.selectedItems()
        if selected_items:
            self.btn_sort.setEnabled(True)
        else:
            self.btn_sort.setEnabled(False)

    def start_message_type_search(self):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.sort_thread = SortMessageSearchThread(self.xlsx_file)
        self.sort_thread.progress.connect(self.progress_bar.setValue)
        self.sort_thread.search_done.connect(self.on_search_finished)
        self.sort_thread.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def on_search_finished(self, unique_types):
        self.list_all_types.clear()

        if unique_types:
            for msg_type in sorted(unique_types):
                if msg_type == 255:
                    continue
                self.list_all_types.addItem(QListWidgetItem(str(msg_type)))
                self.list_select_types.addItem(QListWidgetItem(str(msg_type)))

        self.btn_search.setEnabled(True)
        self.progress_bar.setVisible(False) 

    def apply_sorting(self):
        selected_types = [int(item.text()) for item in self.list_select_types.selectedItems()]
        if selected_types:
            self.sorting_aplied.emit(selected_types)
        self.close()


class SortByTaskNumber(QDialog):
    sorting_applied = pyqtSignal(list)

    def __init__(self, parent=None, xlsx_file=None):
        super().__init__(parent)
        self.xlsx_file = xlsx_file
        self.setWindowTitle("Сортировка по номеру задач")
        self.setGeometry(200, 200, 500, 350)

        main_layout = QVBoxLayout()
        lists_layout = QHBoxLayout()

        # Прогресс-бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # Кнопка поиска номеров задач
        self.btn_search = QPushButton("Поиск номеров задач")
        self.btn_search.clicked.connect(self.start_task_search)
        main_layout.addWidget(self.btn_search)

        # Левая колонка
        left_layout = QVBoxLayout()
        self.label_all = QLabel("Номера задач")
        left_layout.addWidget(self.label_all)

        self.list_all_tasks = QListWidget()
        left_layout.addWidget(self.list_all_tasks)

        lists_layout.addLayout(left_layout)

        # Правая колонка
        right_layout = QVBoxLayout()
        self.label_select = QLabel("Выберите номер/номера задач")
        right_layout.addWidget(self.label_select)

        self.list_select_tasks = QListWidget()
        self.list_select_tasks.setSelectionMode(QListWidget.MultiSelection)
        right_layout.addWidget(self.list_select_tasks)

        lists_layout.addLayout(right_layout)
        main_layout.addLayout(lists_layout)

        # Кнопка сортировки
        self.btn_sort = QPushButton("Выполнить сортировку")
        self.btn_sort.setEnabled(False)
        self.btn_sort.clicked.connect(self.apply_sorting)
        main_layout.addWidget(self.btn_sort)

        self.setLayout(main_layout)

        self.list_select_tasks.itemSelectionChanged.connect(self.check_selection)

    def start_task_search(self):
        self.btn_search.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.task_search_thread = TaskSearchThread(self.xlsx_file)
        self.task_search_thread.progress.connect(self.progress_bar.setValue)
        self.task_search_thread.finished.connect(self.populate_tasks)
        self.task_search_thread.start()

    def populate_tasks(self, task_numbers):
        self.list_all_tasks.clear()
        self.list_select_tasks.clear()

        if not task_numbers:
            QMessageBox.warning(self, "Ошибка", "Не найдены номера задач")
            self.btn_search.setEnabled(True)
            self.progress_bar.setVisible(False)
            return

        for task in task_numbers:
            self.list_all_tasks.addItem(QListWidgetItem(str(task)))
            self.list_select_tasks.addItem(QListWidgetItem(str(task)))

        self.btn_search.setEnabled(True)
        self.progress_bar.setVisible(False)

    def check_selection(self):
        selected_items = self.list_select_tasks.selectedItems()
        self.btn_sort.setEnabled(bool(selected_items))

    def apply_sorting(self):
        selected_tasks = [item.text() for item in self.list_select_tasks.selectedItems()]
        if selected_tasks:
            self.sorting_applied.emit(selected_tasks)
        self.close()


app = QApplication(sys.argv)
ex = fileConverterApp()
ex.show()
sys.exit(app.exec_())