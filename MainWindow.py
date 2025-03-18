from PyQt5.QtWidgets import (QApplication, QLabel, QWidget, QFileDialog, QProgressBar, QTableWidget, QTableWidgetItem,
                             QHBoxLayout, QVBoxLayout, QPushButton, QListWidget, QListWidgetItem, QDialog)
import pandas as pd
from PyQt5.QtCore import pyqtSignal
from Convertation import parse_bin_file
from Sort_by_diag_type import sort_by_diag_type_message
from Global import sorted_by_diag_type_file
import sys
import os


class fileConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.xlsx_file = None
        self.bin_file = None

    def initUI(self):
        main_layout = QVBoxLayout()
        left_layout = QVBoxLayout()
        right_layout = QVBoxLayout()

        self.file_label = QLabel("Вставьте BIN file")
        self.bin_select_file = QPushButton("Вставьте BIN file")
        self.bin_select_file.clicked.connect(self.select_bin_file)
        left_layout.addWidget(self.bin_select_file)

        self.btn_process = QPushButton("Обработать")
        self.btn_process.setEnabled(False)  # Кнопка "Обработать" по умолчанию неактивна
        self.btn_process.clicked.connect(self.process_bin_file)
        left_layout.addWidget(self.btn_process)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        left_layout.addWidget(self.progress_bar)

        self.btn_download = QPushButton("Скачать xlsx file")
        self.btn_download.clicked.connect(self.download_xlsx)
        self.btn_download.setEnabled(False)  # Кнопка "Скачать" неактивна по умолчанию
        left_layout.addWidget(self.btn_download)

        self.btn_download_sorted = QPushButton("Скачать отсортированный файл")
        self.btn_download_sorted.clicked.connect(self.download_sorted_xlsx)
        self.btn_download_sorted.setEnabled(False)  # Кнопка для отсортированного файла неактивна
        left_layout.addWidget(self.btn_download_sorted)

        self.sort_label = QLabel("Сортировка")
        right_layout.addWidget(self.sort_label)
        self.btn_sort_task = QPushButton("По номеру задач")
        self.btn_sort_type = QPushButton("По типу сообщения")
        self.btn_sort_type.clicked.connect(self.open_sort_message_window)
        right_layout.addWidget(self.btn_sort_task)
        right_layout.addWidget(self.btn_sort_type)

        top_layout = QHBoxLayout()
        top_layout.addLayout(left_layout)
        top_layout.addStretch()
        top_layout.addLayout(right_layout)

        main_layout.addLayout(top_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["Порядковый номер", "Время", "Номер задачи", "Тип диагностического сообщения",
                                              "Длина бинарных данных", "Бинарные данные", "Текстовое сообщение разработчику"])
        self.table.setRowCount(200)

        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(0, 1)
        for i in range(self.table.columnCount()):
            self.table.horizontalHeader().setSectionResizeMode(i, 1)

        main_layout.addWidget(self.table)

        self.setWindowTitle("Converter")
        self.setLayout(main_layout)
        self.setGeometry(100, 100, 800, 600)

    def process_bin_file(self):
        if not self.bin_file:
            self.file_label.setText("Ошибка: выберите BIN-файл перед обработкой")
            return

        self.progress_bar.setValue(25)

        self.procces_workbook = parse_bin_file(self.bin_file)  # Получаем Workbook
        self.progress_bar.setValue(75)

        # Временный файл, который будем использовать при сортировке
        self.xlsx_file = os.path.join(os.getcwd(), "converted_file.xlsx")

        # Удаление BIN файла после конвертации
        try:
            self.bin_file = None
            self.file_label.setText(f"Файл {self.bin_file} удален после конвертации")
            self.btn_process.setEnabled(False)
        except Exception as e:
            self.file_label.setText(f"Ошибка при удалении BIN файла: {e}")

        self.btn_download.setEnabled(True)
        self.progress_bar.setValue(100)

        self.load_xlsx_preview()

    def select_bin_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите BIN-файл", "", "BIN files (*.bin)")
        if file_path:
            self.file_label.setText(f"Выбран файл: {file_path}")
            self.bin_file = file_path
            self.btn_process.setEnabled(True)  # Включаем кнопку "Обработать" после выбора BIN файла

    def load_xlsx_preview(self):
        if not hasattr(self, "procces_workbook"):
            self.file_label.setText("Ошибка: файл не обработан")
            return

        data = []

        for sheet in self.procces_workbook.worksheets:
            for row in sheet.iter_rows(min_row=2, values_only=True):
                data.append(row)

        df = pd.DataFrame(data, columns=["Порядковый номер", "Время", "Номер задачи", "Тип диагностического сообщения",
                                         "Длина бинарных данных", "Бинарные данные", "Текстовое сообщение разработчику"])

        self.update_table(df)

    def update_table(self, df):
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
            self.procces_workbook.save(self.xlsx_file)  # Сохраняем Workbook

    def apply_message_type_sorting(self, selected_types):
        self.sorted_workbook = sort_by_diag_type_message(self.xlsx_file, selected_types)

        if self.sorted_workbook:
            self.btn_download_sorted.setEnabled(True)  # Включаем кнопку для скачивания отсортированного файла

    def open_sort_message_window(self):
        if not self.xlsx_file:
            self.file_label.setText("Ошибка: выберите BIN-файл перед обработкой")
            return

        df = pd.read_excel(self.xlsx_file)
        unique_types = df["Тип диагностического сообщения"].dropna().unique().tolist()

        self.sort_window = SortByDiagMessageType(self, unique_types)
        self.sort_window.sorting_aplied.connect(self.apply_message_type_sorting)
        self.sort_window.show()

    def download_sorted_xlsx(self):
        if not hasattr(self, "sorted_workbook") or self.sorted_workbook is None:
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить отсортированный файл", "Sort_diag_type.xlsx", "XLSX Files (*.xlsx)")
        if save_path:
            self.sorted_workbook.save(save_path)  # Сохраняем только при скачивании


class SortByDiagMessageType(QDialog):
    sorting_aplied = pyqtSignal(list)

    def __init__(self, parent=None, message_types=None):
        super().__init__(parent)
        self.setWindowTitle("Сортировка по типу сообщений")
        self.setGeometry(200, 200, 500, 300)

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

        self.btn_sort = QPushButton("Выполнить сортировку")
        self.btn_sort.setEnabled(False)  # Изначально кнопка неактивна
        self.btn_sort.clicked.connect(self.apply_sorting)
        main_layout.addWidget(self.btn_sort)

        self.setLayout(main_layout)

        if message_types:
            for msg_type in message_types:
                if msg_type == 255:
                    continue

                self.list_all_types.addItem(QListWidgetItem(str(msg_type)))
                self.list_select_types.addItem(QListWidgetItem(str(msg_type)))

        # Подключаем сигнал изменения выделенных элементов
        self.list_select_types.itemSelectionChanged.connect(self.check_selection)

    def check_selection(self):
        # Проверяем, есть ли хотя бы один выбранный элемент в списке
        selected_items = self.list_select_types.selectedItems()
        if selected_items:
            self.btn_sort.setEnabled(True)  # Активируем кнопку
        else:
            self.btn_sort.setEnabled(False)  # Деактивируем кнопку, если ничего не выбрано

    def apply_sorting(self):
        selected_types = [int(item.text()) for item in self.list_select_types.selectedItems()]
        if selected_types:
            self.sorting_aplied.emit(selected_types)
        self.close()


app = QApplication(sys.argv)
ex = fileConverterApp()
ex.show()
sys.exit(app.exec_())
