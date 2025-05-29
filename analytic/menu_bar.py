from PyQt5.QtWidgets import (QMenuBar, QDialog, QAction, QMessageBox, QFileDialog, QPushButton, QListWidget, QLabel,
        QVBoxLayout, QHBoxLayout, QComboBox, QTextEdit, QProgressDialog, QSpinBox, QLineEdit, QTableWidget, QTableWidgetItem)
from database.db_manager import DB_PATH
from analytic.methods_bd import (get_all_files_from_db, get_db_size_mb, ExportEntireDatabaseThread, ExportSelectedFilesThread, delete_selected_files,
                                FIELD_MAPPING, CompareWorker, FilterSearchWorker, get_files_list, MetabaseLauncherThread)
from PyQt5.QtCore import pyqtSignal, Qt
import sqlite3, webbrowser, os
from Global import type_names

password = "Petya2003"

class AppMenuBar(QMenuBar):
    # Можно добавить сигналы, если нужно будет обновлять интерфейс
    request_show_files = pyqtSignal()
    request_run_analytics = pyqtSignal()
    request_filter_search = pyqtSignal()
    request_export = pyqtSignal()
    request_delete_file = pyqtSignal()
    request_compare_excel = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)

        db_menu = self.addMenu("База данных")

        show_files_action = QAction("Показать загруженные файлы", self)
        analytics_action = QAction("Аналитика", self)
        filter_search_action = QAction("Фильтр и поиск", self)
        delete_file_action = QAction("Удалить файл", self)
        compare_excel_action = QAction("Сравнить с Excel", self)

        db_menu.addAction(show_files_action)
        db_menu.addAction(analytics_action)
        db_menu.addSeparator()
        db_menu.addAction(filter_search_action)
        db_menu.addSeparator()
        db_menu.addAction(delete_file_action)
        db_menu.addAction(compare_excel_action)

        # Подключаем действия к слотам или сигналам
        show_files_action.triggered.connect(self.request_show_files.emit)
        analytics_action.triggered.connect(self.request_run_analytics.emit)
        filter_search_action.triggered.connect(self.request_filter_search.emit)
        delete_file_action.triggered.connect(self.request_delete_file.emit)
        compare_excel_action.triggered.connect(self.request_compare_excel.emit)


class ExportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Просмотр и экспорт данных из БД")
        self.setMinimumWidth(500)

        layout = QVBoxLayout()

        self.info_label = QLabel("Выберите файлы для экспорта или экспортируйте всю БД:")
        layout.addWidget(self.info_label)

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.MultiSelection)
        layout.addWidget(self.file_list)

        button_layout = QHBoxLayout()

        self.export_selected_btn = QPushButton("Экспортировать выбранные")
        self.export_selected_btn.clicked.connect(self.export_selected)
        button_layout.addWidget(self.export_selected_btn)

        self.export_all_btn = QPushButton("Экспортировать всю базу")
        self.export_all_btn.clicked.connect(self.export_all)
        button_layout.addWidget(self.export_all_btn)

        layout.addLayout(button_layout)
        self.setLayout(layout)

        self.load_file_list()

        # Для хранения активных потоков, чтобы не были собраны сборщиком мусора
        self.export_thread = None

    def load_file_list(self):
        self.file_list.clear()
        self.files = get_all_files_from_db()
        for f in self.files:
            self.file_list.addItem(f"{f['id']} - {f['file_name']} ({f['added_at']})")

    def export_selected(self):
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Нет выбора", "Выберите хотя бы один файл для экспорта.")
            return
        selected_ids = [int(item.text().split(' - ')[0]) for item in selected_items]

        self.export_selected_btn.setEnabled(False)
        self.export_all_btn.setEnabled(False)

        self.export_thread = ExportSelectedFilesThread(selected_ids)
        self.export_thread.finished.connect(self.on_export_finished)
        self.export_thread.start()

    def export_all(self):
        size_mb = get_db_size_mb()
        if size_mb > 5:
            reply = QMessageBox.question(
                self,
                "Большой размер БД",
                f"Размер БД составляет {size_mb:.2f} МБ. Экспорт может занять время. Продолжить?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        self.export_selected_btn.setEnabled(False)
        self.export_all_btn.setEnabled(False)

        self.export_thread = ExportEntireDatabaseThread()
        self.export_thread.finished.connect(self.on_export_finished)
        self.export_thread.start()

    def on_export_finished(self, message):
        QMessageBox.information(self, "Готово", message)
        self.export_selected_btn.setEnabled(True)
        self.export_all_btn.setEnabled(True)
        self.export_thread = None

class DeleteDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Удаление файлов из БД")
        self.setMinimumWidth(500)

        layout = QVBoxLayout()

        self.info_label = QLabel("Выберите файлы для удаления из базы данных:")
        layout.addWidget(self.info_label)

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.MultiSelection)
        layout.addWidget(self.file_list)

        button_layout = QHBoxLayout()

        self.delete_selected_btn = QPushButton("Удалить выбранные")
        self.delete_selected_btn.clicked.connect(self.delete_selected)
        button_layout.addWidget(self.delete_selected_btn)

        layout.addLayout(button_layout)
        self.setLayout(layout)

        self.load_file_list()

    def load_file_list(self):
        self.file_list.clear()
        self.files = get_all_files_from_db()
        for f in self.files:
            self.file_list.addItem(f"{f['id']} - {f['file_name']} ({f['added_at']})")

    def delete_selected(self):
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Нет выбора", "Выберите хотя бы один файл для удаления.")
            return

        file_names = [item.text().split(' - ')[1].split(' (')[0] for item in selected_items]
        file_names_str = "\n".join(file_names)

        reply = QMessageBox.question(
            self,
            "Подтверждение удаления",
            f"Вы действительно хотите удалить следующие файлы из базы данных?\n\n{file_names_str}",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        selected_ids = [int(item.text().split(' - ')[0]) for item in selected_items]

        try:
            delete_selected_files(selected_ids)
            QMessageBox.information(self, "Удалено", "Выбранные файлы успешно удалены.")
            self.load_file_list()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при удалении: {e}")


class SolveDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Сравнение с Excel")

        self.db_conn = sqlite3.connect(DB_PATH)

        self.excel_path = None
        self.db_file_id = None
        self.worker = None

        self.init_ui()
        self.load_files_from_db()

    def init_ui(self):
        layout = QVBoxLayout()

        btn_select_excel = QPushButton("Выбрать Excel-файл")
        btn_select_excel.clicked.connect(self.select_excel_file)
        self.label_excel = QLabel("Excel файл не выбран")
        layout.addWidget(btn_select_excel)
        layout.addWidget(self.label_excel)

        self.combo_db_files = QComboBox()
        self.combo_db_files.currentIndexChanged.connect(self.on_db_file_selected)
        layout.addWidget(QLabel("Выберите файл из базы:"))
        layout.addWidget(self.combo_db_files)

        btn_compare = QPushButton("Сравнить")
        btn_compare.clicked.connect(self.compare)
        layout.addWidget(btn_compare)

        self.text_result = QTextEdit()
        self.text_result.setReadOnly(True)
        layout.addWidget(self.text_result)

        self.setLayout(layout)

    def load_files_from_db(self):
        cursor = self.db_conn.cursor()
        cursor.execute("SELECT id, file_name FROM xlsx_files ORDER BY id")
        self.files_in_db = cursor.fetchall()
        self.combo_db_files.clear()
        for _id, fname in self.files_in_db:
            self.combo_db_files.addItem(fname, _id)
        if self.files_in_db:
            self.db_file_id = self.files_in_db[0][0]

    def select_excel_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите Excel-файл", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.excel_path = path
            self.label_excel.setText(f"Выбран файл: {path}")

    def on_db_file_selected(self, index):
        if index >= 0:
            self.db_file_id = self.combo_db_files.itemData(index)

    def compare(self):
        if not self.excel_path:
            QMessageBox.warning(self, "Ошибка", "Выберите Excel-файл для сравнения")
            return
        if self.db_file_id is None:
            QMessageBox.warning(self, "Ошибка", "Выберите файл из базы")
            return

        # Показываем прогресс-диалог
        self.progress = QProgressDialog("Сравнение файлов...", None, 0, 0, self)
        self.progress.setWindowTitle("Пожалуйста, подождите")
        self.progress.setCancelButton(None)
        self.progress.setWindowModality(Qt.WindowModal)
        self.progress.show()

        # Запускаем поток с обработкой
        self.worker = CompareWorker(self.excel_path, self.db_file_id)
        self.worker.finished.connect(self.on_compare_finished)
        self.worker.error.connect(self.on_compare_error)
        self.worker.start()

    def on_compare_finished(self, diffs, stats):
        self.progress.close()
        report = []
        report.append(f"Всего строк для сравнения: {stats['total_rows']}")
        report.append(f"Совпадающих строк: {stats['matched_rows']}")
        report.append(f"Строк с различиями: {stats['diff_rows']}")
        report.append("Частота различий по полям:")
        for field, count in stats['field_diff_counts'].items():
            report.append(f"  {field}: {count} раз(а)")

        report.append("\nДетали различий:")
        if diffs:
            for diff in diffs:
                row = diff["row"]
                fields = ", ".join(diff["diff_fields"])
                report.append(f"Строка {row}: поля с различиями - {fields}")
                for f in diff["diff_fields"]:
                    db_key = FIELD_MAPPING[f]
                    report.append(f"  {f}: Excel='{diff['excel_values'][f]}', БД='{diff['db_values'].get(db_key, '<нет>')}'")
                report.append("")
        else:
            report.append("Различий не обнаружено.")

        self.text_result.setPlainText("\n".join(report))

    def on_compare_error(self, message):
        self.progress.close()
        QMessageBox.critical(self, "Ошибка при сравнении", message)


class FilterSearcherDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db_conn = sqlite3.connect(DB_PATH)
        self.setWindowTitle("Фильтр и поиск по базе")

        self.file_id = None
        self.worker = None

        self.init_ui()
        self.load_files()

    def init_ui(self):
        layout = QVBoxLayout()

        # Выбор файла
        self.combo_files = QComboBox()
        self.combo_files.currentIndexChanged.connect(self.file_changed)
        layout.addWidget(QLabel("Выберите файл из базы:"))
        layout.addWidget(self.combo_files)

        # Фильтр по типу сообщения
        hbox_type = QHBoxLayout()
        hbox_type.addWidget(QLabel("Тип сообщения:"))
        self.combo_type = QComboBox()
        self.combo_type.addItem("Все", None)
        for key in sorted(type_names, reverse=True):
            self.combo_type.addItem(type_names[key], key)
        hbox_type.addWidget(self.combo_type)
        layout.addLayout(hbox_type)

        # Фильтр по номеру задачи
        hbox_task = QHBoxLayout()
        hbox_task.addWidget(QLabel("Номер задачи:"))
        self.combo_task_number = QComboBox()
        self.combo_task_number.addItem("Все", None)
        hbox_task.addWidget(self.combo_task_number)
        layout.addLayout(hbox_task)

        # Поиск по тексту
        layout.addWidget(QLabel("Поиск по тексту:"))
        self.edit_search = QLineEdit()
        layout.addWidget(self.edit_search)

        # Кнопка применить фильтр
        self.btn_apply = QPushButton("Применить фильтр")
        self.btn_apply.clicked.connect(self.apply_filter)
        layout.addWidget(self.btn_apply)

        # Таблица с результатами
        self.table_result = QTableWidget()
        self.table_result.setColumnCount(len(FIELD_MAPPING))
        self.table_result.setHorizontalHeaderLabels(FIELD_MAPPING.keys())
        self.table_result.horizontalHeader().setStretchLastSection(True)
        self.table_result.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_result.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.table_result)

        self.setLayout(layout)

    def load_files(self):
        files = get_files_list(self.db_conn)
        self.combo_files.clear()
        for fid, fname in files:
            self.combo_files.addItem(fname, fid)
        if files:
            self.file_id = files[0][0]
            self.load_task_numbers()

    def file_changed(self, index):
        if index >= 0:
            self.file_id = self.combo_files.itemData(index)
            self.load_task_numbers()

    def load_task_numbers(self):
        self.combo_task_number.clear()
        self.combo_task_number.addItem("Все", None)
        if self.file_id is None:
            return
        cursor = self.db_conn.cursor()
        cursor.execute(
            "SELECT DISTINCT task_number FROM messages WHERE xlsx_file_id = ? ORDER BY task_number ASC",
            (self.file_id,)
        )
        task_numbers = [row[0] for row in cursor.fetchall()]
        for number in task_numbers:
            self.combo_task_number.addItem(str(number), number)

    def apply_filter(self):
        if self.file_id is None:
            QMessageBox.warning(self, "Ошибка", "Выберите файл из базы")
            return

        filters = {}
        message_type = self.combo_type.currentData()
        filters["message_type"] = message_type

        task_number = self.combo_task_number.currentData()
        if task_number is not None:
            filters["task_number"] = task_number

        search_text = self.edit_search.text().strip()

        self.btn_apply.setEnabled(False)
        self.table_result.setRowCount(0)

        self.worker = FilterSearchWorker(DB_PATH, self.file_id, filters, search_text)
        self.worker.finished.connect(self.on_filter_finished)
        self.worker.error.connect(self.on_filter_error)
        self.worker.start()

    def on_filter_finished(self, results):
        self.btn_apply.setEnabled(True)
        self.table_result.setRowCount(0)

        if not results:
            QMessageBox.information(self, "Результаты", "Данные не найдены.")
            return

        ordered_keys = list(FIELD_MAPPING.keys())
        self.table_result.setRowCount(len(results))

        for row_index, msg in enumerate(results):
            for col_index, key in enumerate(ordered_keys):
                val = msg.get(key, "")
                item = QTableWidgetItem(str(val))
                self.table_result.setItem(row_index, col_index, item)

    def on_filter_error(self, error_msg):
        self.btn_apply.setEnabled(True)
        QMessageBox.critical(self, "Ошибка при загрузке", error_msg)
        self.table_result.setRowCount(0)



class AnalyticsLauncher:
    def __init__(self, parent=None, jar_path="database/metabase/metabase.jar", url="http://localhost:3000"):
        self.parent = parent
        self.jar_path = os.path.abspath(jar_path)
        self.url = url
        self.java_path = r"C:\Program Files\Java\jdk-24\bin\java.exe"
        self.thread = None

    def is_metabase_running(self):
        import requests
        try:
            requests.get(self.url, timeout=2)
            return True
        except:
            return False

    def open_analytics(self):
        if self.is_metabase_running():
            webbrowser.open(self.url)
            return

        self.thread = MetabaseLauncherThread(self.jar_path, self.java_path, self.url)
        self.thread.started_successfully.connect(self.on_metabase_ready)
        self.thread.start()

    def on_metabase_ready(self, success):
        if success:
            webbrowser.open(self.url)
        else:
            QMessageBox.warning(self.parent, "Ошибка", "Не удалось запустить Metabase.")

    def create_action(self):
        action = QAction("Аналитика", self.parent)
        action.triggered.connect(self.open_analytics)
        return action
