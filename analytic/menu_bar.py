from PyQt5.QtWidgets import (QMenuBar, QDialog, QAction, QMessageBox, QFileDialog, QPushButton, QListWidget, QLabel,
        QVBoxLayout, QHBoxLayout, QComboBox, QTextEdit)
from database.db_manager import DB_PATH
from analytic.methods_bd import (get_all_files_from_db, get_db_size_mb, ExportEntireDatabaseThread, ExportSelectedFilesThread, delete_selected_files,
                                 load_db_messages, load_excel_messages, compare_messages, FIELD_MAPPING)
from PyQt5.QtCore import pyqtSignal
import sqlite3



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

        excel_msgs = load_excel_messages(self.excel_path)
        db_msgs = load_db_messages(self.db_file_id, self.db_conn)
        diffs, stats = compare_messages(excel_msgs, db_msgs)

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