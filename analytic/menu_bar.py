from PyQt5.QtWidgets import (QMenuBar, QDialog, QAction, QMessageBox, QFileDialog, QPushButton, QListWidget, QLabel,
        QVBoxLayout, QHBoxLayout)
from database.db_manager import DB_PATH
from analytic.methods_bd import get_all_files_from_db, get_db_size_mb, ExportEntireDatabaseThread, ExportSelectedFilesThread
from PyQt5.QtCore import pyqtSignal


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