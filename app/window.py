from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import tempfile
from openpyxl import load_workbook

from PySide6.QtCore import Qt, QTimer, QDate, Signal
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QGridLayout,
    QTabWidget, QTableWidget, QTableWidgetItem, QPushButton, QLabel,
    QLineEdit, QMessageBox, QFileDialog, QComboBox, QSpinBox, QDateEdit,
    QCheckBox, QTextEdit, QListWidget, QListWidgetItem, QInputDialog,
    QGroupBox, QHeaderView, QAbstractItemView,
)

from app.config import APP_DATA, DATABASE_PATH, INCOMING_FOLDER, READY_FOLDER, ARCHIVE_JOURNAL_DEFAULT
from app.database import Database
from app.file_manager import FileManager
from app.services.file_watcher import FileWatcher
from app.services.archive_importer import read_archive_journal
from app.services.report_exporter import export_report

MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
          "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]


def _qdate(value):
    if not value:
        return None
    try:
        d = date.fromisoformat(str(value)[:10])
        return QDate(d.year, d.month, d.day)
    except (ValueError, TypeError):
        return None


def _display_date(value):
    if not value:
        return ""
    try:
        return datetime.fromisoformat(str(value)[:10]).strftime("%d.%m.%Y")
    except (ValueError, TypeError):
        return str(value)


class OptionalDate(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.enabled = QCheckBox("Указать")
        self.date = QDateEdit(QDate.currentDate())
        self.date.setCalendarPopup(True)
        self.date.setDisplayFormat("dd.MM.yyyy")
        self.date.setEnabled(False)
        self.enabled.toggled.connect(self.date.setEnabled)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.enabled)
        layout.addWidget(self.date)

    def set_value(self, value):
        qd = _qdate(value)
        self.enabled.setChecked(qd is not None)
        if qd:
            self.date.setDate(qd)

    def value(self):
        return self.date.date().toString("yyyy-MM-dd") if self.enabled.isChecked() else None


class RecordDialog(QDialog):
    def __init__(self, parent=None, data=None, executors=None):
        super().__init__(parent)
        self.setWindowTitle("Запись первичной документации")
        self.setMinimumWidth(540)
        data = dict(data or {})
        form = QFormLayout(self)

        self.executor = QComboBox()
        self.executor.setEditable(True)
        self.executor.addItems(executors or [])
        self.executor.setCurrentText(data.get("executor", ""))

        self.month = QComboBox()
        self.month.setEditable(True)
        self.month.addItems(MONTHS)
        self.month.setCurrentText(str(data.get("month", "")).capitalize())
        self.year = QSpinBox()
        self.year.setRange(2000, 2100)
        self.year.setValue(int(data.get("year") or date.today().year))

        self.protocol = QLineEdit(str(data.get("protocol") or ""))
        self.organization = QLineEdit(str(data.get("organization") or ""))
        self.workplaces = QSpinBox()
        self.workplaces.setRange(0, 100000)
        self.workplaces.setValue(int(data.get("workplaces") or 0))

        self.admin_done = OptionalDate()
        self.assembled = OptionalDate()
        self.yandex_uploaded = OptionalDate()
        for widget, key in ((self.admin_done, "admin_done_date"), (self.assembled, "assembled_date"),
                            (self.yandex_uploaded, "yandex_uploaded_date")):
            widget.set_value(data.get(key))

        self.notes = QTextEdit()
        self.notes.setMaximumHeight(75)
        self.notes.setPlainText(data.get("notes") or "")

        form.addRow("Исполнитель", self.executor)
        form.addRow("Месяц", self.month)
        form.addRow("Год", self.year)
        form.addRow("Номер протокола", self.protocol)
        form.addRow("Организация", self.organization)
        form.addRow("Рабочих мест", self.workplaces)
        form.addRow("Административная работа выполнена", self.admin_done)
        form.addRow("Документы собраны", self.assembled)
        form.addRow("Выгружено на Яндекс.Диск", self.yandex_uploaded)
        form.addRow("Комментарий", self.notes)

        buttons = QHBoxLayout()
        save, cancel = QPushButton("Сохранить"), QPushButton("Отмена")
        save.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        buttons.addStretch()
        buttons.addWidget(cancel)
        buttons.addWidget(save)
        form.addRow(buttons)

    def data(self):
        return {
            "executor": self.executor.currentText().strip(),
            "month": self.month.currentText().strip().casefold(),
            "year": self.year.value(),
            "protocol": self.protocol.text().strip(),
            "organization": self.organization.text().strip(),
            "workplaces": self.workplaces.value(),
            "admin_done_date": self.admin_done.value(),
            "assembled_date": self.assembled.value(),
            "yandex_uploaded_date": self.yandex_uploaded.value(),
            "notes": self.notes.toPlainText().strip(),
        }

    def accept(self):
        if not self.protocol.text().strip() or not self.organization.text().strip():
            QMessageBox.warning(self, "Не заполнены данные", "Укажите номер протокола и организацию.")
            return
        if not self.executor.currentText().strip() or not self.month.currentText().strip():
            QMessageBox.warning(self, "Не заполнены данные", "Укажите исполнителя и месяц.")
            return
        super().accept()


class BatchDialog(QDialog):
    FIELDS = [
        ("primary_given_date", "Первичка отдана в работу"),
        ("primary_returned_date", "Первичка возвращена исполнителем"),
        ("primary_report_date", "Отчёт по первичке сдан руководителю"),
        ("primary_paid_date", "Первичка оплачена"),
        ("admin_report_date", "Отчёт по админ. работе сдан"),
        ("admin_paid_date", "Админ. работа оплачена"),
    ]

    def __init__(self, parent, executors, data=None):
        super().__init__(parent)
        self.setWindowTitle("Контроль партии")
        self.setMinimumWidth(560)
        data = dict(data or {})
        form = QFormLayout(self)
        self.executor = QComboBox()
        self.executor.setEditable(True)
        self.executor.addItems(executors)
        self.executor.setCurrentText(data.get("executor", ""))
        self.month = QComboBox()
        self.month.setEditable(True)
        self.month.addItems(MONTHS)
        self.month.setCurrentText(str(data.get("month", "")).capitalize())
        form.addRow("Исполнитель", self.executor)
        form.addRow("Месяц", self.month)

        self.dates = {}
        for key, label in self.FIELDS:
            field = OptionalDate()
            field.set_value(data.get(key))
            self.dates[key] = field
            form.addRow(label, field)
        self.comment = QLineEdit(str(data.get("comment") or ""))
        form.addRow("Комментарий", self.comment)

        buttons = QHBoxLayout()
        save, cancel = QPushButton("Сохранить"), QPushButton("Отмена")
        save.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        buttons.addStretch()
        buttons.addWidget(cancel)
        buttons.addWidget(save)
        form.addRow(buttons)

    def data(self):
        result = {"executor": self.executor.currentText().strip(),
                  "month": self.month.currentText().strip().casefold(),
                  "comment": self.comment.text().strip()}
        result.update({key: field.value() for key, field in self.dates.items()})
        return result

    def accept(self):
        if not self.executor.currentText().strip() or not self.month.currentText().strip():
            QMessageBox.warning(self, "Не заполнены данные", "Укажите исполнителя и месяц.")
            return
        super().accept()


class BatchManagerDialog(QDialog):
    HEADERS = [
        "Исполнитель", "Месяц", "Год", "Статус партии", "Организаций", "Первичка отдана",
        "Первичка возвращена", "К получению, ₽", "Моя доля, ₽",
        "Отчёт по первичке сдан", "Первичка оплачена", "Админ. выполнено",
        "Админ. сумма, ₽", "Отчёт по админке сдан", "Админка оплачена", "Комментарий",
    ]
    FIELDS = [
        "primary_given_date", "primary_returned_date", "primary_report_date",
        "primary_paid_date", "admin_report_date", "admin_paid_date",
    ]

    def __init__(self, parent, db):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Контроль партий")
        self.resize(1450, 650)
        layout = QVBoxLayout(self)
        hint = QLabel("Даты вводите в формате ДД.ММ.ГГГГ. Изменения сохраняются одной кнопкой.")
        layout.addWidget(hint)
        self.table = QTableWidget(0, len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(15, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table)
        self._load_rows()
        buttons = QHBoxLayout()
        add = QPushButton("Добавить партию")
        save = QPushButton("Сохранить изменения")
        close = QPushButton("Закрыть")
        add.clicked.connect(self.add_batch)
        save.clicked.connect(self.save)
        close.clicked.connect(self.accept)
        buttons.addWidget(add)
        buttons.addStretch()
        buttons.addWidget(save)
        buttons.addWidget(close)
        layout.addLayout(buttons)

    def _load_rows(self):
        batches = [dict(row) for row in self.db.list_batches()]
        self.table.setRowCount(len(batches))
        records = self.db.list_records()
        aggregates = {}
        for rec in records:
            key = (rec["executor"], rec["month"], rec["year"])
            item = aggregates.setdefault(key, {"count": 0, "workplaces": 0, "admin_count": 0})
            item["count"] += 1
            item["workplaces"] += int(rec["workplaces"] or 0)
            if rec["admin_done_date"]:
                item["admin_count"] += 1
        for r, batch in enumerate(batches):
            if batch.get("primary_paid_date") and batch.get("admin_paid_date"):
                status = "Оплачена"
            elif batch.get("primary_paid_date") or batch.get("admin_paid_date"):
                status = "Частично оплачена"
            elif batch.get("primary_returned_date"):
                status = "Возвращена"
            elif batch.get("primary_given_date"):
                status = "В работе"
            else:
                status = "Не передана"
            totals = aggregates.get((batch["executor"], batch["month"], batch["year"]), {})
            workplaces = totals.get("workplaces", 0)
            admin_count = totals.get("admin_count", 0)
            values = [
                batch["executor"], batch["month"], batch["year"], status, totals.get("count", 0),
                _display_date(batch.get("primary_given_date")),
                _display_date(batch.get("primary_returned_date")),
                workplaces * 20, workplaces * 10,
                _display_date(batch.get("primary_report_date")),
                _display_date(batch.get("primary_paid_date")),
                admin_count, admin_count * 100,
                _display_date(batch.get("admin_report_date")),
                _display_date(batch.get("admin_paid_date")),
                batch.get("comment") or "",
            ]
            for c, value in enumerate(values):
                item = QTableWidgetItem(str(value if value is not None else ""))
                if c in (3, 4, 7, 8, 11, 12):
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(r, c, item)

    def add_batch(self):
        value, ok = QInputDialog.getText(
            self, "Новая партия", "Введите исполнителя и месяц через дефис (например: Лилия - февраль):"
        )
        if not ok or not value.strip():
            return
        parts = value.rsplit("-", 1)
        if len(parts) != 2 or not all(part.strip() for part in parts):
            QMessageBox.warning(self, "Неверный формат", "Укажите исполнителя и месяц через дефис.")
            return
        year, ok = QInputDialog.getInt(self, "Год партии", "Год:", date.today().year, 2000, 2100)
        if not ok:
            return
        row = self.table.rowCount()
        self.table.insertRow(row)
        values = [parts[0].strip(), parts[1].strip().casefold(), str(year), "Не передана", "0",
                  "", "", "0", "0", "", "", "0", "0", "", "", ""]
        for c, value in enumerate(values):
            self.table.setItem(row, c, QTableWidgetItem(value))

    @staticmethod
    def _to_iso(value):
        value = (value or "").strip()
        if not value:
            return None
        for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(value, fmt).date().isoformat()
            except ValueError:
                pass
        raise ValueError(f"Неверная дата «{value}». Используйте ДД.ММ.ГГГГ.")

    def save(self):
        try:
            for r in range(self.table.rowCount()):
                executor = self.table.item(r, 0).text().strip()
                month = self.table.item(r, 1).text().strip().casefold()
                year = int(self.table.item(r, 2).text())
                if not executor or not month:
                    raise ValueError(f"Не заполнены исполнитель или месяц в строке {r + 1}.")
                data = {"executor": executor, "month": month, "year": year,
                        "comment": self.table.item(r, 15).text().strip()}
                for i, field in zip((5, 6, 9, 10, 13, 14), self.FIELDS):
                    data[field] = self._to_iso(self.table.item(r, i).text())
                self.db.save_batch(data)
            QMessageBox.information(self, "Сохранено", "Данные по партиям обновлены.")
            self.accept()
        except Exception as exc:
            QMessageBox.warning(self, "Не удалось сохранить", str(exc))


class ReportPreviewDialog(QDialog):
    SHEETS = [
        ("2026 Сводка по партиям", "Партии"),
        ("2026 Отчёт руководителю", "Отчёт руководителю"),
        ("2026 Годовой отчет", "Годовой отчёт"),
    ]

    def __init__(self, parent, workbook_path):
        super().__init__(parent)
        self.setWindowTitle("Предпросмотр отчётности")
        self.resize(1350, 760)
        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        wb = load_workbook(workbook_path, data_only=False, read_only=True)
        for sheet_name, title in self.SHEETS:
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            table = QTableWidget()
            max_rows = min(ws.max_row, 250)
            max_cols = min(ws.max_column, 20)
            table.setRowCount(max_rows)
            table.setColumnCount(max_cols)
            table.setHorizontalHeaderLabels([
                str(ws.cell(1, c).value or f"Столбец {c}") for c in range(1, max_cols + 1)
            ])
            for r in range(1, max_rows + 1):
                for c in range(1, max_cols + 1):
                    value = ws.cell(r, c).value
                    if value is not None:
                        table.setItem(r - 1, c - 1, QTableWidgetItem(str(value)))
            table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            table.setAlternatingRowColors(True)
            table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
            self.tabs.addTab(table, title)
        wb.close()
        layout.addWidget(self.tabs)
        note = QLabel("Показаны первые 250 строк каждого листа. Итоговый Excel сохранит все строки и листы шаблона.")
        layout.addWidget(note)
        buttons = QHBoxLayout()
        cancel = QPushButton("Отмена")
        proceed = QPushButton("Выгрузить отчёт")
        cancel.clicked.connect(self.reject)
        proceed.clicked.connect(self.accept)
        buttons.addStretch()
        buttons.addWidget(cancel)
        buttons.addWidget(proceed)
        layout.addLayout(buttons)


class MainWindow(QMainWindow):
    scan_files_changed = Signal()
    COLUMNS = [
        ("executor", "Исполнитель"), ("month", "Месяц"), ("year", "Год"), ("protocol", "№ протокола"),
        ("organization", "Организация"), ("workplaces", "Мест"),
        ("fee", "К получению, ₽"), ("share", "Моя доля, ₽"),
        ("admin", "Админ, ₽"), ("archive_index", "Индекс дела"),
        ("assembled_date", "Документы собраны"), ("scanned_date", "Отсканированы"),
        ("yandex_uploaded_date", "Яндекс.Диск"),
    ]

    def __init__(self):
        super().__init__()
        self.setWindowTitle("YURKON Archive")
        self.resize(1450, 820)
        self.db = Database(DATABASE_PATH)
        self.file_manager = FileManager(INCOMING_FOLDER, READY_FOLDER)
        self.current_pdf = None
        self.current_archive_record = None
        self._build_ui()
        self.refresh_table()
        self.refresh_scan_queue()
        self.scan_files_changed.connect(self.refresh_scan_queue)
        self.watcher = FileWatcher(INCOMING_FOLDER, self.scan_files_changed.emit)
        self.watcher.start()
        self.statusBar().showMessage(
            f"База: {DATABASE_PATH} · Записей архива: {self.db.archive_count()}"
        )

    def _build_ui(self):
        tabs = QTabWidget()
        self.setCentralWidget(tabs)
        tabs.addTab(self._build_registry_tab(), "Учёт организаций")
        tabs.addTab(self._build_scanner_tab(), "Обработка сканов")

    def _build_registry_tab(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        toolbar = QHBoxLayout()
        actions = [
            ("Добавить запись", self.add_record),
            ("Изменить", self.edit_record),
            ("Удалить", self.delete_record),
            ("Контроль партий", self.manage_batches),
            ("Импорт архива", self.import_archive),
            ("Загрузить отчетность", self.import_legacy),
            ("Выгрузить отчёт", self.export_reports),
            ("Резервная копия", self.backup_database),
        ]
        for label, callback in actions:
            button = QPushButton(label)
            button.clicked.connect(callback)
            toolbar.addWidget(button)
        layout.addLayout(toolbar)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Поиск"))
        self.search = QLineEdit()
        self.search.setPlaceholderText("Организация, протокол, исполнитель, месяц, год или индекс дела")
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self.refresh_table)
        self.search.textChanged.connect(lambda _text: self._search_timer.start(250))
        search_row.addWidget(self.search)
        layout.addLayout(search_row)

        filter_grid = QGridLayout()
        self.filter_boxes = {}
        filter_specs = (("year", "Год"), ("executor", "Исполнитель"), ("month", "Месяц"),
                        ("organization", "Организация"), ("protocol", "Протокол"),
                        ("archive_index", "Индекс дела"),
                        ("assembled", "Документы собраны"), ("scanned", "Отсканированы"),
                        ("yandex", "Яндекс.Диск"))
        for index, (key, label) in enumerate(filter_specs):
            grid_row, group = divmod(index, 4)
            filter_grid.addWidget(QLabel(label), grid_row, group * 2)
            combo = QComboBox()
            combo.setMaximumWidth(190)
            combo.addItem("Все", "")
            combo.currentIndexChanged.connect(self.refresh_table)
            self.filter_boxes[key] = combo
            filter_grid.addWidget(combo, grid_row, group * 2 + 1)
        layout.addLayout(filter_grid)

        self.table = QTableWidget(0, len(self.COLUMNS) + 1)
        self.table.setHorizontalHeaderLabels(["ID"] + [item[1] for item in self.COLUMNS])
        self.table.setColumnHidden(0, True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.doubleClicked.connect(self.edit_record)
        self.table.itemChanged.connect(self.update_status_from_table)
        self._filter_values_cache = {}
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table)
        self.summary = QLabel()
        layout.addWidget(self.summary)
        return page

    def _build_scanner_tab(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        group = QGroupBox("Новые PDF в папке сканера")
        group_layout = QVBoxLayout(group)
        self.scan_file_label = QLabel("Новых PDF нет")
        self.scan_queue = QListWidget()
        self.scan_queue.currentItemChanged.connect(self.select_pdf)
        group_layout.addWidget(self.scan_file_label)
        group_layout.addWidget(self.scan_queue)
        layout.addWidget(group)

        search_group = QGroupBox("Поиск организации по номеру протокола")
        search_layout = QVBoxLayout(search_group)
        row = QHBoxLayout()
        self.scan_protocol = QLineEdit()
        self.scan_protocol.setPlaceholderText("Например, 7609/2024")
        self.scan_protocol.returnPressed.connect(self.search_archive)
        find = QPushButton("Найти")
        find.clicked.connect(self.search_archive)
        row.addWidget(self.scan_protocol)
        row.addWidget(find)
        search_layout.addLayout(row)

        self.archive_results = QListWidget()
        self.archive_results.currentItemChanged.connect(self.select_archive_record)
        search_layout.addWidget(self.archive_results)
        self.archive_detail = QLabel("Выберите найденную запись архива.")
        self.archive_detail.setWordWrap(True)
        search_layout.addWidget(self.archive_detail)
        self.rename_button = QPushButton("Переименовать и перенести PDF")
        self.rename_button.clicked.connect(self.move_scanned_pdf)
        self.rename_button.setEnabled(False)
        search_layout.addWidget(self.rename_button)
        layout.addWidget(search_group)
        return page

    def refresh_table(self, *_):
        if not hasattr(self, "table"):
            return
        rows = self.db.list_records(self.search.text() if hasattr(self, "search") else "")
        if hasattr(self, "filter_boxes"):
            filter_values = {
                key: sorted({str(r[key]) for r in rows if r.get(key)})
                for key in ("year", "executor", "month", "organization", "protocol", "archive_index")
            }
            for key, values in filter_values.items():
                combo = self.filter_boxes[key]
                selected = combo.currentData()
                if self._filter_values_cache.get(key) != values:
                    combo.blockSignals(True)
                    combo.clear()
                    combo.addItem("Все", "")
                    for value in values:
                        combo.addItem(value, value)
                    idx = combo.findData(selected)
                    if idx >= 0:
                        combo.setCurrentIndex(idx)
                    elif key == "year":
                        current_idx = combo.findData(str(date.today().year))
                        if current_idx >= 0:
                            combo.setCurrentIndex(current_idx)
                    combo.blockSignals(False)
                    self._filter_values_cache[key] = values
            for key in ("assembled", "scanned", "yandex"):
                combo = self.filter_boxes[key]
                if self._filter_values_cache.get(key) != ["yes", "no"]:
                    selected = combo.currentData()
                    combo.blockSignals(True)
                    combo.clear()
                    combo.addItem("Все", "")
                    combo.addItem("Есть", "yes")
                    combo.addItem("Нет", "no")
                    idx = combo.findData(selected)
                    if idx >= 0:
                        combo.setCurrentIndex(idx)
                    combo.blockSignals(False)
                    self._filter_values_cache[key] = ["yes", "no"]
            filtered = []
            for record in rows:
                ok = True
                for key in ("year", "executor", "month", "organization", "protocol", "archive_index"):
                    value = self.filter_boxes[key].currentData()
                    if value and str(record.get(key) or "") != value:
                        ok = False
                for key, field in (("assembled", "assembled_date"), ("scanned", "scanned_date"),
                                   ("yandex", "yandex_uploaded_date")):
                    value = self.filter_boxes[key].currentData()
                    if value == "yes" and not record[field]:
                        ok = False
                    if value == "no" and record[field]:
                        ok = False
                if ok:
                    filtered.append(record)
            rows = filtered

        self.table.blockSignals(True)
        self.table.setRowCount(len(rows))
        total_workplaces = total_fee = total_share = total_admin = 0
        for row_num, record in enumerate(rows):
            workplaces = int(record["workplaces"] or 0)
            fee, share = workplaces * 20, workplaces * 10
            admin = 100 if record["admin_done_date"] else 0
            total_workplaces += workplaces
            total_fee += fee
            total_share += share
            total_admin += admin
            values = [
                record["id"], record["executor"], record["month"], record["year"], record["protocol"],
                record["organization"], workplaces, fee, share, admin, record["archive_index"] or "",
                record["assembled_date"], record["scanned_date"], record["yandex_uploaded_date"],
            ]
            for col, value in enumerate(values):
                item = QTableWidgetItem("")
                if col in (11, 12, 13):
                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                    item.setCheckState(Qt.CheckState.Checked if value else Qt.CheckState.Unchecked)
                    item.setData(Qt.ItemDataRole.UserRole, record["id"])
                else:
                    item.setText(str(value if value is not None else ""))
                if col in (0, 6, 7, 8, 9):
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.table.setItem(row_num, col, item)
        self.table.blockSignals(False)
        self._visible_admin_total = total_admin
        self.summary.setText(
            f"Записей: {len(rows)} · Рабочих мест: {total_workplaces} · "
            f"К получению: {total_fee:,} ₽ · Моя доля: {total_share:,} ₽ · "
            f"Админ. работа: {total_admin:,} ₽".replace(",", " ")
        )

    def update_status_from_table(self, item):
        if item.column() not in (11, 12, 13):
            return
        record_id = item.data(Qt.ItemDataRole.UserRole)
        if not record_id:
            return
        field = {11: "assembled_date", 12: "scanned_date", 13: "yandex_uploaded_date"}[item.column()]
        stamp = date.today().isoformat() if item.checkState() == Qt.CheckState.Checked else None
        try:
            self.db.update_status(int(record_id), field, stamp)
            row = item.row()
            if item.column() == 12:
                admin_item = self.table.item(row, 9)
                if admin_item:
                    admin_item.setText("100" if stamp else "0")
                # Update the visible summary only; do not reload the full registry.
                if hasattr(self, "_visible_admin_total"):
                    self._visible_admin_total += 100 if stamp else -100
                    self.summary.setText(self.summary.text().split(" · Админ. работа:")[0] +
                                         f" · Админ. работа: {self._visible_admin_total:,} ₽".replace(",", " "))
        except Exception as exc:
            QMessageBox.critical(self, "Не удалось сохранить статус", str(exc))

    def selected_record_id(self):
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return int(item.text()) if item else None

    def add_record(self):
        dialog = RecordDialog(self, executors=self.db.list_executors())
        if dialog.exec() == QDialog.DialogCode.Accepted:
            try:
                data = dialog.data()
                self.db.add_record(data)
                self.refresh_table()
                year_combo = self.filter_boxes.get("year")
                if year_combo:
                    idx = year_combo.findData(str(data["year"]))
                    if idx >= 0:
                        year_combo.setCurrentIndex(idx)
            except ValueError as exc:
                QMessageBox.warning(self, "Дубликат протокола", str(exc))

    def edit_record(self, *_):
        record_id = self.selected_record_id()
        if not record_id:
            QMessageBox.information(self, "Запись не выбрана", "Выберите строку в таблице.")
            return
        record = self.db.get_record(record_id)
        dialog = RecordDialog(self, dict(record), self.db.list_executors())
        if dialog.exec() == QDialog.DialogCode.Accepted:
            try:
                self.db.update_record(record_id, dialog.data())
                self.refresh_table()
            except ValueError as exc:
                QMessageBox.warning(self, "Дубликат протокола", str(exc))

    def delete_record(self):
        record_id = self.selected_record_id()
        if not record_id:
            QMessageBox.information(self, "Запись не выбрана", "Выберите строку в таблице.")
            return
        answer = QMessageBox.question(self, "Удалить запись", "Удалить выбранную запись?")
        if answer == QMessageBox.StandardButton.Yes:
            self.db.delete_record(record_id)
            self.refresh_table()

    def manage_batches(self):
        dialog = BatchManagerDialog(self, self.db)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.statusBar().showMessage("Данные партий сохранены.", 5000)

    def import_archive(self):
        default = ARCHIVE_JOURNAL_DEFAULT if Path(ARCHIVE_JOURNAL_DEFAULT).exists() else str(APP_DATA)
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите журнал архивных дел", default, "Excel (*.xlsx *.xlsm)"
        )
        if not path:
            return
        try:
            rows = read_archive_journal(path)
            result = self.db.import_archive_rows(rows)
            self.statusBar().showMessage(
                f'Архив: добавлено {result["added"]}, уже были {result["updated"]}, пропущено {result["skipped"]}.', 10000
            )
            QMessageBox.information(
                self, "Импорт завершён",
                f'Строк прочитано: {len(rows)}\nДобавлено: {result["added"]}\n'
                f'Уже были в базе: {result["updated"]}\nПропущено: {result["skipped"]}'
            )
            self.refresh_table()
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка импорта", str(exc))

    def import_legacy(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите старую книгу учёта", str(APP_DATA), "Excel (*.xlsx *.xlsm)"
        )
        if not path:
            return
        answer = QMessageBox.question(
            self, "Загрузка отчетности",
            "Будут загружены записи из листа «2026 Реестр» и данные по партиям. "
            "Повторные записи не добавляются. Если в отчётности указано 100 ₽ за админ. работу, "
            "будут отмечены статусы «Документы собраны» и «Отсканированы». Продолжить?"
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            result = self.db.import_legacy_workbook(path)
            self.db.set_setting("report_template", path)
            self.refresh_table()
            QMessageBox.information(
                self, "Загрузка завершена",
                f'Добавлено записей: {result["imported"]}\n'
                f'Отмечено «Документы собраны» и «Отсканированы»: {result["statuses_updated"]}\n'
                f'Повторная загрузка той же книги не создаст дубликаты.'
            )
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка переноса", str(exc))

    def export_reports(self):
        template = self.db.get_setting("report_template")
        if not template or not Path(template).exists():
            template, _ = QFileDialog.getOpenFileName(
                self, "Выберите исходный шаблон 2026_учет_контор.xlsx",
                str(APP_DATA), "Excel (*.xlsx *.xlsm)"
            )
            if not template:
                return
            self.db.set_setting("report_template", template)
        available_years = sorted({str(r["year"]) for r in self.db.list_records()}, reverse=True)
        if not available_years:
            QMessageBox.information(self, "Нет данных", "Сначала добавьте или загрузите организации.")
            return
        current_filter = self.filter_boxes["year"].currentData()
        default_year = current_filter if current_filter in available_years else available_years[0]
        year, ok = QInputDialog.getItem(self, "Выгрузить отчёт", "За какой год подготовить отчёт?", available_years,
                                        available_years.index(default_year), False)
        if not ok:
            return
        default_name = str(Path(template).with_name(f"Отчёт_YURKON_{year}_{date.today():%Y-%m-%d}.xlsx"))
        try:
            records = sorted((dict(r) for r in self.db.list_records() if str(r["year"]) == year),
                             key=lambda item: item["id"])
            batches = [dict(b) for b in self.db.list_batches() if str(b["year"]) == year]
            with tempfile.TemporaryDirectory() as temp_dir:
                preview_path = Path(temp_dir) / "preview.xlsx"
                export_report(template, preview_path, records, batches)
                preview = ReportPreviewDialog(self, preview_path)
                if preview.exec() != QDialog.DialogCode.Accepted:
                    return
            output, _ = QFileDialog.getSaveFileName(self, "Сохранить отчёт", default_name, "Excel (*.xlsx)")
            if not output:
                return
            if not output.lower().endswith(".xlsx"):
                output += ".xlsx"
            export_report(template, output, records, batches)
            QMessageBox.information(
                self, "Экспорт готов",
                f"Файл создан:\n{output}\n\nГод отчётности: {year}. Все исходные листы шаблона сохранены."
            )
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка экспорта", str(exc))

    def backup_database(self):
        default = str(APP_DATA / f"yurkon_archive_backup_{datetime.now():%Y%m%d_%H%M%S}.sqlite3")
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить резервную копию базы", default, "SQLite (*.sqlite3)")
        if not path:
            return
        if not path.lower().endswith(".sqlite3"):
            path += ".sqlite3"
        try:
            self.db.backup_to(path)
            QMessageBox.information(self, "Резервная копия создана", f"База сохранена:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка резервного копирования", str(exc))

    def refresh_scan_queue(self):
        files = self.file_manager.get_new_files()
        selected_path = self.current_pdf
        self.scan_queue.blockSignals(True)
        self.scan_queue.clear()
        for path in files:
            item = QListWidgetItem(path.name)
            item.setData(Qt.ItemDataRole.UserRole, str(path))
            self.scan_queue.addItem(item)
            if selected_path and str(path) == str(selected_path):
                self.scan_queue.setCurrentItem(item)
        if self.scan_queue.count() and self.scan_queue.currentRow() < 0:
            self.scan_queue.setCurrentRow(0)
        self.scan_queue.blockSignals(False)
        if not files:
            self.current_pdf = None
            self.scan_file_label.setText("Новых PDF нет")
        else:
            self.scan_file_label.setText(f"В папке новых PDF: {len(files)}")
            if self.current_pdf is None:
                self.current_pdf = files[0]
        self._update_scan_button()

    def select_pdf(self, current, _previous=None):
        if current:
            self.current_pdf = Path(current.data(Qt.ItemDataRole.UserRole))
        self._update_scan_button()

    def search_archive(self):
        protocol = self.scan_protocol.text().strip()
        self.archive_results.clear()
        self.current_archive_record = None
        if not protocol:
            QMessageBox.warning(self, "Введите номер", "Укажите номер протокола для поиска.")
            return
        matches = self.db.find_archive(protocol)
        if not matches:
            QMessageBox.information(
                self, "Не найдено",
                "Совпадений нет. Проверьте, что импортирована актуальная версия архивного журнала."
            )
            self.archive_detail.setText("Совпадений нет.")
            self._update_scan_button()
            return
        for record in matches:
            item = QListWidgetItem(
                f'{record["protocol"]} — {record["organization"]} — {record["archive_index"]}'
            )
            item.setData(Qt.ItemDataRole.UserRole, dict(record))
            self.archive_results.addItem(item)
        if len(matches) == 1:
            self.archive_results.setCurrentRow(0)
        else:
            self.archive_detail.setText(f"Найдено совпадений: {len(matches)}. Выберите нужную организацию.")
        self._update_scan_button()

    def select_archive_record(self, current, _previous=None):
        self.current_archive_record = current.data(Qt.ItemDataRole.UserRole) if current else None
        if self.current_archive_record:
            rec = self.current_archive_record
            self.archive_detail.setText(
                f'Организация: {rec["organization"]}\nПротокол: {rec["protocol"]}\nИндекс дела: {rec["archive_index"]}'
            )
        self._update_scan_button()

    def _update_scan_button(self):
        if hasattr(self, "rename_button"):
            self.rename_button.setEnabled(bool(self.current_pdf and self.current_archive_record))

    def move_scanned_pdf(self):
        if not self.current_pdf or not self.current_archive_record:
            return
        rec = self.current_archive_record
        answer = QMessageBox.question(
            self, "Подтверждение",
            f'Перенести файл {self.current_pdf.name}?\n\n{rec["organization"]}\n'
            f'Протокол: {rec["protocol"]}\nИндекс дела: {rec["archive_index"]}'
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            destination = self.file_manager.move_pdf(self.current_pdf, rec["archive_index"])
            self.db.mark_scanned(
                rec["protocol"], rec["organization"], str(destination), date.today().isoformat()
            )
            QMessageBox.information(self, "Готово", f"Файл сохранён:\n{destination}")
            self.current_pdf = None
            self.current_archive_record = None
            self.archive_detail.setText("Файл обработан. Выберите следующий PDF.")
            self.refresh_scan_queue()
            self.refresh_table()
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка обработки PDF", str(exc))

    def closeEvent(self, event):
        try:
            self.watcher.stop()
            backup_dir = APP_DATA / "backups"
            backup_dir.mkdir(parents=True, exist_ok=True)
            self.db.backup_to(backup_dir / f"yurkon_archive_{datetime.now():%Y%m%d_%H%M%S}.sqlite3")
        finally:
            self.db.close()
        super().closeEvent(event)
