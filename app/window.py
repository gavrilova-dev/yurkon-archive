from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, QDate
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
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

        self.protocol = QLineEdit(str(data.get("protocol") or ""))
        self.organization = QLineEdit(str(data.get("organization") or ""))
        self.workplaces = QSpinBox()
        self.workplaces.setRange(0, 100000)
        self.workplaces.setValue(int(data.get("workplaces") or 0))

        self.admin_done = OptionalDate()
        self.assembled = OptionalDate()
        self.returned = OptionalDate()
        self.signed = OptionalDate()
        for widget, key in ((self.admin_done, "admin_done_date"), (self.assembled, "assembled_date"),
                            (self.returned, "returned_date"), (self.signed, "signed_date")):
            widget.set_value(data.get(key))

        self.notes = QTextEdit()
        self.notes.setMaximumHeight(75)
        self.notes.setPlainText(data.get("notes") or "")

        form.addRow("Исполнитель", self.executor)
        form.addRow("Месяц", self.month)
        form.addRow("Номер протокола", self.protocol)
        form.addRow("Организация", self.organization)
        form.addRow("Рабочих мест", self.workplaces)
        form.addRow("Административная работа выполнена", self.admin_done)
        form.addRow("Документы собраны", self.assembled)
        form.addRow("Возвращены с подписи", self.returned)
        form.addRow("Подписаны", self.signed)
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
            "protocol": self.protocol.text().strip(),
            "organization": self.organization.text().strip(),
            "workplaces": self.workplaces.value(),
            "admin_done_date": self.admin_done.value(),
            "assembled_date": self.assembled.value(),
            "returned_date": self.returned.value(),
            "signed_date": self.signed.value(),
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


class MainWindow(QMainWindow):
    COLUMNS = [
        ("executor", "Исполнитель"), ("month", "Месяц"), ("protocol", "№ протокола"),
        ("organization", "Организация"), ("workplaces", "Мест"),
        ("fee", "К получению, ₽"), ("share", "Моя доля, ₽"),
        ("admin", "Админ, ₽"), ("archive_index", "Индекс дела"),
        ("assembled_date", "Собрано"), ("returned_date", "Возвращено"),
        ("signed_date", "Подписано"), ("scanned_date", "Скан"),
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
        self.watcher = FileWatcher(INCOMING_FOLDER, lambda: QTimer.singleShot(0, self.refresh_scan_queue))
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
            ("Перенести старый Excel", self.import_legacy),
            ("Экспорт отчётов", self.export_reports),
        ]
        for label, callback in actions:
            button = QPushButton(label)
            button.clicked.connect(callback)
            toolbar.addWidget(button)
        layout.addLayout(toolbar)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Поиск"))
        self.search = QLineEdit()
        self.search.setPlaceholderText("Организация, протокол, исполнитель, месяц или индекс дела")
        self.search.textChanged.connect(self.refresh_table)
        search_row.addWidget(self.search)
        layout.addLayout(search_row)

        self.table = QTableWidget(0, len(self.COLUMNS) + 1)
        self.table.setHorizontalHeaderLabels(["ID"] + [item[1] for item in self.COLUMNS])
        self.table.setColumnHidden(0, True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.doubleClicked.connect(self.edit_record)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
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
                record["id"], record["executor"], record["month"], record["protocol"],
                record["organization"], workplaces, fee, share, admin, record["archive_index"] or "",
                _display_date(record["assembled_date"]), _display_date(record["returned_date"]),
                _display_date(record["signed_date"]), _display_date(record["scanned_date"]),
            ]
            for col, value in enumerate(values):
                item = QTableWidgetItem(str(value if value is not None else ""))
                if col in (0, 5, 6, 7, 8):
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.table.setItem(row_num, col, item)
        self.summary.setText(
            f"Записей: {len(rows)} · Рабочих мест: {total_workplaces} · "
            f"К получению: {total_fee:,} ₽ · Моя доля: {total_share:,} ₽ · "
            f"Админ. работа: {total_admin:,} ₽".replace(",", " ")
        )

    def selected_record_id(self):
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return int(item.text()) if item else None

    def add_record(self):
        dialog = RecordDialog(self, executors=self.db.list_executors())
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.db.add_record(dialog.data())
            self.refresh_table()

    def edit_record(self, *_):
        record_id = self.selected_record_id()
        if not record_id:
            QMessageBox.information(self, "Запись не выбрана", "Выберите строку в таблице.")
            return
        record = self.db.get_record(record_id)
        dialog = RecordDialog(self, dict(record), self.db.list_executors())
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.db.update_record(record_id, dialog.data())
            self.refresh_table()

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
        batches = self.db.list_batches()
        labels = [f'{b["executor"]} — {b["month"]}' for b in batches]
        options = ["Добавить / изменить партию"] + labels
        choice, ok = QInputDialog.getItem(self, "Контроль партий", "Выберите партию:", options, 0, False)
        if not ok:
            return
        data = None
        if choice != options[0]:
            data = dict(batches[options.index(choice) - 1])
        dialog = BatchDialog(self, self.db.list_executors(), data)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.db.save_batch(dialog.data())
            self.statusBar().showMessage("Данные партии сохранены.", 5000)

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
            self, "Перенос данных",
            "Будут перенесены записи из листа «2026 Реестр» и даты партий. "
            "Повторные записи не добавляются. Продолжить?"
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            result = self.db.import_legacy_workbook(path)
            self.refresh_table()
            QMessageBox.information(
                self, "Перенос завершён",
                f'Добавлено записей: {result["imported"]}\nПовторный перенос той же книги не создаст дубликаты.'
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
        default_name = str(Path(template).with_name(f"Отчёт_YURKON_{date.today():%Y-%m-%d}.xlsx"))
        output, _ = QFileDialog.getSaveFileName(self, "Сохранить отчёт", default_name, "Excel (*.xlsx)")
        if not output:
            return
        if not output.lower().endswith(".xlsx"):
            output += ".xlsx"
        try:
            records = [dict(r) for r in self.db.list_records()]
            batches = [dict(b) for b in self.db.list_batches()]
            export_report(template, output, records, batches)
            QMessageBox.information(
                self, "Экспорт готов",
                f"Файл создан:\n{output}\n\nВсе исходные листы шаблона сохранены."
            )
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка экспорта", str(exc))

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
            self.db.connection.execute("""
                UPDATE primary_records SET scanned_date=?,pdf_path=?,updated_at=CURRENT_TIMESTAMP
                WHERE id=(SELECT id FROM primary_records
                          WHERE lower(trim(protocol))=lower(trim(?))
                            AND lower(trim(organization))=lower(trim(?))
                          ORDER BY id DESC LIMIT 1)
            """, (date.today().isoformat(), str(destination), rec["protocol"], rec["organization"]))
            self.db.connection.commit()
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
        finally:
            self.db.close()
        super().closeEvent(event)
