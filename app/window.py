from app.services.file_watcher import FileWatcher

from pathlib import Path

from PySide6.QtWidgets import (
    QWidget,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QMessageBox,
)

from app.config import (
    EXCEL_PATH,
    INCOMING_FOLDER,
    READY_FOLDER,
)

from app.services.excel_service import ExcelService
from app.file_manager import FileManager


class MainWindow(QWidget):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("YURKON Archive")
        self.resize(500, 320)

        self.excel = ExcelService(EXCEL_PATH)
        self.excel.load()

        self.file_manager = FileManager(
            INCOMING_FOLDER,
            READY_FOLDER
        )

        self.current_pdf = None
        self.current_record = None


        self.file_label = QLabel("Файл: нет")


        self.protocol_edit = QLineEdit()
        self.protocol_edit.setPlaceholderText(
            "Введите номер протокола"
        )


        self.find_button = QPushButton(
            "Найти"
        )


        self.organization_label = QLabel("")


        self.index_label = QLabel("")


        self.rename_button = QPushButton(
            "Переименовать"
        )

        self.rename_button.setEnabled(False)


        layout = QVBoxLayout()

        layout.addWidget(self.file_label)
        layout.addWidget(self.protocol_edit)
        layout.addWidget(self.find_button)
        layout.addWidget(self.organization_label)
        layout.addWidget(self.index_label)
        layout.addWidget(self.rename_button)

        self.setLayout(layout)


        self.find_button.clicked.connect(
            self.find_protocol
        )


        self.protocol_edit.returnPressed.connect(
            self.find_protocol
        )


        self.rename_button.clicked.connect(
            self.rename_file
        )


        self.load_first_pdf()

        self.watcher = FileWatcher(
            INCOMING_FOLDER,
            self.load_first_pdf
        )

        self.watcher.start()



    def load_first_pdf(self):

        pdfs = self.file_manager.get_new_files()


        if not pdfs:

            self.current_pdf = None

            self.file_label.setText(
                "Нет новых PDF"
            )

            return


        self.current_pdf = pdfs[0]


        self.file_label.setText(
            f"Найдено файлов: {len(pdfs)}\n\n"
            f"Текущий:\n{self.current_pdf.name}"
        )



    def find_protocol(self):

        protocol = self.protocol_edit.text().strip()


        record = self.excel.find_by_protocol(
            protocol
        )


        if record is None:

            QMessageBox.warning(
                self,
                "Ошибка",
                "Протокол не найден."
            )

            self.current_record = None

            self.rename_button.setEnabled(
                False
            )

            return



        self.current_record = record


        self.organization_label.setText(
            f"Организация:\n{record.organization}"
        )


        self.index_label.setText(
            f"Индекс дела:\n{record.index}"
        )


        self.rename_button.setEnabled(
            True
        )



    def rename_file(self):

        if self.current_pdf is None:

            QMessageBox.warning(
                self,
                "Ошибка",
                "Нет PDF файла для обработки."
            )

            return



        if self.current_record is None:

            QMessageBox.warning(
                self,
                "Ошибка",
                "Сначала найдите протокол."
            )

            return



        reply = QMessageBox.question(
            self,
            "Подтверждение",
            f"Переименовать файл?\n\n"
            f"Протокол: {self.current_record.protocol}\n"
            f"Организация: {self.current_record.organization}\n"
            f"Индекс дела: {self.current_record.index}",
            QMessageBox.Yes | QMessageBox.No
        )


        if reply != QMessageBox.Yes:

            return



        try:

            destination = self.file_manager.move_pdf(
                self.current_pdf,
                self.current_record.index
            )


            self.excel.mark_scanned(
                self.current_record.excel_row
            )


            # self.excel.save()


            QMessageBox.information(
                self,
                "Готово",
                f"Файл сохранен:\n{destination.name}"
            )


            self.protocol_edit.clear()

            self.organization_label.clear()

            self.index_label.clear()


            self.rename_button.setEnabled(
                False
            )


            self.current_record = None


            self.load_first_pdf()



        except Exception as e:

            QMessageBox.critical(
                self,
                "Ошибка",
                str(e)
            )