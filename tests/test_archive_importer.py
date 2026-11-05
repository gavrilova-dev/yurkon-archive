import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook
from app.services.archive_importer import read_archive_journal

class ArchiveImporterTests(unittest.TestCase):
    def test_reads_journal_headers_and_skips_numeric_header_row(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "journal.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Вход Архив ИЛ-10"
            ws.append(["Журнал"])
            ws.append(["№, п/п", "Индекс дела", "Наименование организации", "Номер протокола", "ФИО сдавшего дело в архив"])
            ws.append([1, 2, 3, 4, 5])
            ws.append([1, "ИЛ-10-0001/2025", "ООО Альфа", "9510/2025", "Кормин М.В."])
            ws.append([2, "ИЛ-10-0002/2025", "ООО Бета", "9511/2025", "Иванов И.И."])
            wb.save(path)
            rows = read_archive_journal(path)
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]["organization"], "ООО Альфа")
            self.assertEqual(rows[0]["archive_index"], "ИЛ-10-0001/2025")
            self.assertEqual(rows[0]["protocol"], "9510/2025")

    def test_does_not_skip_first_real_row_when_no_numeric_helper_row(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "journal.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Вход Архив ИЛ-10"
            ws.append(["Журнал"])
            ws.append(["№, п/п", "Индекс дела", "Наименование организации", "Номер протокола", "ФИО сдавшего дело в архив"])
            ws.append([1, "ИЛ-10-0001/2025", "ООО Альфа", "9510/2025", "Кормин М.В."])
            ws.append([2, "ИЛ-10-0002/2025", "ООО Бета", "9511/2025", "Кормин М.В."])
            wb.save(path)
            rows = read_archive_journal(path)
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0]["archive_index"], "ИЛ-10-0001/2025")

if __name__ == "__main__":
    unittest.main()
