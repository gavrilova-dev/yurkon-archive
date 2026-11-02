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
            ws.append(["№, п/п", "Индекс дела", "Наименование организации", "Номер протокола"])
            ws.append([1, 2, 3, 4])
            ws.append([1, "ИЛ-10-0001/2025", "ООО Альфа", "9510/2025"])
            wb.save(path)
            rows = read_archive_journal(path)
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]["archive_index"], "ИЛ-10-0001/2025")
            self.assertEqual(rows[0]["protocol"], "9510/2025")

if __name__ == "__main__":
    unittest.main()
