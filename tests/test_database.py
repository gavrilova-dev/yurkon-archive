import tempfile
import unittest
from pathlib import Path
from app.database import Database
class DatabaseTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.db = Database(Path(self.tmp.name) / "test.sqlite3")
    def tearDown(self):
        self.db.close()
        self.tmp.cleanup()
    def test_archive_import_is_idempotent_and_preserves_same_protocol_for_other_orgs(self):
        rows = [
            {"archive_index":"ИЛ-10-0001/2025","organization":"ООО Альфа","protocol":"9510/2025","source_row":4},
            {"archive_index":"ИЛ-10-0002/2025","organization":"ООО Бета","protocol":"9510/2025","source_row":5},
        ]
        result = self.db.import_archive_rows(rows)
        self.assertEqual(result["added"], 2)
        self.assertEqual(len(self.db.find_archive("9510/2025")), 2)
        result = self.db.import_archive_rows(rows)
        self.assertEqual(result["added"], 0)
        self.assertEqual(result["updated"], 2)
    def test_calculation_fields(self):
        self.db.add_record({"executor":"Аня","month":"октябрь","protocol":"1234/2026",
                            "organization":"ООО Тест","workplaces":7,"admin_done_date":"2026-10-01"})
        row = self.db.list_records()[0]
        self.assertEqual(row["workplaces"] * 20, 140)
        self.assertEqual(row["workplaces"] * 10, 70)
        self.assertEqual(100 if row["admin_done_date"] else 0, 100)
if __name__ == "__main__":
    unittest.main()
