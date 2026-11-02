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
    def test_cyrillic_search_archive_match_and_scan_status(self):
        self.db.add_record({
            "executor":"Аня","month":"октябрь","protocol":"1234/2026",
            "organization":"ООО АЛЬФА","workplaces":2
        })
        self.db.import_archive_rows([{
            "archive_index":"ИЛ-10-0001/2026","organization":"ооо альфа",
            "protocol":"1234/2026","source_row":4
        }])
        rows = self.db.list_records("альфа")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["archive_index"], "ИЛ-10-0001/2026")
        self.db.mark_scanned("1234/2026", "ооо альфа", "C:/ready/test.pdf", "2026-11-02")
        self.assertEqual(self.db.list_records()[0]["scanned_date"], "2026-11-02")

    def test_archive_index_matches_when_organization_punctuation_differs(self):
        self.db.add_record({
            "executor":"Лилия","month":"февраль","protocol":"9356/2025",
            "organization":'ООО "ДРТ Групп"',"workplaces":15
        })
        self.db.import_archive_rows([{
            "archive_index":"ИЛ-10-0001/2025","organization":"ООО ДРТ Групп",
            "protocol":"9356/2025","source_row":4
        }])
        row = self.db.list_records()[0]
        self.assertEqual(row["archive_index"], "ИЛ-10-0001/2025")

    def test_duplicate_protocol_is_rejected_even_for_another_organization(self):
        self.db.add_record({
            "executor":"Анна","month":"февраль","protocol":"9356/2025",
            "organization":"ООО Альфа","workplaces":1
        })
        with self.assertRaisesRegex(ValueError, "уже есть в базе"):
            self.db.add_record({
                "executor":"Мария","month":"март","protocol":"9356/2025",
                "organization":"ООО Бета","workplaces":2
            })

    def test_calculation_fields(self):
        self.db.add_record({"executor":"Аня","month":"октябрь","protocol":"1234/2026",
                            "organization":"ООО Тест","workplaces":7,"admin_done_date":"2026-10-01"})
        row = self.db.list_records()[0]
        self.assertEqual(row["workplaces"] * 20, 140)
        self.assertEqual(row["workplaces"] * 10, 70)
        self.assertEqual(100 if row["admin_done_date"] else 0, 100)
if __name__ == "__main__":
    unittest.main()
