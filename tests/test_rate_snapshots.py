import tempfile
import unittest
from pathlib import Path
from app.database import Database


class RateSnapshotTests(unittest.TestCase):
    def test_existing_records_keep_rates_and_new_records_use_updated_rates(self):
        with tempfile.TemporaryDirectory() as folder:
            db = Database(Path(folder) / "archive.sqlite3")
            try:
                old_id = db.add_record({"executor":"Старый","month":"январь","year":2026,
                    "protocol":"1/2026","organization":"ООО Старое","workplaces":5})
                db.save_rates({"primary_rate":30,"primary_share_rate":15,"admin_rate":150})
                new_id = db.add_record({"executor":"Новый","month":"февраль","year":2027,
                    "protocol":"1/2027","organization":"ООО Новое","workplaces":4})
                old, new = dict(db.get_record(old_id)), dict(db.get_record(new_id))
                self.assertEqual((old["primary_rate"],old["primary_share_rate"],old["admin_rate"]),(20,10,100))
                self.assertEqual((new["primary_rate"],new["primary_share_rate"],new["admin_rate"]),(30,15,150))
            finally:
                db.close()

    def test_custom_admin_amount_remains_override(self):
        with tempfile.TemporaryDirectory() as folder:
            db = Database(Path(folder) / "archive.sqlite3")
            try:
                rid = db.add_record({"executor":"Лилия","month":"март","year":2026,
                    "protocol":"2/2026","organization":"ООО Тест","workplaces":1,"admin_amount":350})
                db.save_rates({"primary_rate":40,"primary_share_rate":20,"admin_rate":200})
                record = dict(db.get_record(rid))
                self.assertEqual(record["admin_amount"],350)
                self.assertEqual(record["admin_rate"],100)
            finally:
                db.close()


if __name__ == "__main__":
    unittest.main()
