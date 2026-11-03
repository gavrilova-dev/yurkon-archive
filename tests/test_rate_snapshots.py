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

    def test_missing_snapshot_values_are_backfilled_from_existing_settings(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "archive.sqlite3"
            db = Database(path)
            rid = db.add_record({"executor":"Старый","month":"апрель","year":2026,
                "protocol":"3/2026","organization":"ООО Архив","workplaces":2})
            with db.connection:
                db.connection.execute(
                    "UPDATE primary_records SET primary_rate=NULL, primary_share_rate=NULL, admin_rate=NULL WHERE id=?",
                    (rid,),
                )
                db.connection.execute(
                    "UPDATE settings SET value='27' WHERE key='primary_rate'"
                )
                db.connection.execute(
                    "UPDATE settings SET value='13' WHERE key='primary_share_rate'"
                )
                db.connection.execute(
                    "UPDATE settings SET value='125' WHERE key='admin_rate'"
                )
            db.close()
            migrated = Database(path)
            try:
                record = dict(migrated.get_record(rid))
                self.assertEqual((record["primary_rate"],record["primary_share_rate"],record["admin_rate"]),
                                 (27,13,125))
            finally:
                migrated.close()

    def test_mark_scanned_reports_when_record_is_missing(self):
        with tempfile.TemporaryDirectory() as folder:
            db = Database(Path(folder) / "archive.sqlite3")
            try:
                result = db.mark_scanned("404/2026", "Неизвестная организация",
                                         "ready/404.pdf", "2026-11-03")
                self.assertFalse(result)
            finally:
                db.close()

    def test_mark_scanned_returns_true_and_saves_status_for_matching_record(self):
        with tempfile.TemporaryDirectory() as folder:
            db = Database(Path(folder) / "archive.sqlite3")
            try:
                rid = db.add_record({"executor":"Тест","month":"ноябрь","year":2026,
                    "protocol":"55/2026","organization":"ООО Проверка","workplaces":1})
                result = db.mark_scanned("55/2026", "ООО Проверка",
                                         "ready/55-2026.pdf", "2026-11-03")
                self.assertTrue(result)
                record = dict(db.get_record(rid))
                self.assertEqual(record["scanned_date"], "2026-11-03")
                self.assertEqual(record["pdf_path"], "ready/55-2026.pdf")
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
