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
        record = self.db.list_records()[0]
        self.assertEqual(record["scanned_date"], "2026-11-02")
        self.assertEqual(record["assembled_date"], "2026-11-02")

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

    def test_existing_legacy_duplicate_can_still_be_updated_without_changing_protocol(self):
        record_id = self.db.add_record({
            "executor":"Анна","month":"февраль","protocol":"9356/2025",
            "organization":"ООО Альфа","workplaces":1
        })
        self.db.connection.execute("""
            INSERT INTO primary_records(executor,month,protocol,organization,workplaces)
            VALUES(?,?,?,?,?)
        """, ("Мария","март","9356/2025","ООО Бета",2))
        self.db.connection.commit()
        record = dict(self.db.get_record(record_id))
        record["assembled_date"] = "2026-11-02"
        self.db.update_record(record_id, record)
        self.assertEqual(self.db.get_record(record_id)["assembled_date"], "2026-11-02")

    def test_same_executor_and_month_can_exist_in_different_years(self):
        self.db.add_record({
            "executor":"Лилия","month":"январь","year":2026,
            "protocol":"100/2026","organization":"ООО Альфа","workplaces":1
        })
        self.db.add_record({
            "executor":"Лилия","month":"январь","year":2027,
            "protocol":"100/2027","organization":"ООО Бета","workplaces":2
        })
        batches = self.db.list_batches()
        self.assertEqual(len(batches), 2)
        self.assertEqual({row["year"] for row in batches}, {2026, 2027})

    def test_scanned_status_also_marks_admin_work_done(self):
        record_id = self.db.add_record({
            "executor":"Аня","month":"ноябрь","year":2026,
            "protocol":"1001/2026","organization":"ООО Тест","workplaces":1
        })
        self.db.update_status(record_id, "scanned_date", "2026-11-02")
        record = self.db.get_record(record_id)
        self.assertEqual(record["scanned_date"], "2026-11-02")
        self.assertEqual(record["assembled_date"], "2026-11-02")
        self.assertEqual(record["admin_done_date"], "2026-11-02")
        self.db.update_status(record_id, "scanned_date", None)
        record = self.db.get_record(record_id)
        self.assertIsNone(record["scanned_date"])
        self.assertEqual(record["assembled_date"], "2026-11-02")

    def test_archive_index_lookup_uses_protocol_not_organization_name(self):
        self.db.add_record({
            "executor":"Анна","month":"ноябрь","year":2026,
            "protocol":"555/2026","organization":"Название организации изменилось","workplaces":1
        })
        self.db.import_archive_rows([{
            "archive_index":"ИЛ-10-0555/2026","organization":"Старое название",
            "protocol":"555/2026","source_row":4
        }])
        record = self.db.list_records()[0]
        self.assertEqual(record["archive_index"], "ИЛ-10-0555/2026")


    def test_legacy_import_reads_only_five_fields_and_detects_year(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "source.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "2027 Реестр"
            ws.append(["№","Исполнитель","Месяц","Партия","№ документа","Название конторы","Кол-во рабочих мест",
                       "Первичка: к получению (по 20 ₽)","Первичка: моя доля (по 10 ₽)",
                       "Админ. работа выполнена (дата)","Админ. сумма (100 ₽)"])
            ws.append([1,"Лилия","январь",None,"10/2027","ООО Тест",4,999,999,"2027-01-10",100])
            wb.save(path)
            result = self.db.import_legacy_workbook(path)
            self.assertEqual(result["imported"], 1)
            record = self.db.list_records()[0]
            self.assertEqual(record["year"], 2027)
            self.assertEqual(record["workplaces"], 4)
            self.assertIsNone(record["admin_done_date"])
            self.assertIsNone(record["assembled_date"])
            self.assertIsNone(record["scanned_date"])


    def test_custom_admin_amount_and_configurable_rates_are_saved(self):
        self.db.save_rates({"primary_rate": 30, "primary_share_rate": 15, "admin_rate": 125})
        record_id = self.db.add_record({
            "executor": "Тест", "month": "сентябрь", "year": 2027,
            "protocol": "77/2027", "organization": "ООО Большая", "workplaces": 4,
            "admin_amount": 350, "admin_done_date": "2027-09-01"
        })
        record = self.db.get_record(record_id)
        self.assertEqual(record["admin_amount"], 350)
        self.assertEqual(self.db.get_rates(), {
            "primary_rate": 30, "primary_share_rate": 15, "admin_rate": 125
        })

    def test_clear_archive_records_preserves_primary_records_batches_and_settings(self):
        record_id = self.db.add_record({
            "executor": "Анна", "month": "ноябрь", "year": 2026,
            "protocol": "123/2026", "organization": "ООО Основная", "workplaces": 3
        })
        self.db.set_setting("custom_setting", "keep")
        self.db.import_archive_rows([{
            "archive_index": "ИЛ-10-0001/2026", "organization": "ООО Архив",
            "protocol": "123/2026", "source_row": 4
        }])
        deleted = self.db.clear_archive_records()
        self.assertEqual(deleted, 1)
        self.assertEqual(self.db.archive_count(), 0)
        self.assertEqual(len(self.db.list_records()), 1)
        self.assertEqual(self.db.get_record(record_id)["organization"], "ООО Основная")
        self.assertEqual(len(self.db.list_batches()), 1)
        self.assertEqual(self.db.get_setting("custom_setting"), "keep")

    def test_backup_restore_replaces_database_contents(self):
        with tempfile.TemporaryDirectory() as folder:
            backup = Path(folder) / "backup.sqlite3"
            self.db.add_record({
                "executor": "Аня", "month": "октябрь", "year": 2026,
                "protocol": "1/2026", "organization": "ООО Старое", "workplaces": 1
            })
            self.db.backup_to(backup)
            self.db.add_record({
                "executor": "Аня", "month": "ноябрь", "year": 2026,
                "protocol": "2/2026", "organization": "ООО Новое", "workplaces": 2
            })
            self.db.restore_from(backup)
            self.assertEqual(len(self.db.list_records()), 1)
            self.assertEqual(self.db.list_records()[0]["organization"], "ООО Старое")

    def test_calculation_fields(self):
        self.db.add_record({"executor":"Аня","month":"октябрь","protocol":"1234/2026",
                            "organization":"ООО Тест","workplaces":7,"admin_done_date":"2026-10-01"})
        row = self.db.list_records()[0]
        self.assertEqual(row["workplaces"] * 20, 140)
        self.assertEqual(row["workplaces"] * 10, 70)
        self.assertEqual(100 if row["admin_done_date"] else 0, 100)
if __name__ == "__main__":
    unittest.main()


class ProtocolOcrParsingTests(unittest.TestCase):
    def test_extracts_printed_protocol_numbers(self):
        from app.services.protocol_ocr import extract_protocol_candidates

        self.assertEqual(
            extract_protocol_candidates("Протокол № 7609/2024, второй: 81-2025"),
            ["7609/2024", "81-2025"],
        )

    def test_ignores_numbers_without_four_digit_year_and_duplicates(self):
        from app.services.protocol_ocr import extract_protocol_candidates

        self.assertEqual(
            extract_protocol_candidates("№ 123/24; 123/24; 4567/2026"),
            ["4567/2026"],
        )
