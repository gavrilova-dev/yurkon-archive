import tempfile
import unittest
from pathlib import Path
from openpyxl import load_workbook
from app.services.report_exporter import export_report


class ReportExporterTests(unittest.TestCase):
    def test_year_is_part_of_party_key_and_values_are_calculated(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "summary.xlsx"
            records = [
                {"year": 2026, "executor": "Тест", "month": "сентябрь",
                 "protocol": "1/2026", "organization": "ООО 2026", "workplaces": 5,
                 "admin_done_date": "2026-09-01"},
                {"year": 2027, "executor": "Тест", "month": "сентябрь",
                 "protocol": "1/2027", "organization": "ООО 2027", "workplaces": 8,
                 "admin_done_date": "2027-09-01"},
            ]
            batches = [
                {"executor": "Тест", "month": "сентябрь", "year": 2026},
                {"executor": "Тест", "month": "сентябрь", "year": 2027},
            ]
            rates = {"primary_rate": 20, "primary_share_rate": 10, "admin_rate": 100}
            export_report(None, output, records, batches, 2026,
                          report_type="summary", rates=rates)
            wb = load_workbook(output, data_only=False)
            self.assertEqual(wb.sheetnames, ["2026 Сводка по партиям"])
            ws = wb.active
            self.assertEqual(ws["A1"].value, "Партия")
            self.assertEqual(ws["A2"].value, "Тест - сентябрь")
            self.assertEqual(ws["B2"].value, 1)
            self.assertEqual(ws["E2"].value, 100)
            self.assertEqual(ws["F2"].value, 50)
            self.assertEqual(ws["I2"].value, 1)
            self.assertEqual(ws["J2"].value, 100)
            self.assertEqual(ws["A3"].value, "ИТОГО ЗА ГОД")
            self.assertTrue(ws["A1"].font.bold)
            self.assertTrue(ws["A3"].font.bold)
            for sheet in wb.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        self.assertFalse(isinstance(cell.value, str) and cell.value.startswith("="))
            wb.close()

    def test_manager_admin_export_uses_custom_amount_and_is_standalone(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "manager.xlsx"
            records = [
                {"year": 2027, "executor": "Лилия", "month": "январь",
                 "protocol": "1/2027", "organization": "ООО Тест", "workplaces": 3,
                 "admin_done_date": "2027-01-10", "admin_amount": 350},
            ]
            batches = [{"executor": "Лилия", "month": "январь", "year": 2027}]
            rates = {"primary_rate": 30, "primary_share_rate": 15, "admin_rate": 120}
            export_report(None, output, records, batches, 2027, report_type="manager",
                          party="Лилия - январь", manager_kind="admin", rates=rates)
            wb = load_workbook(output, data_only=False)
            self.assertEqual(wb.sheetnames, ["2027 Отчёт руководителю"])
            ws = wb.active
            self.assertEqual(ws["A1"].value, "№ протокола")
            self.assertEqual(ws["A2"].value, "1/2027")
            self.assertEqual(ws["D2"].value, 350)
            self.assertEqual(ws["D3"].value, 350)
            self.assertTrue(ws["A1"].font.bold)
            self.assertTrue(ws["A3"].font.bold)
            self.assertFalse(any(cell.data_type == "f" for row in ws.iter_rows() for cell in row))
            wb.close()

    def test_primary_export_uses_configured_rate(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "primary.xlsx"
            records = [{"year": 2027, "executor": "Лилия", "month": "январь",
                        "protocol": "1/2027", "organization": "ООО Тест", "workplaces": 3}]
            rates = {"primary_rate": 35, "primary_share_rate": 12, "admin_rate": 100}
            export_report(None, output, records, [], 2027, report_type="manager",
                          party="Лилия - январь", manager_kind="primary", rates=rates)
            ws = load_workbook(output, data_only=True).active
            self.assertEqual(ws["D2"].value, 105)
            self.assertEqual(ws["D3"].value, 105)

    def test_annual_report_has_three_separate_tables_and_bold_headers(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "annual.xlsx"
            records = [{"year": 2027, "executor": "Лилия", "month": "январь",
                        "protocol": "1/2027", "organization": "ООО Тест", "workplaces": 3,
                        "admin_done_date": "2027-01-10", "admin_amount": None}]
            batches = [{"executor": "Лилия", "month": "январь", "year": 2027,
                        "primary_paid_date": "2027-02-01"}]
            rates = {"primary_rate": 20, "primary_share_rate": 10, "admin_rate": 100}
            export_report(None, output, records, batches, 2027, report_type="annual", rates=rates)
            ws = load_workbook(output, data_only=False).active
            self.assertEqual(ws["A3"].value, "Месяц")
            self.assertEqual(ws["A4"].value, "Январь")
            self.assertEqual(ws["A20"].value, "Исполнитель")
            self.assertEqual(ws["A21"].value, "Лилия")
            self.assertEqual(ws["A26"].value, "Показатель")
            self.assertEqual(ws["A27"].value, "Первичка оплачена")
            self.assertTrue(ws["A3"].font.bold)
            self.assertTrue(ws["A20"].font.bold)
            self.assertTrue(ws["A26"].font.bold)
            self.assertEqual(ws["F4"].value, 100)
            self.assertEqual(ws["B27"].value, 60)
            self.assertFalse(any(cell.data_type == "f" for row in ws.iter_rows() for cell in row))
            wb = ws.parent
            self.assertEqual(len(wb.sheetnames), 1)
            wb.close()


if __name__ == "__main__":
    unittest.main()
