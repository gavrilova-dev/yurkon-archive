import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook, load_workbook
from app.services.report_exporter import export_report


class ReportExporterTests(unittest.TestCase):
    def _template(self, path):
        wb = Workbook()
        wb.active.title = "2026 Сводка по партиям"
        sw = wb.active
        sw["A1"] = "Сводка по партиям — 2026"
        sw["A4"] = "Партия"
        sw["B4"] = "Кол-во контор"
        sw["E4"] = "Первичка: сумма к получению (20 ₽)"
        sw["F4"] = "Первичка: моя доля (10 ₽)"
        sw["I4"] = "Админ: контор сделано"
        sw["J4"] = "Админ: сумма (100 ₽)"
        manager = wb.create_sheet("2026 Отчёт руководителю")
        manager["A1"] = "ОТЧЁТ ПО ПАРТИИ КОНТОР"
        manager["A5"] = "БЛОК 1. ПЕРВИЧНАЯ ДОКУМЕНТАЦИЯ"
        manager["A6"] = "№ документа"
        manager["B6"] = "Название конторы"
        manager["C6"] = "Кол-во мест"
        manager["D6"] = "Сумма (20 ₽/место)"
        manager["A139"] = "БЛОК 2. АДМИНИСТРАТИВНАЯ РАБОТА"
        manager["A140"] = "№ документа"
        manager["B140"] = "Название конторы"
        manager["C140"] = "Дата"
        manager["D140"] = "Сумма"
        yearly = wb.create_sheet("2026 Годовой отчет")
        yearly["A1"] = "ГОДОВОЙ ОТЧЁТ — 2026"
        yearly["A4"] = "Месяц"
        yearly["A20"] = "Исполнитель"
        yearly["A24"] = "Статус оплаты"
        yearly["A25"] = "Показатель"
        yearly["B25"] = "Сумма"
        wb.save(path)

    def test_reports_contain_calculated_values_not_formulas(self):
        with tempfile.TemporaryDirectory() as folder:
            template = Path(folder) / "template.xlsx"
            output = Path(folder) / "all.xlsx"
            self._template(template)
            records = [
                {"id": 1, "year": 2026, "executor": "Аня", "month": "октябрь",
                 "protocol": "1234/2026", "organization": "ООО Тест", "workplaces": 5,
                 "admin_done_date": "2026-10-01"},
                {"id": 2, "year": 2026, "executor": "Аня", "month": "октябрь",
                 "protocol": "1235/2026", "organization": "ООО Альфа", "workplaces": 2,
                 "admin_done_date": None},
            ]
            batches = [{"executor": "Аня", "month": "октябрь", "year": 2026,
                        "primary_given_date": None, "primary_returned_date": None,
                        "primary_report_date": None, "primary_paid_date": "2026-10-20",
                        "admin_report_date": None, "admin_paid_date": None, "comment": ""}]
            export_report(template, output, records, batches, 2026)
            result = load_workbook(output, data_only=False)
            sw = result["2026 Сводка по партиям"]
            self.assertEqual(sw["A5"].value, "Аня - октябрь")
            self.assertEqual(sw["B5"].value, 2)
            self.assertEqual(sw["E5"].value, 140)
            self.assertEqual(sw["F5"].value, 70)
            self.assertEqual(sw["I5"].value, 1)
            self.assertEqual(sw["J5"].value, 100)
            yearly = result["2026 Годовой отчет"]
            self.assertEqual(yearly["B14"].value, 2)
            self.assertEqual(yearly["D14"].value, 140)
            self.assertEqual(yearly["B26"].value, 140)
            manager = result["2026 Отчёт руководителю"]
            self.assertEqual(manager["A7"].value, "1234/2026")
            self.assertEqual(manager["D7"].value, 100)
            self.assertEqual(manager["B137"].value, "ИТОГО ПО ПЕРВИЧНОЙ ДОКУМЕНТАЦИИ:")
            for sheet in result.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        self.assertFalse(isinstance(cell.value, str) and cell.value.startswith("="),
                                         f"Formula remains in {sheet.title}!{cell.coordinate}")

    def test_each_export_is_a_separate_report_file(self):
        with tempfile.TemporaryDirectory() as folder:
            template = Path(folder) / "template.xlsx"
            output = Path(folder) / "manager.xlsx"
            self._template(template)
            records = [{"year": 2027, "executor": "Лилия", "month": "январь",
                        "protocol": "1/2027", "organization": "ООО Тест", "workplaces": 3,
                        "admin_done_date": "2027-01-10"}]
            batches = [{"executor": "Лилия", "month": "январь", "year": 2027}]
            export_report(template, output, records, batches, 2027,
                          report_type="manager", party="Лилия - январь", manager_kind="admin")
            result = load_workbook(output, data_only=False)
            self.assertEqual(result.sheetnames, ["2027 Отчёт руководителю"])
            ws = result.active
            self.assertEqual(ws["A141"].value, "1/2027")
            self.assertEqual(ws["D141"].value, 100)
            self.assertEqual(ws["C271"].value, 1)
            for row in ws.iter_rows():
                for cell in row:
                    self.assertFalse(isinstance(cell.value, str) and cell.value.startswith("="))


if __name__ == "__main__":
    unittest.main()
