import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook, load_workbook
from app.services.report_exporter import export_report

class ReportExporterTests(unittest.TestCase):
    def test_export_keeps_all_sheets_and_does_not_write_archive_index(self):
        with tempfile.TemporaryDirectory() as folder:
            template = Path(folder) / "template.xlsx"
            output = Path(folder) / "out.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "2026 Реестр"
            ws.append(["№","Исполнитель","Месяц","Партия","№ документа","Название конторы",
                       "Кол-во рабочих мест","Первичка: к получению (по 20 ₽)",
                       "Первичка: моя доля (по 10 ₽)","Админ. работа выполнена (дата)",
                       "Админ. сумма (100 ₽)","Столбец1"])
            wb.create_sheet("2026 Сводка по партиям")
            yearly = wb.create_sheet("2026 Годовой отчет")
            yearly["A24"] = "Статус оплаты (по партиям)"
            yearly["A26"] = "Первичка: оплачено"
            yearly["B26"] = "=SUMIF(Свод2026[Первичка: оплачено (дата)],\"<>\",Свод2026[Первичка: сумма к получению (20 ₽)])"
            wb.create_sheet("2026 Отчёт руководителю")
            wb.create_sheet("Инструкция")
            wb.create_sheet("Списки")
            wb.save(template)
            records = [{
                "executor":"Аня","month":"октябрь","protocol":"1234/2026",
                "organization":"ООО Тест","workplaces":5,"admin_done_date":None
            }]
            batches = [{
                "executor":"Аня","month":"октябрь","primary_given_date":None,
                "primary_returned_date":None,"primary_report_date":None,"primary_paid_date":None,
                "admin_report_date":None,"admin_paid_date":None,"comment":""
            }]
            export_report(template, output, records, batches)
            result = load_workbook(output, data_only=False)
            self.assertEqual(result.sheetnames, wb.sheetnames)
            self.assertEqual(result["2026 Реестр"]["E2"].value, "1234/2026")
            self.assertEqual(result["2026 Реестр"]["F2"].value, "ООО Тест")
            self.assertEqual(result["2026 Реестр"]["H2"].value, '=IF(G2="","",G2*20)')
            self.assertNotIn("Индекс дела", [result["2026 Реестр"].cell(1,c).value for c in range(1,13)])
            self.assertEqual(result["2026 Годовой отчет"]["A24"].value, "Статус оплаты (по партиям)")
            self.assertEqual(result["2026 Годовой отчет"]["A26"].value, "Первичка: оплачено")
            self.assertEqual(result["2026 Отчёт руководителю"]["B3"].value, "Аня - октябрь")
            self.assertEqual(len(result["2026 Отчёт руководителю"].data_validations.dataValidation), 1)

if __name__ == "__main__":
    unittest.main()
