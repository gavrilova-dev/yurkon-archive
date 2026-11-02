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
            wb.create_sheet("2026 Годовой отчет")
            wb.create_sheet("2026 Отчёт руководителю")
            wb.create_sheet("Инструкция")
            wb.create_sheet("Списки")
            wb.save(template)
            records = [{
                "executor":"Аня","month":"октябрь","protocol":"1234/2026",
                "organization":"ООО Тест","workplaces":5,"admin_done_date":None
            }]
            export_report(template, output, records, [])
            result = load_workbook(output, data_only=False)
            self.assertEqual(result.sheetnames, wb.sheetnames)
            self.assertEqual(result["2026 Реестр"]["E2"].value, "1234/2026")
            self.assertEqual(result["2026 Реестр"]["F2"].value, "ООО Тест")
            self.assertEqual(result["2026 Реестр"]["H2"].value, '=IF(G2="","",G2*20)')
            self.assertNotIn("Индекс дела", [result["2026 Реестр"].cell(1,c).value for c in range(1,13)])

if __name__ == "__main__":
    unittest.main()
