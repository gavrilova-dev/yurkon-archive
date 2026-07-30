from app.services.excel_service import ExcelService


excel_path = r"C:\ЮРКОН\Административка\Реестр\Журнал_архивов.xlsx"


excel = ExcelService(excel_path)

excel.load()


protocol = "7609/2024"


record = excel.find_by_protocol(protocol)


print(record)


excel.mark_scanned(
    record.excel_row
)


excel.save()


print("Строка отмечена")