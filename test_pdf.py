from pathlib import Path

from app.services.pdf_service import PDFService


pdf = Path(
    r"C:\ЮРКОН\Административка\Сканы_новые\doc04476720260730130359.pdf"
)


service = PDFService()


image = service.get_first_page_image(pdf)


print("Создано изображение:")
print(image)