from pathlib import Path

from app.services.ocr_service import OCRService


image = Path(
    r"C:\ЮРКОН\Административка\Сканы_новые\doc04476720260730130359.png"
)


ocr = OCRService()


text = ocr.extract_text(
    image
)


print("ТЕКСТ:")
print("----------------")
print(text)
print("----------------")


protocol = ocr.find_protocol(
    text
)


print("Номер протокола:")
print(protocol)