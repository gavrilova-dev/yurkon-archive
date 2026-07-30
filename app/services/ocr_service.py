import re
from pathlib import Path

import pytesseract
from PIL import Image


class OCRService:

    def __init__(self):

        pytesseract.pytesseract.tesseract_cmd = (
            r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        )


    def extract_text(self, image_path: Path):

        image = Image.open(image_path)

        text = pytesseract.image_to_string(
            image,
            lang="rus+eng"
        )

        return text


    def find_protocol(self, text: str):

        pattern = r"\b\d{4,5}/\d{4}\b"

        result = re.search(
            pattern,
            text
        )

        if result:
            return result.group(0)

        return None