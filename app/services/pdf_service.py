from pathlib import Path
import fitz  # PyMuPDF
class PDFService:


    def wait_until_ready(self, pdf_path: Path):

        last_size = -1

        while True:

            current_size = pdf_path.stat().st_size

            if current_size == last_size:
                break

            last_size = current_size


    def get_first_page_image(
        self,
        pdf_path: Path
    ):

        self.wait_until_ready(pdf_path)


        document = fitz.open(pdf_path)

        page = document[0]


        pix = page.get_pixmap(
            dpi=300
        )


        image_path = pdf_path.with_suffix(
            ".png"
        )


        pix.save(
            image_path
        )


        document.close()


        return image_path