from pathlib import Path
class FileManager:
    def __init__(self, incoming_folder, ready_folder):
        self.incoming_folder = Path(incoming_folder)
        self.ready_folder = Path(ready_folder)
        self.incoming_folder.mkdir(parents=True, exist_ok=True)
        self.ready_folder.mkdir(parents=True, exist_ok=True)
    def get_new_files(self):
        return sorted(self.incoming_folder.glob("*.pdf"), key=lambda p: p.name.casefold())
    def move_pdf(self, pdf_path, index):
        pdf_path = Path(pdf_path)
        safe_index = str(index).replace("/", "-").replace("\\", "-").strip()
        destination = self.ready_folder / f"{safe_index}.pdf"
        if destination.exists():
            raise FileExistsError(f"Файл уже существует: {destination.name}")
        pdf_path.rename(destination)
        return destination
