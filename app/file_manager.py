from pathlib import Path
import time


class FileManager:
    def __init__(self, incoming_folder, ready_folder):
        self.incoming_folder = Path(incoming_folder)
        self.ready_folder = Path(ready_folder)
        self.incoming_folder.mkdir(parents=True, exist_ok=True)
        self.ready_folder.mkdir(parents=True, exist_ok=True)

    def get_new_files(self):
        return sorted(self.incoming_folder.glob("*.pdf"), key=lambda p: p.name.casefold())

    @staticmethod
    def _wait_until_ready(pdf_path: Path, timeout: float = 30.0) -> None:
        deadline = time.monotonic() + timeout
        previous_size = None
        stable_checks = 0
        while time.monotonic() < deadline:
            size = pdf_path.stat().st_size
            if size > 0 and size == previous_size:
                stable_checks += 1
                if stable_checks >= 2:
                    return
            else:
                stable_checks = 0
            previous_size = size
            time.sleep(0.5)
        raise TimeoutError(f"Файл ещё записывается или пустой: {pdf_path.name}")

    def move_pdf(self, pdf_path, index):
        pdf_path = Path(pdf_path)
        safe_index = str(index).replace("/", "-").replace("\\", "-").strip()
        destination = self.ready_folder / f"{safe_index}.pdf"
        if destination.exists():
            raise FileExistsError(f"Файл уже существует: {destination.name}")
        self._wait_until_ready(pdf_path)
        last_error = None
        for _ in range(10):
            try:
                pdf_path.rename(destination)
                return destination
            except PermissionError as exc:
                last_error = exc
                time.sleep(0.5)
        raise PermissionError(f"Не удалось переместить {pdf_path.name}: {last_error}")
