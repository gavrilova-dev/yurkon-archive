from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
class PDFHandler(FileSystemEventHandler):
    def __init__(self, callback):
        self.callback = callback
    def on_created(self, event):
        if not event.is_directory and Path(event.src_path).suffix.lower() == ".pdf":
            self.callback()
    def on_moved(self, event):
        if not event.is_directory and Path(event.dest_path).suffix.lower() == ".pdf":
            self.callback()
class FileWatcher:
    def __init__(self, folder, callback):
        self.folder = str(folder)
        self.callback = callback
        self.observer = Observer()
    def start(self):
        Path(self.folder).mkdir(parents=True, exist_ok=True)
        self.observer.schedule(PDFHandler(self.callback), self.folder, recursive=False)
        self.observer.start()
    def stop(self):
        if self.observer.is_alive():
            self.observer.stop()
            self.observer.join(timeout=3)
