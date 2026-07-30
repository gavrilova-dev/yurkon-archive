from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path


class PDFHandler(FileSystemEventHandler):

    def __init__(self, callback):

        self.callback = callback


    def on_created(self, event):

        path = Path(event.src_path)


        if path.suffix.lower() == ".pdf":

            self.callback()



class FileWatcher:


    def __init__(
        self,
        folder,
        callback
    ):

        self.folder = folder
        self.callback = callback

        self.observer = Observer()



    def start(self):

        handler = PDFHandler(
            self.callback
        )


        self.observer.schedule(
            handler,
            self.folder,
            recursive=False
        )


        self.observer.start()