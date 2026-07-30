from app.services.file_watcher import FileWatcher


folder = r"C:\ЮРКОН\Административка\Сканы_новые"


def new_file(path):

    print("Новый файл:")
    print(path)


watcher = FileWatcher(folder)

watcher.start(new_file)


print("Слежение запущено...")


try:

    while True:
        pass

except KeyboardInterrupt:

    watcher.stop()