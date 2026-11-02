from pathlib import Path
import os
INCOMING_FOLDER = r"C:\ЮРКОН\Административка\Сканы_новые"
READY_FOLDER = r"C:\ЮРКОН\Административка\Сканы_готовые"
ARCHIVE_JOURNAL_DEFAULT = r"C:\ЮРКОН\Административка\Реестр\Журнал_архивов.xlsx"
APP_DATA = Path(os.environ.get("APPDATA", Path.home())) / "YURKON Archive"
DATABASE_PATH = APP_DATA / "yurkon_archive.sqlite3"
