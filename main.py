import sys
from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from app.window import MainWindow


if sys.platform == "win32":
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("YURKON.Archive")

icon_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
app = QApplication(sys.argv)
app.setWindowIcon(QIcon(str(icon_root / "icon.ico")))

window = MainWindow()
window.show()

sys.exit(app.exec())
