import tempfile
import unittest
from pathlib import Path
from app.file_manager import FileManager

class FileManagerTests(unittest.TestCase):
    def test_moves_ready_pdf_and_sanitizes_archive_index(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            incoming, ready = root / "new", root / "ready"
            incoming.mkdir()
            source = incoming / "scan.pdf"
            source.write_bytes(b"%PDF-test")
            manager = FileManager(incoming, ready)
            destination = manager.move_pdf(source, "ИЛ-10-0001/2025")
            self.assertTrue(destination.exists())
            self.assertEqual(destination.name, "ИЛ-10-0001-2025.pdf")
            self.assertFalse(source.exists())

if __name__ == "__main__":
    unittest.main()
