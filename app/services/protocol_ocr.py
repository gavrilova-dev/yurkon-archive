from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path


PROTOCOL_PATTERN = re.compile(r"(?<!\d)(\d{1,6})\s*([-/])\s*(\d{4})(?!\d)")
PROTOCOL_LABEL_PATTERN = re.compile(r"(?:№|N[oо]?\.?)\s*протокол(?:а)?", re.IGNORECASE)


def normalize_protocol(value: str) -> str:
    return re.sub(r"\s+", "", str(value or "")).replace("ё", "е").casefold()


def extract_protocol_candidates(text: str) -> list[str]:
    """Return printed protocol-like number candidates in their order of appearance."""
    source = text or ""
    candidates = []
    seen = set()

    # Ignore dates and identifiers elsewhere on the page. Accept a number
    # only when it follows the explicit "№ протокола" label.
    for label in PROTOCOL_LABEL_PATTERN.finditer(source):
        nearby_text = source[label.end():label.end() + 160]
        match = PROTOCOL_PATTERN.search(nearby_text)
        if match:
            value = f"{match.group(1)}{match.group(2)}{match.group(3)}"
            normalized = normalize_protocol(value)
            if normalized not in seen:
                seen.add(normalized)
                candidates.append(value)
    return candidates


def _tesseract_paths() -> tuple[Path, Path]:
    if getattr(sys, "frozen", False):
        root = Path(sys._MEIPASS) / "tesseract"
    else:
        root = Path(__file__).resolve().parents[2] / "build" / "tesseract"
    executable = root / "tesseract.exe"
    tessdata = root / "tessdata"
    if not executable.is_file():
        raise FileNotFoundError("В сборке не найден модуль распознавания Tesseract.")
    if not (tessdata / "rus.traineddata").is_file():
        raise FileNotFoundError("Не найден языковой файл русского OCR.")
    return executable, tessdata


def recognize_protocol_candidates(pdf_path: str | Path) -> list[str]:
    """Read the first PDF page and OCR it if it contains no selectable text."""
    import fitz

    with fitz.open(str(pdf_path)) as document:
        if not document:
            return []
        page = document.load_page(0)
        candidates = extract_protocol_candidates(page.get_text("text"))
        if candidates:
            return candidates

        executable, tessdata = _tesseract_paths()
        pixmap = page.get_pixmap(matrix=fitz.Matrix(3, 3), colorspace=fitz.csGRAY, alpha=False)
        result = subprocess.run(
            [
                str(executable), "stdin", "stdout",
                "--tessdata-dir", str(tessdata),
                "-l", "rus+eng", "--oem", "1", "--psm", "6",
            ],
            input=pixmap.tobytes("png"),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=45,
            check=False,
        )
        if result.returncode != 0:
            message = result.stderr.decode("utf-8", errors="replace").strip()
            raise RuntimeError(message or "Tesseract не смог распознать страницу.")
        return extract_protocol_candidates(result.stdout.decode("utf-8", errors="replace"))
