@echo off
setlocal
cd /d %~dp0
if not exist .venv\Scripts\python.exe (
  py -3.12 -m venv .venv
)
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
pyinstaller --noconfirm --clean "YURKON Archive.spec"
echo.
echo Build complete: dist\YURKON Archive.exe
pause
