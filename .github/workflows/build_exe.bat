@echo off
REM Builds collections_report.exe (command-prompt app) from collections_report.py
REM Requires Python 3.9+ installed and on PATH.
cd /d "%~dp0"
python -m pip install --upgrade pip
python -m pip install pyinstaller pandas numpy openpyxl tqdm
python -m PyInstaller --onefile --console --name collections_report collections_report.py
echo.
echo Done. Your executable is in the "dist" folder: dist\collections_report.exe
pause
