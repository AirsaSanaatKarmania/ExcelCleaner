@echo off
chcp 65001 >nul
echo ===== Building ExcelCleaner.exe =====
py -m pip install --upgrade pip
py -m pip install -r requirements.txt pyinstaller
py -m PyInstaller --onefile --windowed --clean --collect-submodules tksheet --name ExcelCleaner excel_cleaner.py
echo.
if exist dist\ExcelCleaner.exe (
    echo SUCCESS: dist\ExcelCleaner.exe
) else (
    echo FAILED - see messages above
)
pause
