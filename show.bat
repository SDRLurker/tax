@echo off
if "%~1"=="" (
    pyinstaller --onefile --icon=icon.ico --exclude-module sqlite3 --exclude-module html --exclude-module unittest --exclude-module matplotlib --exclude-module tkinter --exclude-module pdb --exclude-module ftplib --exclude-module smtplib --exclude-module poplib --hidden-import=openpyxl --hidden-import=openpyxl.cell._writer --add-data "./*;." show.py
) else (
    pyinstaller --onefile --upx-dir="%~1" --icon=icon.ico --exclude-module sqlite3 --exclude-module html --exclude-module unittest --exclude-module matplotlib --exclude-module tkinter --exclude-module pdb --exclude-module ftplib --exclude-module smtplib --exclude-module poplib --hidden-import=openpyxl --hidden-import=openpyxl.cell._writer --add-data "./*;." show.py
)
