@echo off
setlocal
call venv\Scripts\activate
python -m pip install -r requirements.txt
pyinstaller --noconfirm --clean --onedir --name KUIN-G ^
  --add-data "dashboard;dashboard" ^
  --add-data "input;input" ^
  --add-data "output;output" ^
  --add-data "qr_codes;qr_codes" ^
  UID.py
echo.
echo Build complete: dist\KUIN-G\KUIN-G.exe
pause