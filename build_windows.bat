@echo off
setlocal
call venv\Scripts\activate
python -m pip install -r requirements.txt
pyinstaller --noconfirm --clean --console --onedir --name KUIN-G ^
  --hidden-import=tools.dashboard_server ^
  --hidden-import=tools.create_dashboard_data ^
  --hidden-import=src.generate_UIDS ^
  --add-data "dashboard;dashboard" ^
  --add-data "input;input" ^
  --add-data "output;output" ^
  --add-data "qr_codes;qr_codes" ^
  UID.py
echo.
echo Build complete: dist\KUIN-G\KUIN-G.exe
pause