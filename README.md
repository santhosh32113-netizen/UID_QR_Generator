# KUIN-G Generator

Offline asset register and QR generator for Windows deployment.

## Windows build

The GitHub Actions workflow builds the application on a Windows runner. Run it manually from the **Actions** tab, or push a version tag such as `v1.0.0`.

The workflow publishes the `KUIN-G-Windows` artifact. Download it and extract it once on the standalone Windows PC or pendrive, then run `KUIN-G.exe` from the extracted folder. Keep the complete extracted `KUIN-G-Windows` folder together; do not move only the EXE. The package keeps `dashboard\`, `input\`, `output\`, and `qr_codes\` beside the EXE, and saves to paths relative to that folder:

```text
<drive>:\KUIN-G-Windows\input\Sample.xlsx
<drive>:\KUIN-G-Windows\output\
<drive>:\KUIN-G-Windows\qr_codes\
```

The USB drive must not be write-protected. Do not use a hard-coded `C:\Users\user\Desktop\...` path.

## Local development

```text
python UID.py
```

Open `http://127.0.0.1:8765/index.html` in a browser.

## Outputs

- `input/Sample.xlsx`: asset data
- `output/master_register.xlsx`: Drone ID and KUIN-G mapping
- `output/distributable_register.xlsx`: KUIN-G and embedded QR images
- `output/qr_register.csv`: QR tracking register
- `qr_codes/<KUIN-G>.png`: printable QR image files
- `qr_codes/<KUIN-G>.svg`: path-based scalable QR files for SolidWorks and print workflows
- `qr_stl_backup/<KUIN-G>.stl`: raised 3D QR mesh backup for SolidWorks import
- `archive/deleted_records/<timestamp>_<KUIN-G>/`: deleted record snapshot and QR assets

The QR payload and filename are both KUIN-G. No separate UID is generated.

The database is preserved during EXE upgrades on Windows under `%LOCALAPPDATA%\KUIN-G`. Deleted records are archived before removal and are not silently discarded.
