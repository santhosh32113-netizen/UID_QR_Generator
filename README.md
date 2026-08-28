# KUIN-G Generator

Offline asset register and QR generator for Windows deployment.

## Windows build

The GitHub Actions workflow builds the application on a Windows runner. Run it manually from the **Actions** tab, or push a version tag such as `v1.0.0`.

The workflow publishes the `KUIN-G-Windows` artifact. Download it and extract it once on the standalone Windows PC, then run `KUIN-G.exe` from a writable folder such as `C:\KUIN-G`. Keep the complete extracted `KUIN-G` folder together; do not move only the EXE. The package keeps `dashboard\`, `input\`, `output\`, and `qr_codes\` beside the EXE; the application saves data in `input\Sample.xlsx`, `output\`, and `qr_codes\`.

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
- `qr_codes/<KUIN-G>.svg`: scalable QR files for SolidWorks and print workflows

The QR payload and filename are both KUIN-G. No separate UID is generated.
