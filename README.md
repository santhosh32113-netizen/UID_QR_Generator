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

The database is preserved during EXE upgrades under the configured data root. By default this is the folder beside the EXE (and the repository root during local development); set `KUIN_G_DATA_ROOT` to a separate, stable folder for production deployments. Deleted records are archived before removal and are not silently discarded.

## Upgrades and version control

The application version is stored in `VERSION`. Release builds must use a matching Git tag, for example `VERSION` `1.0.1` with tag `v1.0.1`. Keep the source in a Git repository and deploy only tagged build artifacts.

Before the first startup and before any version change, the application creates an atomic ZIP snapshot in `backups/` under the data root. The snapshot includes `input/Sample.xlsx`, generated registers, QR files, uploaded images, the deletion archive, `passwords.json`, and `secret.key`. Startup stops if the snapshot cannot be created or the version state cannot be read, preventing an upgrade from proceeding without a recoverable copy. Copy the newest `backups/kuin-g-*.zip` to separate storage as part of regular operations; it is the recovery source if the application folder or drive is lost. To recover, stop the application, extract the archive into the configured data root, and start the matching or newer release.

## KUIN-G generation algorithm

KUIN-G is derived from the Drone ID (the primary key) using a 256-bit secret HMAC key:

1. `HMAC-SHA256(secret_key, Drone_ID)` produces a 256-bit digest.
2. The digest is Base32 encoded using the alphanumeric alphabet `A-Z2-7`.
3. The first 16 characters are used as the KUIN-G.

The resulting KUIN-G is exactly 16 characters and is case-insensitive for storage/display. The secret key is 32 bytes (256 bits) and is stored in `secret.key` under the data root on first run. For production deployment, protect this key and back it up securely; losing it means the same Drone ID cannot reproduce the same KUIN-G. You can instead supply the key through the `KUIN_G_SECRET_KEY_B64` environment variable.

**Important:** truncating the 256-bit HMAC digest to 16 Base32 characters reduces the identifier to 80 bits of output. The HMAC itself remains SHA-256/256-bit; the KUIN-G representation is an 80-bit truncated identifier.
