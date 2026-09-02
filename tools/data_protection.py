"""Version-aware snapshots for the live KUIN-G data directory."""

from __future__ import annotations

import datetime
import json
import os
import shutil
import tempfile
import zipfile
from pathlib import Path

STATE_FILE_NAME = "app_state.json"
BACKUP_DIR_NAME = "backups"
PERSISTED_PATHS = (
    Path("input") / "Sample.xlsx",
    Path("output"),
    Path("qr_codes"),
    Path("qr_stl_backup"),
    Path("dashboard") / "uploads",
    Path("archive"),
    Path("secret.key"),
    Path("passwords.json"),
)


def read_app_version(resource_root: Path) -> str:
    version_file = resource_root / "VERSION"
    version = version_file.read_text(encoding="ascii").strip() if version_file.exists() else "0.0.0-dev"
    if not version:
        raise ValueError(f"Application version file is empty: {version_file}")
    return version


def _atomic_write_json(path: Path, payload: dict[str, str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary_name = tempfile.mkstemp(prefix=f".{path.stem}-", suffix=".json", dir=str(path.parent))
    os.close(fd)
    temporary_path = Path(temporary_name)
    try:
        temporary_path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
        os.replace(temporary_path, path)
    finally:
        temporary_path.unlink(missing_ok=True)


def snapshot_data(data_root: Path, reason: str, version: str) -> Path | None:
    existing_paths = [relative for relative in PERSISTED_PATHS if (data_root / relative).exists()]
    if not existing_paths:
        return None

    backup_dir = data_root / BACKUP_DIR_NAME
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    archive_path = backup_dir / f"kuin-g-{timestamp}-{reason}-v{version}.zip"
    fd, temporary_name = tempfile.mkstemp(prefix=".kuin-g-backup-", suffix=".zip", dir=str(backup_dir))
    os.close(fd)
    temporary_path = Path(temporary_name)
    try:
        with zipfile.ZipFile(temporary_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for relative in existing_paths:
                source = data_root / relative
                if source.is_file():
                    archive.write(source, relative.as_posix())
                else:
                    for path in source.rglob("*"):
                        if path.is_file():
                            archive.write(path, path.relative_to(data_root).as_posix())
            archive.writestr("backup.json", json.dumps({
                "created_utc": timestamp,
                "application_version": version,
                "reason": reason,
                "format_version": "1",
            }, indent=2) + "\n")
        os.replace(temporary_path, archive_path)
        return archive_path
    finally:
        temporary_path.unlink(missing_ok=True)


def protect_data_before_startup(data_root: Path, resource_root: Path) -> Path | None:
    """Snapshot existing data before the first run or an application upgrade."""
    version = read_app_version(resource_root)
    state_path = data_root / STATE_FILE_NAME
    previous_version = None
    if state_path.exists():
        try:
            previous_version = json.loads(state_path.read_text(encoding="utf-8")).get("application_version")
        except (json.JSONDecodeError, OSError) as error:
            raise RuntimeError(f"Cannot read {state_path}; refusing to start without data protection") from error

    reason = "initial" if previous_version is None else "upgrade" if previous_version != version else None
    backup_path = snapshot_data(data_root, reason, version) if reason else None
    _atomic_write_json(state_path, {
        "application_version": version,
        "last_started_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(),
    })
    return backup_path