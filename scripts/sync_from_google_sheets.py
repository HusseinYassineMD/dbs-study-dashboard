#!/usr/bin/env python3
"""Download Google Sheets workbooks used by the DBS dashboard."""

from __future__ import annotations

import json
import os
import sys
from io import BytesIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "google_sheets.json"
EXAMPLE_CONFIG_PATH = ROOT / "config" / "google_sheets.example.json"
DEFAULT_CREDENTIALS_PATH = ROOT / "credentials" / "google_service_account.json"


def load_config() -> dict:
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    if EXAMPLE_CONFIG_PATH.exists():
        return json.loads(EXAMPLE_CONFIG_PATH.read_text(encoding="utf-8"))
    raise FileNotFoundError("Missing config/google_sheets.json")


def credentials_path() -> Path:
    env_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if env_path:
        return Path(env_path)
    return DEFAULT_CREDENTIALS_PATH


def extract_sheet_id(value: str) -> str:
    value = value.strip()
    if "/d/" in value:
        return value.split("/d/")[1].split("/")[0]
    return value


def export_spreadsheet(spreadsheet_id: str, output_path: Path, credentials_file: Path) -> None:
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload

    scopes = ["https://www.googleapis.com/auth/drive.readonly"]
    creds = Credentials.from_service_account_file(str(credentials_file), scopes=scopes)
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)

    request = drive.files().export_media(
        fileId=spreadsheet_id,
        mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    buffer = BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    output_path.write_bytes(buffer.getvalue())


def service_account_email(credentials_file: Path) -> str:
    data = json.loads(credentials_file.read_text(encoding="utf-8"))
    return data["client_email"]


def main() -> int:
    config = load_config()
    creds_file = credentials_path()

    if not creds_file.exists():
        print("Google credentials not found.", file=sys.stderr)
        print(f"Expected credentials at: {creds_file}", file=sys.stderr)
        print("", file=sys.stderr)
        print("Next steps:", file=sys.stderr)
        print("1. Create a Google Cloud service account with Drive API access.", file=sys.stderr)
        print("2. Save the JSON key to credentials/google_service_account.json", file=sys.stderr)
        print("3. Ask Jimena to share both Google Sheets with the service account email.", file=sys.stderr)
        print("4. Fill in config/google_sheets.json with the spreadsheet IDs.", file=sys.stderr)
        return 1

    tracker_id = extract_sheet_id(config["tracker_spreadsheet_id"])
    meeting_id = extract_sheet_id(config["meeting_spreadsheet_id"])
    tracker_output = ROOT / config.get("tracker_output", "DBS Tracker.xlsx")
    meeting_output = ROOT / config.get("meeting_output", "Monthly Meeting Updates.xlsx")

    if tracker_id.startswith("PASTE_") or meeting_id.startswith("PASTE_"):
        print("Google Sheet IDs are not configured yet.", file=sys.stderr)
        print(f"Edit {CONFIG_PATH} and add the spreadsheet IDs or full Google Sheets URLs.", file=sys.stderr)
        return 1

    print(f"Using service account: {service_account_email(creds_file)}")
    print(f"Downloading tracker sheet -> {tracker_output.name}")
    export_spreadsheet(tracker_id, tracker_output, creds_file)
    print(f"Downloading meeting updates sheet -> {meeting_output.name}")
    export_spreadsheet(meeting_id, meeting_output, creds_file)
    print("Google Sheets sync complete.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
