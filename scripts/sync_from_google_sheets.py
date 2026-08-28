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
USER_TOKEN_PATH = ROOT / "credentials" / "google_token.json"


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


def get_credentials():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials as UserCredentials
    from google.oauth2.service_account import Credentials as ServiceCredentials

    scopes = ["https://www.googleapis.com/auth/drive.readonly"]

    creds_file = credentials_path()
    if creds_file.exists():
        return ServiceCredentials.from_service_account_file(str(creds_file), scopes=scopes)

    if USER_TOKEN_PATH.exists():
        creds = UserCredentials.from_authorized_user_file(str(USER_TOKEN_PATH), scopes)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            USER_TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
        return creds

    token_json = os.environ.get("GOOGLE_OAUTH_TOKEN_JSON")
    if token_json:
        creds = UserCredentials.from_authorized_user_info(json.loads(token_json), scopes)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        return creds

    return None


def export_spreadsheet(spreadsheet_id: str, output_path: Path, credentials) -> None:
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload

    drive = build("drive", "v3", credentials=credentials, cache_discovery=False)

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


def main() -> int:
    config = load_config()
    credentials = get_credentials()

    if credentials is None:
        print("Google credentials not found.", file=sys.stderr)
        print("", file=sys.stderr)
        print("Jimena already shared the sheets with ajagadis@usc.edu.", file=sys.stderr)
        print("Run this one-time setup:", file=sys.stderr)
        print("  1. Create an OAuth Desktop client in Google Cloud Console", file=sys.stderr)
        print("  2. Save JSON to credentials/google_oauth_client.json", file=sys.stderr)
        print("  3. python scripts/google_auth_setup.py", file=sys.stderr)
        print("  4. python scripts/sync_from_google_sheets.py", file=sys.stderr)
        return 1

    tracker_id = extract_sheet_id(config["tracker_spreadsheet_id"])
    meeting_id = extract_sheet_id(config["meeting_spreadsheet_id"])
    tracker_output = ROOT / config.get("tracker_output", "DBS Tracker.xlsx")
    meeting_output = ROOT / config.get("meeting_output", "Monthly Meeting Updates.xlsx")

    if tracker_id.startswith("PASTE_") or meeting_id.startswith("PASTE_"):
        print("Google Sheet IDs are not configured yet.", file=sys.stderr)
        print(f"Edit {CONFIG_PATH} and add the spreadsheet IDs or full Google Sheets URLs.", file=sys.stderr)
        return 1

    owner = config.get("owner_email", "ajagadis@usc.edu")
    print(f"Syncing Google Sheets using access for {owner}")
    print(f"Downloading tracker sheet -> {tracker_output.name}")
    export_spreadsheet(tracker_id, tracker_output, credentials)
    print(f"Downloading meeting updates sheet -> {meeting_output.name}")
    export_spreadsheet(meeting_id, meeting_output, credentials)
    print("Google Sheets sync complete.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
