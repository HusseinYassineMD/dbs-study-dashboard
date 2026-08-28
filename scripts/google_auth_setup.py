#!/usr/bin/env python3
"""One-time Google sign-in for ajagadis@usc.edu to sync shared Google Sheets."""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CLIENT_SECRETS = ROOT / "credentials" / "google_oauth_client.json"
TOKEN_PATH = ROOT / "credentials" / "google_token.json"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


def main() -> None:
    if not CLIENT_SECRETS.exists():
        raise SystemExit(
            "Missing credentials/google_oauth_client.json\n"
            "Create an OAuth Desktop client in Google Cloud Console and download the JSON key."
        )

    from google_auth_oauthlib.flow import InstalledAppFlow

    flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_SECRETS), SCOPES)
    creds = flow.run_local_server(port=0)
    TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
    print(f"Saved Google token to {TOKEN_PATH}")
    print("You can now run: python scripts/sync_from_google_sheets.py")


if __name__ == "__main__":
    main()
