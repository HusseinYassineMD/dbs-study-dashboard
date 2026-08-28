#!/usr/bin/env bash
# One-time setup for live Google Sheets → dashboard daily sync.
set -euo pipefail
cd "$(dirname "$0")"

echo "=== DBS Live Google Sheets Setup ==="
echo ""
echo "This connects Jimena's Google Sheets to the live dashboard"
echo "and enables daily automatic refresh."
echo ""

mkdir -p credentials

if [[ ! -f credentials/google_oauth_client.json ]]; then
  echo "STEP 1: Create Google OAuth credentials"
  echo "  1. Open https://console.cloud.google.com/apis/credentials"
  echo "  2. Create project (or select existing)"
  echo "  3. Enable Google Drive API"
  echo "  4. Create OAuth client -> Desktop app"
  echo "  5. Download JSON -> save as:"
  echo "     credentials/google_oauth_client.json"
  echo ""
  echo "Then run this script again."
  exit 1
fi

source .venv/bin/activate
pip install -q -r requirements.txt

if [[ ! -f credentials/google_token.json ]]; then
  echo "STEP 2: Sign in with ajagadis@usc.edu (browser will open)..."
  python scripts/google_auth_setup.py
fi

echo ""
echo "STEP 3: Syncing from Google Sheets..."
python scripts/sync_from_google_sheets.py

echo ""
echo "STEP 4: Building and publishing live dashboard..."
python scripts/build_dashboard_data.py
git add dashboard/data.json "DBS Tracker.xlsx" "Monthly Meeting Updates.xlsx" 2>/dev/null || git add dashboard/data.json
git diff --cached --quiet || git commit -m "Sync dashboard from Google Sheets"
git push origin main

echo ""
echo "STEP 5: Add GitHub secrets for DAILY auto-refresh (copy/paste these commands):"
echo ""
echo "gh secret set GOOGLE_OAUTH_TOKEN_JSON --repo HusseinYassineMD/dbs-study-dashboard < credentials/google_token.json"
echo "gh secret set GOOGLE_TRACKER_SHEET_ID --repo HusseinYassineMD/dbs-study-dashboard --body \"1vx20S9kklt4WSOO3qc7fCZW4fBVvKaRUYMJBcQ5Wonc\""
echo "gh secret set GOOGLE_MEETING_SHEET_ID --repo HusseinYassineMD/dbs-study-dashboard --body \"1zAHmTDBxkPDdFXChnIiHR11jTyx4h-Siu-IcJFR3Zf0\""
echo ""
echo "Done! Live site: https://husseinyassinemd.github.io/dbs-study-dashboard/"
echo "Daily sync runs automatically at 7:00 AM Pacific once secrets are added."
