#!/usr/bin/env bash
# Rebuild data and publish the live dashboard to GitHub Pages.
set -euo pipefail
cd "$(dirname "$0")"

source .venv/bin/activate

if [[ -f credentials/google_service_account.json && -f config/google_sheets.json ]] || [[ -f credentials/google_token.json ]]; then
  python scripts/sync_from_google_sheets.py
fi

python scripts/build_dashboard_data.py

git add dashboard/data.json "DBS Tracker.xlsx" "Monthly Meeting Updates.xlsx" 2>/dev/null || git add dashboard/data.json
git diff --cached --quiet || git commit -m "Update dashboard data $(date +%Y-%m-%d)"
git push origin main

echo ""
echo "Live site: https://husseinyassinemd.github.io/dbs-study-dashboard/"
echo "GitHub Pages deploys automatically after push (~1 minute)."
