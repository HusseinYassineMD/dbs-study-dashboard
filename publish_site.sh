#!/usr/bin/env bash
# Rebuild data and publish the live dashboard to GitHub Pages.
set -euo pipefail
cd "$(dirname "$0")"

source .venv/bin/activate
python scripts/build_dashboard_data.py

SITE_DIR=$(mktemp -d)
trap 'rm -rf "$SITE_DIR"' EXIT

git clone --depth 1 https://github.com/HusseinYassineMD/dbs-study-dashboard-site.git "$SITE_DIR"
cp dashboard/index.html dashboard/app.js dashboard/styles.css dashboard/data.json "$SITE_DIR/"
cd "$SITE_DIR"
git add -A
git commit -m "Update dashboard data $(date +%Y-%m-%d)" || { echo "No changes to publish."; exit 0; }
git push origin main

gh workflow run deploy-pages.yml --repo HusseinYassineMD/dbs-study-dashboard-site >/dev/null 2>&1 || true

echo ""
echo "Live site: https://husseinyassinemd.github.io/dbs-study-dashboard-site/"
echo "Deployment triggered — live in ~1 minute."
