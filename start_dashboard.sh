#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

if [[ ! -d .venv ]]; then
  python3 -m venv .venv
  source .venv/bin/activate
  pip install -r requirements.txt
else
  source .venv/bin/activate
fi

python scripts/build_dashboard_data.py
echo ""
echo "Dashboard ready at http://localhost:8080"
python -m http.server 8080 --directory dashboard
