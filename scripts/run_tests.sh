#!/usr/bin/env bash
# Run all report/dashboard validation and unit tests.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

PYTHON="python3"
if [[ -x "$ROOT/.venv/bin/python" ]]; then
  PYTHON="$ROOT/.venv/bin/python"
fi

echo "== Dashboard data validation =="
"$PYTHON" scripts/validate_dashboard_data.py

echo
echo "== Multi-month PDF validation (Jimena decks) =="
"$PYTHON" scripts/validate_all_monthly_reports.py

echo
echo "== August 2026 PDF validation (legacy check) =="
"$PYTHON" scripts/validate_monthly_report.py

echo
echo "== Report metric unit tests (Node) =="
node --test scripts/test_report_metrics.test.js

echo
echo "== Report metric unit tests (Python) =="
"$PYTHON" scripts/test_report_validation.py

echo
echo "✅ All validations and unit tests passed"
