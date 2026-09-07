#!/usr/bin/env bash
# Run all report/dashboard validation and unit tests.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

echo "== Dashboard data validation =="
python3 scripts/validate_dashboard_data.py

echo
echo "== August 2026 PDF validation =="
python3 scripts/validate_monthly_report.py

echo
echo "== Report metric unit tests (Node) =="
node --test scripts/test_report_metrics.test.js

echo
echo "== Report metric unit tests (Python) =="
python3 scripts/test_report_validation.py

echo
echo "✅ All validations and unit tests passed"
