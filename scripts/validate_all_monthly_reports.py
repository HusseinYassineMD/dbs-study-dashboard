#!/usr/bin/env python3
"""Validate generated report metrics against team meeting PDFs for all available months."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "dashboard" / "data.json"
REPORT_JS = ROOT / "dashboard" / "report.js"
FIXTURES = ROOT / "scripts" / "fixtures" / "report_expectations.json"

GROUP_SUM_PATHS = {
    "scheduling.y3InProgress",
    "scheduling.y3ToContact",
    "scheduling.y4ToContact",
}


def load_report_metrics(year: int, month: int) -> dict:
    node = """
const fs = require('fs');
global.DATA = JSON.parse(fs.readFileSync('dashboard/data.json','utf8'));
eval(fs.readFileSync('dashboard/report.js','utf8').replace(/^async function generateMonthlyReportPPT[\\s\\S]*/,''));
const m = collectReportMetrics(%d, %d);
console.log(JSON.stringify(m));
""" % (
        year,
        month,
    )
    out = subprocess.check_output(["node", "-e", node], cwd=ROOT, text=True)
    return json.loads(out)


def get_path(obj: dict, path: str):
    cur = obj
    for part in path.split("."):
        cur = cur[part]
    return cur


def resolve_value(metrics: dict, path: str):
    if path in GROUP_SUM_PATHS:
        groups = get_path(metrics, path)
        return sum(g["count"] for g in groups)
    return get_path(metrics, path)


def within_tolerance(actual: int, expected: int, tolerance: int) -> bool:
    return abs(int(actual) - int(expected)) <= tolerance


def validate_month(key: str, spec: dict) -> tuple[list[str], list[str], list[str]]:
    passed: list[str] = []
    failures: list[str] = []
    warnings: list[str] = []
    metrics = load_report_metrics(spec["year"], spec["month"])
    tolerance_map = spec.get("tolerance") or {}

    for label, path, expected in spec.get("metrics", []):
        try:
            actual = resolve_value(metrics, path)
        except (KeyError, TypeError):
            actual = None
        tol = int(tolerance_map.get(path, 0))
        if actual is None:
            failures.append(f"{label}: missing (PDF {expected})")
        elif within_tolerance(actual, expected, tol):
            note = f" (±{tol})" if tol and int(actual) != int(expected) else ""
            passed.append(f"{label} = {expected}{note}")
            if tol and int(actual) != int(expected):
                warnings.append(f"{label}: PDF {expected}, report {actual} (within ±{tol} meeting-date tolerance)")
        else:
            failures.append(f"{label}: PDF {expected}, report {actual}")

    for field, expected in (spec.get("narrative") or {}).items():
        actual = metrics.get("narrative", {}).get(field)
        if actual == expected:
            passed.append(f"narrative.{field} = {expected}")
        else:
            failures.append(f"narrative.{field}: PDF {expected!r}, report {actual!r}")

    if not metrics.get("reportSnapshot"):
        warnings.append("No study_progress_snapshot applied — using current Study Progress sheet")

    return passed, failures, warnings


def main() -> int:
    if not DATA.exists() or not REPORT_JS.exists():
        print("Missing dashboard/data.json or dashboard/report.js")
        return 1
    if not FIXTURES.exists():
        print(f"Missing {FIXTURES}")
        return 1

    fixtures = json.loads(FIXTURES.read_text())
    months: dict[str, dict] = fixtures.get("months") or {}
    if not months:
        print("No month fixtures found.")
        return 1

    total_pass = 0
    total_fail = 0
    total_warn = 0
    month_failures: list[str] = []

    print("=" * 72)
    print("MULTI-MONTH REPORT VALIDATION — TEAM MEETING PDFs")
    print("=" * 72)
    print(fixtures.get("_notes", ""))
    print()

    for key in sorted(months.keys()):
        spec = months[key]
        label = spec.get("label", key)
        pdf = spec.get("pdf", "—")
        passed, failures, warnings = validate_month(key, spec)
        total_pass += len(passed)
        total_fail += len(failures)
        total_warn += len(warnings)

        status = "✅ PASS" if not failures else "❌ FAIL"
        print(f"{status}  {label}  ({pdf})")
        print(f"       passed {len(passed)}, failed {len(failures)}, warnings {len(warnings)}")
        if failures:
            for item in failures:
                print(f"         ✗ {item}")
                month_failures.append(f"{label}: {item}")
        if warnings:
            for item in warnings[:5]:
                print(f"         ⚠ {item}")
            if len(warnings) > 5:
                print(f"         ⚠ … and {len(warnings) - 5} more warnings")
        print()

    print("=" * 72)
    print(f"TOTAL: {total_pass} passed, {total_fail} failed, {total_warn} warnings across {len(months)} months")
    print("=" * 72)

    if total_fail:
        print("\n❌ VALIDATION FAILED — see mismatches above")
        return 1

    print("\n✅ ALL PDF VALIDATIONS PASSED (including allowed meeting-date tolerances)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
