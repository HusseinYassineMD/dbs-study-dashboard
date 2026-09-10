#!/usr/bin/env python3
"""
Validate report CALCULATIONS against team meeting PDFs (validation reference only).

Reports are always generated from live Excel/sheet data. PDFs prove the logic is
correct — especially month-level visit and activity counts. Cumulative Study
Progress (Y3/Y4 completed, dropouts) is only compared to PDF for the latest
deck (August 2026), because those numbers grow each month.

Meetings were typically held on the last Friday; small variance (±1–2) on visit
and due-date counts is expected per Jimena.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "dashboard" / "data.json"
REPORT_JS = ROOT / "dashboard" / "report.js"
FIXTURES = ROOT / "scripts" / "fixtures" / "report_expectations.json"

# Latest PDF with full Study Progress — cumulative metrics validated here only.
LATEST_CUMULATIVE_MONTH = "2026-08"

GROUP_SUM_PATHS = {
    "scheduling.y3InProgress",
    "scheduling.y3ToContact",
    "scheduling.y4ToContact",
}

# Paths that reflect study-wide totals as-of meeting time (not month-local).
CUMULATIVE_PATHS = {
    "y3.completed",
    "y4.completed",
    "y2.completed",
    "enrolled",
    "baseline.completed",
    "dropouts.total",
    "dropouts.baseline",
    "dropouts.year2",
    "dropouts.year3",
    "dropouts.year4",
    "scheduling.y3Components.MRI Date",
    "scheduling.y3Components.BD Date",
    "scheduling.y3Components.NP Date",
    "scheduling.y3Components.CV Date",
    "scheduling.y4Components.BD Date",
    "scheduling.y4Components.NP Date",
    "scheduling.y4Components.CV Date",
    "scheduling.futureVisits.mri",
    "scheduling.futureVisits.bd",
    "scheduling.futureVisits.np",
    "scheduling.futureVisits.cv",
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


def is_cumulative_metric(path: str) -> bool:
    return path in CUMULATIVE_PATHS


def validate_month(key: str, spec: dict) -> tuple[list[str], list[str], list[str], list[str]]:
    passed: list[str] = []
    failures: list[str] = []
    warnings: list[str] = []
    skipped: list[str] = []
    metrics = load_report_metrics(spec["year"], spec["month"])
    tolerance_map = spec.get("tolerance") or {}
    validate_cumulative = key == LATEST_CUMULATIVE_MONTH

    for label, path, expected in spec.get("metrics", []):
        if is_cumulative_metric(path) and not validate_cumulative:
            skipped.append(f"{label} (cumulative — PDF was point-in-time at meeting)")
            continue

        try:
            actual = resolve_value(metrics, path)
        except (KeyError, TypeError):
            actual = None
        tol = int(tolerance_map.get(path, 0))
        if actual is None:
            failures.append(f"{label}: missing (PDF {expected})")
        elif within_tolerance(actual, expected, tol):
            passed.append(f"{label} = {expected}")
            if tol and int(actual) != int(expected):
                warnings.append(
                    f"{label}: PDF {expected}, report {actual} (within ±{tol} last-Friday tolerance)"
                )
        else:
            failures.append(f"{label}: PDF {expected}, report {actual}")

    if validate_cumulative:
        for field, expected in (spec.get("narrative") or {}).items():
            actual = metrics.get("narrative", {}).get(field)
            if actual == expected:
                passed.append(f"narrative.{field} = {expected}")
            else:
                failures.append(f"narrative.{field}: PDF {expected!r}, report {actual!r}")

    return passed, failures, warnings, skipped


def validate_september_smoke() -> tuple[list[str], list[str]]:
    """September 2026 has no PDF yet — ensure live calculations produce sensible output."""
    passed: list[str] = []
    failures: list[str] = []
    try:
        m = load_report_metrics(2026, 9)
    except Exception as exc:
        failures.append(f"September 2026 metrics failed to compute: {exc}")
        return passed, failures

    checks = [
        ("September label", m.get("label") == "September 2026"),
        ("Y3 visit total > 0", (m.get("y3TotalMonthVisits") or 0) > 0),
        ("Y4 visit total > 0", (m.get("y4TotalMonthVisits") or 0) > 0),
        ("Scheduling block present", bool(m.get("scheduling"))),
        ("Narrative block present", bool(m.get("narrative"))),
        ("Study progress populated", m.get("y3", {}).get("completed") is not None),
    ]
    for label, ok in checks:
        if ok:
            passed.append(label)
        else:
            failures.append(label)
    return passed, failures


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
    total_skip = 0

    print("=" * 72)
    print("REPORT CALCULATION VALIDATION vs TEAM PDFs (reference only)")
    print("=" * 72)
    print(
        "Reports use live sheet data. PDFs validate the calculation logic.\n"
        "Month-level visits/activity: checked for every PDF month.\n"
        f"Cumulative Study Progress: checked against PDF for {LATEST_CUMULATIVE_MONTH} only.\n"
        "Small variance allowed — meetings were on the last Friday of the month."
    )
    print()

    for key in sorted(months.keys()):
        spec = months[key]
        label = spec.get("label", key)
        pdf = spec.get("pdf", "—")
        passed, failures, warnings, skipped = validate_month(key, spec)
        total_pass += len(passed)
        total_fail += len(failures)
        total_warn += len(warnings)
        total_skip += len(skipped)

        status = "✅ PASS" if not failures else "❌ FAIL"
        print(f"{status}  {label}  ({pdf})")
        print(
            f"       passed {len(passed)}, failed {len(failures)}, "
            f"warnings {len(warnings)}, skipped {len(skipped)} cumulative"
        )
        if failures:
            for item in failures:
                print(f"         ✗ {item}")
        if warnings:
            for item in warnings[:4]:
                print(f"         ⚠ {item}")
            if len(warnings) > 4:
                print(f"         ⚠ … and {len(warnings) - 4} more")
        print()

    print("=" * 72)
    print("SEPTEMBER 2026 — LIVE CALCULATION SMOKE TEST (no PDF yet)")
    print("=" * 72)
    sep_pass, sep_fail = validate_september_smoke()
    total_pass += len(sep_pass)
    total_fail += len(sep_fail)
    for item in sep_pass:
        print(f"  ✓ {item}")
    for item in sep_fail:
        print(f"  ✗ {item}")
    print()

    print("=" * 72)
    print(
        f"TOTAL: {total_pass} passed, {total_fail} failed, "
        f"{total_warn} warnings, {total_skip} cumulative skipped"
    )
    print("=" * 72)

    if total_fail:
        print("\n❌ VALIDATION FAILED")
        return 1

    print("\n✅ CALCULATION LOGIC VALIDATED — safe to generate September+ from live sheets")
    return 0


if __name__ == "__main__":
    sys.exit(main())
