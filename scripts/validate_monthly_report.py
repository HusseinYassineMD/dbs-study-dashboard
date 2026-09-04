#!/usr/bin/env python3
"""Validate August 2026 monthly report metrics against the team PDF and Study Progress."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "dashboard" / "data.json"
REPORT_JS = ROOT / "dashboard" / "report.js"

KPI = {
    "enrolled": 200,
    "baseline_completed": 187,
    "baseline_dropped": 13,
    "year2_completed": 171,
    "year2_dropped": 15,
    "year3_completed": 120,
    "year3_dropped": 15,
    "year4_completed": 28,
    "year4_dropped": 0,
    "total_dropouts": 43,
    "active": 157,
    "recruitment_total": 1735,
}

# Every number on August DBS Monthly Progress Meeting.pdf (slides 2–10)
PDF_AUG_2026 = [
    ("Slide 2 · Y3 completed", "y3.completed", 120),
    ("Slide 2 · Y4 completed", "y4.completed", 28),
    ("Slide 2 · Y3 MRI components", "scheduling.y3Components.MRI Date", 119),
    ("Slide 2 · Y3 BD components", "scheduling.y3Components.BD Date", 122),
    ("Slide 2 · Y3 NP components", "scheduling.y3Components.NP Date", 121),
    ("Slide 2 · Y3 CV components", "scheduling.y3Components.CV Date", 122),
    ("Slide 2 · Future MRI", "scheduling.futureVisits.mri", 11),
    ("Slide 2 · Future BD", "scheduling.futureVisits.bd", 13),
    ("Slide 2 · Future NP", "scheduling.futureVisits.np", 14),
    ("Slide 2 · Future CV", "scheduling.futureVisits.cv", 13),
    ("Slide 2 · Dropouts total", "dropouts.total", 43),
    ("Slide 3 · Y3 MRI Aug", "y3Visits.mri", 5),
    ("Slide 3 · Y3 BD Aug", "y3Visits.bloodDraw", 5),
    ("Slide 3 · Y3 NP Aug", "y3Visits.np", 7),
    ("Slide 3 · Y3 CV Aug", "y3Visits.cv", 7),
    ("Slide 3 · Y4 BD Aug", "y4Visits.bloodDraw", 10),
    ("Slide 3 · Y4 NP Aug", "y4Visits.np", 10),
    ("Slide 3 · Y4 CV Aug", "y4Visits.cv", 9),
    ("Slide 4 · Y3 total visits Aug", "y3TotalMonthVisits", 24),
    ("Slide 4 · Y3 due n", "scheduling.y3Activity.dueCount", 6),
    ("Slide 4 · Y3 due complete", "scheduling.y3Activity.dueComplete", 3),
    ("Slide 4 · Y3 due scheduled", "scheduling.y3Activity.dueScheduled", 3),
    ("Slide 4 · Sep Y3 scheduled", "scheduling.y3Activity.nextVisitTotal", 35),
    ("Slide 6 · Y3 in progress", "scheduling.y3InProgress", 15),
    ("Slide 7 · Y3 to contact", "scheduling.y3ToContact", 22),
    ("Slide 8 · Y4 total visits Aug", "y4TotalMonthVisits", 29),
    ("Slide 8 · Y4 due n Aug", "scheduling.y4Activity.dueCount", 12),
    ("Slide 8 · Y4 due complete Aug", "scheduling.y4Activity.dueComplete", 11),
    ("Slide 8 · Sep Y4 scheduled", "scheduling.y4Activity.nextVisitTotal", 2),
    ("Slide 9 · Y4 to contact", "scheduling.y4ToContact", 26),
    ("Slide 10 · Enrolled", "enrolled", 200),
    ("Slide 10 · Baseline completed", "baseline.completed", 187),
    ("Slide 10 · Y2 completed", "y2.completed", 171),
]


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
        if part.endswith("]"):
            key, idx = part[:-1].split("[")
            cur = cur[key][int(idx)]
        else:
            cur = cur[part]
    return cur


def resolve_value(metrics: dict, path: str):
    if path in {"scheduling.y3InProgress", "scheduling.y3ToContact", "scheduling.y4ToContact"}:
        groups = get_path(metrics, path)
        return sum(g["count"] for g in groups)
    return get_path(metrics, path)


def main() -> int:
    if not DATA.exists() or not REPORT_JS.exists():
        print("Missing dashboard/data.json or dashboard/report.js")
        return 1

    data = json.loads(DATA.read_text())
    m = load_report_metrics(2026, 8)
    passed: list[str] = []
    failures: list[str] = []

    for label, path, expected in PDF_AUG_2026:
        try:
            actual = resolve_value(m, path)
        except (KeyError, TypeError):
            actual = None
        if actual is None:
            failures.append(f"{label}: MISSING (PDF {expected})")
        elif int(actual) == int(expected):
            passed.append(f"{label} = {expected}")
        else:
            failures.append(f"{label}: PDF {expected}, report {actual}")

    meeting = {str(r["year"]): r for r in data["study_progress_meeting"]}
    for label, key, kpi_key in [
        ("Baseline dropouts", "Baseline", "baseline_dropped"),
        ("Year 2 dropouts", "2", "year2_dropped"),
        ("Year 3 dropouts", "3", "year3_dropped"),
        ("Year 4 dropouts", "4", "year4_dropped"),
    ]:
        actual = meeting[key]["dropped_out_dq"]
        expected = KPI[kpi_key]
        if int(actual) == int(expected):
            passed.append(f"{label} = {expected}")
        else:
            failures.append(f"{label}: expected {expected}, got {actual}")

    print("=" * 72)
    print("AUGUST 2026 REPORT vs TEAM PDF — FULL VALIDATION")
    print("=" * 72)
    if m.get("usesMeetingSnapshot"):
        print(f"Snapshot: {m.get('reportSource')} (as of {m.get('reportAsOf')})")
    print(f"PASSED:   {len(passed)}")
    print(f"FAILURES: {len(failures)}")
    print()
    print(f"{'Metric':<42} {'PDF':>6} {'Report':>6} {'OK':>4}")
    print("-" * 62)
    for label, path, expected in PDF_AUG_2026:
        try:
            actual = resolve_value(m, path)
        except (KeyError, TypeError):
            actual = None
        ok = "✓" if actual is not None and int(actual) == int(expected) else "✗"
        print(f"{label:<42} {expected:>6} {str(actual):>6} {ok:>4}")

    if failures:
        print("\nFAILURES:")
        for f in failures:
            print(f"  ✗ {f}")
        print("\n❌ VALIDATION FAILED")
        return 1

    print("\n✅ ALL AUGUST NUMBERS MATCH THE TEAM PDF")
    return 0


if __name__ == "__main__":
    sys.exit(main())
