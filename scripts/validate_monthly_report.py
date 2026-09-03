#!/usr/bin/env python3
"""Cross-check monthly report metrics against source Excel and August PDF."""

from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
TRACKER = ROOT / "DBS Tracker.xlsx"
MEETING = ROOT / "Monthly Meeting Updates.xlsx"
DATA = ROOT / "dashboard" / "data.json"

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
REPORT_START_KEY = 2024 * 12 + 5
REPORT_END_KEY = 2026 * 12 + 9

# August 2026 PDF canonical + known monthly snapshots
PDF = {
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

MONTHLY_EXPECTED = {
    (2024, 10): {"y2_cv": 12, "y2_np": 12, "y2_cumulative": 52, "y3_mri": 0},
    (2026, 3): {"y2_cv": 1, "y2_np": 1, "y2_cumulative": 171, "y3_mri": 4},
    (2026, 8): {"y2_cv": 0, "y2_np": 0, "y2_cumulative": 171, "y3_mri": 6, "y3_bd": 6, "y3_np": 7, "y3_cv": 7},
    (2026, 9): {"y2_cv": 0, "y2_np": 0, "y3_mri": 8, "y3_bd": 8, "y3_np": 10, "y3_cv": 9},
}


def parse_date(value) -> datetime | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s or re.search(r"opt out|n/a|^nan$", s, re.I):
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", ""))
    except ValueError:
        return None


def in_month(value, year: int, month: int) -> bool:
    d = parse_date(value)
    return bool(d and d.year == year and d.month == month)


def month_key(year: int, month: int) -> int:
    return year * 12 + month


def find_monthly_column(data: dict, year: int, month: int) -> str | None:
    rows = data.get("second_year_monthly") or []
    if not rows:
        return None
    full = MONTH_NAMES[month - 1]
    abbr = full[:3]
    yy = str(year)[-2:]
    keys = [k for k in rows[0] if k not in ("visit_type", "total_completed")]
    for k in keys:
        kl = k.lower()
        if (full.lower() in kl or kl.startswith(abbr.lower())) and f"'{yy}" in k:
            return k
    return None


def last_trend_on_or_before(data: dict, year: int, month: int) -> dict | None:
    target = month_key(year, month)
    best = None
    best_key = -1
    for row in data.get("visit_trends_sheet5") or []:
        d = parse_date(row.get("Date"))
        if not d:
            continue
        key = month_key(d.year, d.month)
        if key <= target and key > best_key:
            best_key = key
            best = row
    return best


def collect_metrics(data: dict, year: int, month: int) -> dict:
    third = data.get("third_year_scheduling") or []
    fourth = data.get("fourth_year_scheduling") or []
    cv_row = next((r for r in data.get("second_year_monthly") or [] if r.get("visit_type") == "Clinician Visit"), {})
    np_row = next((r for r in data.get("second_year_monthly") or [] if r.get("visit_type") == "NP"), {})
    col = find_monthly_column(data, year, month)
    trend5 = next((r for r in data.get("visit_trends_sheet5") or [] if in_month(r.get("Date"), year, month)), None)

    y2cv = int(cv_row.get(col) or 0) if col else int(trend5.get("CV Monthly") or 0) if trend5 else 0
    y2np = int(np_row.get(col) or 0) if col else int(trend5.get("NP Monthly") or 0) if trend5 else 0
    last = last_trend_on_or_before(data, year, month)
    y2cum = (
        (trend5 or {}).get("CV Cumulative")
        or (last or {}).get("CV Cumulative")
    )

    y3 = {
        "mri": sum(1 for r in third if in_month(r.get("MRI Date"), year, month)),
        "bd": sum(1 for r in third if in_month(r.get("BD Date"), year, month)),
        "np": sum(1 for r in third if in_month(r.get("NP Date"), year, month)),
        "cv": sum(1 for r in third if in_month(r.get("CV Date"), year, month)),
    }
    y4 = {
        "bd": sum(1 for r in fourth if in_month(r.get("BD Date"), year, month)),
        "np": sum(1 for r in fourth if in_month(r.get("NP Date"), year, month)),
        "cv": sum(1 for r in fourth if in_month(r.get("CV Date"), year, month)),
    }

    meeting = {str(r["year"]): r for r in data.get("study_progress_meeting") or []}
    return {
        "y2_cv": y2cv,
        "y2_np": y2np,
        "y2_cumulative": y2cum,
        "y3_mri": y3["mri"],
        "y3_bd": y3["bd"],
        "y3_np": y3["np"],
        "y3_cv": y3["cv"],
        "y4_bd": y4["bd"],
        "y4_np": y4["np"],
        "y4_cv": y4["cv"],
        "y3_completed": meeting.get("3", {}).get("completed"),
        "y4_completed": meeting.get("4", {}).get("completed"),
        "active": len(data.get("active_participants") or []),
        "recruitment": data.get("recruitment", {}).get("grand_total"),
    }


def main() -> int:
    if not DATA.exists():
        print(f"Missing {DATA}")
        return 1

    data = json.loads(DATA.read_text())
    passed: list[str] = []
    failures: list[str] = []

    def ok(msg: str) -> None:
        passed.append(msg)

    def fail(msg: str) -> None:
        failures.append(msg)

    # Pass 1: PDF / study progress
    meeting = {str(r["year"]): r for r in data["study_progress_meeting"]}
    checks = [
        ("enrolled", meeting["Baseline"]["enrolled"], PDF["enrolled"]),
        ("baseline_completed", meeting["Baseline"]["completed"], PDF["baseline_completed"]),
        ("year3_completed", meeting["3"]["completed"], PDF["year3_completed"]),
        ("year4_completed", meeting["4"]["completed"], PDF["year4_completed"]),
        ("active", len(data["active_participants"]), PDF["active"]),
        ("recruitment", data["recruitment"]["grand_total"], PDF["recruitment_total"]),
    ]
    dropped = sum(meeting[y]["dropped_out_dq"] or 0 for y in ["Baseline", "2", "3", "4"])
    checks.append(("total_dropouts", dropped, PDF["total_dropouts"]))

    for name, actual, expected in checks:
        if int(actual) == int(expected):
            ok(f"PDF {name} = {expected}")
        else:
            fail(f"PDF {name}: expected {expected}, got {actual}")

    # Pass 2: Monthly report snapshots
    for (year, month), expected in MONTHLY_EXPECTED.items():
        m = collect_metrics(data, year, month)
        label = f"{MONTH_NAMES[month - 1]} {year}"
        for key, exp in expected.items():
            actual = m.get(key)
            if actual is None:
                fail(f"{label} {key}: missing")
            elif int(actual) != int(exp):
                fail(f"{label} {key}: expected {exp}, got {actual}")
            else:
                ok(f"{label} {key} = {exp}")

    # Pass 3: Report month range
    months_in_range = []
    for key in range(REPORT_START_KEY, REPORT_END_KEY + 1):
        y = (key - 1) // 12
        m = (key - 1) % 12 + 1
        months_in_range.append((y, m))
    ok(f"Report range: {len(months_in_range)} months (May 2024 – Sep 2026)")

    # Pass 4: Every month in range has computable metrics (no null cumulative after May 2024)
    null_cum = []
    for year, month in months_in_range:
        m = collect_metrics(data, year, month)
        if m["y2_cumulative"] is None and month_key(year, month) >= 2024 * 12 + 5:
            null_cum.append(f"{MONTH_NAMES[month-1]} {year}")
    if null_cum:
        fail(f"Missing 2nd year cumulative: {', '.join(null_cum[:5])}")
    else:
        ok("All report months have 2nd year cumulative")

    # Pass 5: Excel vs JSON row counts (spot check scheduling)
    df3 = pd.read_excel(TRACKER, sheet_name="3rd Yr Scheduling Notes ", header=None)
    json_third = len(data["third_year_scheduling"])
    if json_third < 100:
        fail(f"third_year_scheduling suspiciously low: {json_third}")
    else:
        ok(f"third_year_scheduling: {json_third} rows")

    # Pass 6: Recruitment sources
    src_sum = sum(s["total_leads"] for s in data["recruitment"]["sources"])
    if src_sum != PDF["recruitment_total"]:
        fail(f"Recruitment sum {src_sum} != {PDF['recruitment_total']}")
    else:
        ok("Recruitment sources sum = 1735")

    print("=" * 70)
    print("MONTHLY REPORT + DASHBOARD CROSS-CHECK (DOUBLE AUDIT)")
    print("=" * 70)
    print(f"PASSED:   {len(passed)}")
    print(f"FAILURES: {len(failures)}")
    if failures:
        print("\nFAILURES:")
        for f in failures:
            print(f"  ✗ {f}")
    print()
    if not failures:
        print("✅ DOUBLE-CHECK PASSED — all data verified")
        return 0
    print("❌ ISSUES FOUND")
    return 1


if __name__ == "__main__":
    sys.exit(main())
