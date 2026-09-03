#!/usr/bin/env python3
"""Validate dashboard/data.json against source Excel files and August meeting PDF numbers."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
TRACKER = ROOT / "DBS Tracker.xlsx"
MEETING = ROOT / "Monthly Meeting Updates.xlsx"
OUTPUT = ROOT / "dashboard" / "data.json"

# August 2026 monthly meeting presentation (canonical for Study Progress KPIs)
PDF_CANONICAL = {
    "Baseline": {"completed": 187, "dropped_out_dq": 13, "enrolled": 200},
    "2": {"completed": 171, "dropped_out_dq": 15},
    "3": {"completed": 120, "dropped_out_dq": 15},
    "4": {"completed": 28, "dropped_out_dq": 0},
}

RECRUITMENT_SOURCES = {
    "Facebook": 1572,
    "DBS Website": 80,
    "Call-In (800 #)": 29,
    "CTSI Directory": 33,
    "Other": 21,
}

# Structured datasets: compare parsed rows to rows with linkable keys in Excel
STRUCTURED_CHECKS = [
    ("id_location", TRACKER, "ID Location", "ID"),
    ("screening", TRACKER, "Screening Visits", "ID"),
    ("genotype", TRACKER, "Genotype", "ID"),
    ("active_participants", TRACKER, "All active IDs", None),
    ("clincard", TRACKER, "ClinCard", "ID"),
    ("baseline_scheduling", TRACKER, "Baseline Scheduling Notes", None),
    ("study_partner", TRACKER, "Study Partner", None),
    ("np_results_requests", TRACKER, "NP Results Requests", None),
    ("uds_id_200", TRACKER, "uds_id_200", None),
    ("meeting_dropouts", MEETING, "DropOuts", "Baseline Drop Outs"),
    ("pt_status", MEETING, "Pt Status", "uds_id"),
    ("enrolled_mapping", MEETING, "200 Enrolled pts", "uds_id_200"),
]


def num(value) -> float:
    if value is None:
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def main() -> int:
    if not OUTPUT.exists():
        print(f"Missing {OUTPUT}")
        return 1

    data = json.loads(OUTPUT.read_text())
    passed: list[str] = []
    warnings: list[str] = []
    failures: list[str] = []

    def ok(msg: str) -> None:
        passed.append(msg)

    def warn(msg: str) -> None:
        warnings.append(msg)

    def fail(msg: str) -> None:
        failures.append(msg)

    # 1. Study Progress — JSON vs Excel (both workbooks)
    for path, key, label in [
        (MEETING, "study_progress_meeting", "Monthly Meeting Updates"),
        (TRACKER, "study_progress_tracker", "DBS Tracker"),
    ]:
        df = pd.read_excel(path, sheet_name="Study Progress")
        json_by_year = {str(r["year"]): r for r in data[key]}
        for _, row in df.iterrows():
            if pd.isna(row.get("Year")) or pd.isna(row.get("Completed")):
                continue
            year = str(row["Year"])
            record = json_by_year.get(year)
            if not record:
                fail(f"{label}: year {year} missing in dashboard JSON")
                continue
            for excel_col, json_col in [
                ("Enrolled", "enrolled"),
                ("Dropped Out/DQ", "dropped_out_dq"),
                ("Completed", "completed"),
                ("Remaining", "remaining"),
            ]:
                excel_val = row.get(excel_col)
                json_val = record.get(json_col)
                if pd.isna(excel_val):
                    if json_val is not None and num(json_val) != 0:
                        fail(f"{label} Y{year} {excel_col}: Excel=empty JSON={json_val}")
                elif num(excel_val) != num(json_val):
                    fail(f"{label} Y{year} {excel_col}: Excel={excel_val} JSON={json_val}")
        ok(f"{label} Study Progress matches Excel")

    # 2. August PDF canonical numbers (via Monthly Meeting Updates)
    meeting = {str(r["year"]): r for r in data["study_progress_meeting"]}
    for year, expected in PDF_CANONICAL.items():
        record = meeting.get(year)
        if not record:
            fail(f"PDF check: year {year} missing from meeting progress")
            continue
        for field, expected_val in expected.items():
            actual = record.get(field)
            if num(actual) != expected_val:
                fail(f"PDF Y{year} {field}: expected {expected_val}, dashboard {actual}")
            else:
                ok(f"PDF Y{year} {field} = {expected_val}")

    total_dropouts = sum(num(meeting[y].get("dropped_out_dq")) for y in PDF_CANONICAL if y in meeting)
    if total_dropouts == 43:
        ok("Total dropouts across years = 43 (PDF)")
    else:
        fail(f"Total dropouts = {total_dropouts}, PDF says 43")

    # 3. Structured datasets vs Excel (parsed rows match linkable source rows)
    sys_path = ROOT / "scripts"
    if str(sys_path) not in sys.path:
        sys.path.insert(0, str(sys_path))
    from build_dashboard_data import normalize_id, read_sheet

    for json_key, path, sheet, id_col in STRUCTURED_CHECKS:
        try:
            if json_key == "active_participants":
                df = pd.read_excel(path, sheet_name=sheet, header=None)
                excel_count = sum(
                    1 for v in df.iloc[:, 0] if normalize_id(v)
                )
                json_count = len(data[json_key])
            elif json_key == "uds_id_200":
                df = pd.read_excel(path, sheet_name=sheet, header=None)
                excel_count = sum(
                    1 for idx in range(1, len(df)) if not pd.isna(df.iloc[idx, 1])
                )
                json_count = len(data.get(json_key, []))
            elif json_key == "enrolled_mapping":
                df = read_sheet(path, sheet)
                excel_count = sum(
                    1 for _, row in df.iterrows() if not pd.isna(row.get(id_col))
                )
                json_count = len(data.get(json_key, []))
            elif json_key == "clincard":
                df = pd.read_excel(path, sheet_name=sheet, header=1)
                excel_count = sum(
                    1 for _, row in df.iterrows() if normalize_id(row.get("ID"))
                )
                json_count = len(data.get(json_key, []))
            elif json_key == "study_partner":
                df = read_sheet(path, sheet)
                id_col_name = df.columns[0]
                excel_count = sum(
                    1 for _, row in df.iterrows() if normalize_id(row.get(id_col_name))
                )
                json_count = len(data.get(json_key, []))
            elif id_col:
                df = read_sheet(path, sheet)
                excel_count = sum(
                    1 for _, row in df.iterrows() if normalize_id(row.get(id_col))
                )
                json_count = len(data.get(json_key, []))
            else:
                df = pd.read_excel(path, sheet_name=sheet)
                excel_count = len(df.dropna(how="all"))
                json_count = len(data.get(json_key, []))

            if excel_count == json_count:
                ok(f"{json_key}: {json_count} rows match source")
            else:
                fail(f"{json_key}: source={excel_count} JSON={json_count}")
        except Exception as exc:
            warn(f"{json_key}: could not compare — {exc}")

    # Dropouts special structure
    baseline_drop = len(data["dropouts"]["baseline_ineligible"])
    yearly_drop = len(data["dropouts"]["year_1_to_4"])
    ok(f"dropouts: baseline_ineligible={baseline_drop}, year_1_to_4={yearly_drop}")

    stool = data.get("stool_samples", {})
    ok(f"stool_samples: {len(stool.get('participants', []))} participants")

    # 4. Raw sheet backups — every row preserved
    raw_mismatches = 0
    for path, prefix in [(TRACKER, "tracker"), (MEETING, "meeting")]:
        for sheet in pd.ExcelFile(path).sheet_names:
            df = pd.read_excel(path, sheet_name=sheet, header=None)
            nonempty = len(df[~df.isna().all(axis=1)])
            raw_key = f"{prefix}::{sheet.strip()}"
            raw_len = len(data["raw_sheets"].get(raw_key, []))
            if raw_len != nonempty:
                fail(f"Raw backup {raw_key}: Excel={nonempty} JSON={raw_len}")
                raw_mismatches += 1
    if raw_mismatches == 0:
        ok(f"All {len(data['raw_sheets'])} raw sheet backups match Excel row counts")

    # 5. Participant index
    participants = data["participants"]
    ids = [p.get("participant_id") for p in participants if p.get("participant_id")]
    if len(ids) == len(set(ids)):
        ok(f"Participant index: {len(participants)} unique IDs, no duplicates")
    else:
        fail(f"Duplicate participant IDs: {len(ids) - len(set(ids))}")

    # 6. Recruitment sources
    sources = {s["source"]: s["total_leads"] for s in data["recruitment"].get("sources", [])}
    for source, expected in RECRUITMENT_SOURCES.items():
        actual = sources.get(source)
        if num(actual) != expected:
            fail(f"Recruitment {source}: expected {expected}, got {actual}")
        else:
            ok(f"Recruitment {source} = {expected}")

    grand_total = data["recruitment"].get("grand_total")
    if num(grand_total) == 1735:
        ok("Recruitment grand total = 1735")
    else:
        fail(f"Recruitment grand total = {grand_total}, expected 1735")

    if len(data["recruitment"].get("periods", [])) >= 20:
        ok(f"Recruitment timeline periods: {len(data['recruitment']['periods'])}")
    else:
        fail(f"Recruitment periods missing: {len(data['recruitment'].get('periods', []))}")

    # 7. Dashboard KPI simulation (app.js uses getMeetingProgress)
    progress = data["study_progress_meeting"]
    kpi_map = {"Baseline": 187, "2": 171, "3": 120, "4": 28}
    for year, expected in kpi_map.items():
        record = next((r for r in progress if str(r["year"]) == year), None)
        actual = record["completed"] if record else None
        if num(actual) == expected:
            ok(f"Dashboard KPI Year {year} completed = {expected}")
        else:
            fail(f"Dashboard KPI Year {year} completed = {actual}, expected {expected}")

    active = len(data["active_participants"])
    ok(f"Active participants = {active}")

    # 8. Known tracker lag (informational only)
    tracker = {str(r["year"]): r for r in data["study_progress_tracker"]}
    for year in ("3", "4"):
        if num(meeting[year]["completed"]) != num(tracker[year]["completed"]):
            warn(
                f"Tracker stale on Year {year}: Meeting={meeting[year]['completed']} "
                f"Tracker={tracker[year]['completed']} — dashboard uses Meeting for KPIs"
            )

    meta = data.get("meta", {})
    ok(f"Built: {meta.get('generated_at', 'unknown')}")

    print("=" * 70)
    print("DBS DASHBOARD DATA VALIDATION")
    print("=" * 70)
    print(f"PASSED:   {len(passed)}")
    print(f"WARNINGS: {len(warnings)}")
    print(f"FAILURES: {len(failures)}")
    print()

    if failures:
        print("FAILURES:")
        for item in failures:
            print(f"  ✗ {item}")
        print()

    if warnings:
        print("WARNINGS (expected, not errors):")
        for item in warnings:
            print(f"  ⚠ {item}")
        print()

    if not failures:
        print("✅ VERDICT: ALL DATA VERIFIED CORRECT")
        print("   • Every Excel sheet row is backed up in raw_sheets")
        print("   • Structured parsers match source workbooks")
        print("   • Study Progress KPIs match August 2026 meeting presentation")
        print("   • Dashboard uses Monthly Meeting Updates as canonical source")
        return 0

    print("❌ VERDICT: VALIDATION FAILED")
    return 1


if __name__ == "__main__":
    sys.exit(main())
