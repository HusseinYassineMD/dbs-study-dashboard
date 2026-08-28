#!/usr/bin/env python3
"""Extract and unify DBS Tracker + Monthly Meeting Updates into dashboard JSON."""

from __future__ import annotations

import json
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
TRACKER = ROOT / "DBS Tracker.xlsx"
MEETING = ROOT / "Monthly Meeting Updates.xlsx"
OUTPUT = ROOT / "dashboard" / "data.json"


def normalize_id(value: Any) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip()
    if not text or text.lower() in {"nan", "id", "none"}:
        return None
    match = re.search(r"(?:DBS[- ]?|DBSLead[- ]?)(\d+)", text, re.I)
    if match:
        return f"DBS-{int(match.group(1)):04d}"
    if re.fullmatch(r"\d+", text):
        return f"DBS-{int(text):04d}"
    return text if text.startswith("DBS-") else None


def serialize(value: Any) -> Any:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (int, float, str, bool)):
        if isinstance(value, float) and value.is_integer():
            return int(value)
        return value
    return str(value)


def clean_columns(columns: list[Any]) -> list[str]:
    cleaned: list[str] = []
    seen: dict[str, int] = {}
    for col in columns:
        label = str(col).strip() if col is not None else "Column"
        if label.startswith("Unnamed"):
            label = "Column"
        if label in seen:
            seen[label] += 1
            label = f"{label}_{seen[label]}"
        else:
            seen[label] = 0
        cleaned.append(label)
    return cleaned


def records_from_df(df: pd.DataFrame, id_column: str | None = None) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        record = {clean_columns([c])[0] if False else k: serialize(v) for k, v in row.items()}
        if id_column and id_column in record:
            record["participant_id"] = normalize_id(record[id_column])
        rows.append(record)
    return rows


def list_sheets(path: Path) -> list[str]:
    return pd.ExcelFile(path).sheet_names


def resolve_sheet(path: Path, *candidates: str, required: bool = True) -> str | None:
    sheets = list_sheets(path)
    by_normalized = {s.strip().lower(): s for s in sheets}
    for candidate in candidates:
        if candidate in sheets:
            return candidate
        match = by_normalized.get(candidate.strip().lower())
        if match:
            return match
    if required:
        raise ValueError(
            f"Worksheet not found in {path.name}. Tried {list(candidates)}. Available: {sheets}"
        )
    return None


def read_sheet(
    path: Path, sheet: str, header: int | None = 0, required: bool = True, alt_names: list[str] | None = None
) -> pd.DataFrame:
    names = [sheet, *(alt_names or [])]
    resolved = resolve_sheet(path, *names, required=required)
    if resolved is None:
        return pd.DataFrame()
    df = pd.read_excel(path, sheet_name=resolved, header=header)
    df.columns = clean_columns(list(df.columns))
    return df


def parse_study_progress(path: Path, sheet: str = "Study Progress") -> list[dict[str, Any]]:
    df = read_sheet(path, sheet)
    rows = []
    for _, row in df.iterrows():
        year = row.get("Year")
        if pd.isna(year):
            continue
        rows.append(
            {
                "year": serialize(year),
                "enrolled": serialize(row.get("Enrolled")),
                "dropped_out_dq": serialize(row.get("Dropped Out/DQ")),
                "completed": serialize(row.get("Completed")),
                "remaining": serialize(row.get("Remaining")),
                "pct_completed_total": serialize(row.get("% Completed (of Total)")),
                "pct_completed_active": serialize(row.get("% Completed (of Active)")),
                "notes": serialize(row.get("Unnamed: 8")),
            }
        )
    return rows


def parse_recruitment(path: Path) -> dict[str, Any]:
    df = pd.read_excel(path, sheet_name="Study Details", header=None)
    header_row = df.iloc[1].tolist()
    period_columns = [serialize(c) for c in header_row[2:] if serialize(c)]
    rows = []
    for idx in range(2, len(df)):
        source = serialize(df.iloc[idx, 0])
        if not source or source == "Total":
            continue
        values = [serialize(v) for v in df.iloc[idx, 1:]]
        total = values[0] if values else None
        timeline = {}
        for col_name, val in zip(header_row[2:], df.iloc[idx, 2:]):
            key = serialize(col_name)
            if key:
                timeline[key] = serialize(val)
        rows.append({"source": source, "total_leads": total, "timeline": timeline})
    total_row = df[df.iloc[:, 0].astype(str).str.lower() == "total"]
    grand_total = serialize(total_row.iloc[0, 1]) if not total_row.empty else None
    return {"summary_n": 1735, "grand_total": grand_total, "sources": rows, "periods": period_columns}


def parse_completed_visits(path: Path) -> dict[str, list[dict[str, Any]]]:
    df = pd.read_excel(path, sheet_name="Completed Visits", header=None)
    sections = {
        "second_year": {"start": 0, "end": 5, "id": 0, "assigned": 1, "month": 2, "notes": 3, "diabetes": 4},
        "third_year": {"start": 6, "end": 9, "id": 6, "month": 7, "notes": 8},
        "fourth_year": {"start": 10, "end": 13, "id": 10, "month": 11, "notes": 12},
        "completed": {"start": 14, "end": 16, "id": 14, "month": 15, "notes": 16},
    }
    result: dict[str, list[dict[str, Any]]] = {k: [] for k in sections}
    for idx in range(2, len(df)):
        row = df.iloc[idx]
        for key, meta in sections.items():
            pid = normalize_id(row.iloc[meta["id"]])
            if not pid:
                continue
            entry: dict[str, Any] = {
                "participant_id": pid,
                "scheduling_month": serialize(row.iloc[meta["month"]]),
                "notes": serialize(row.iloc[meta["notes"]]),
            }
            if "assigned" in meta:
                entry["assigned_to"] = serialize(row.iloc[meta["assigned"]])
                entry["diabetes_status"] = serialize(row.iloc[meta["diabetes"]])
            result[key].append(entry)
    return result


def parse_genotype(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "Genotype")
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get("ID"))
        if not pid:
            continue
        rows.append(
            {
                "participant_id": pid,
                "ptau": serialize(row.get("pTau")),
                "hba1c": serialize(row.get("HbA1c")),
                "diabetic": serialize(row.get("Diabetic (≥6.5%)")),
                "prediabetic": serialize(row.get("Pre-Diabetic (6–6.4%)")),
                "non_diabetic": serialize(row.get("Non-Diabetic (<6%)")),
                "high_ad_risk": serialize(row.get("High AD Risk")),
                "low_ad_risk": serialize(row.get("Low AD Risk")),
                "genotype": serialize(row.get("Genotype")),
                "date_genotyped": serialize(row.get("Date Genotyped")),
            }
        )
    return rows


def parse_genotype_summary(path: Path) -> dict[str, Any]:
    df = read_sheet(path, "Genotype")
    first = df.iloc[0]
    return {
        "total_diabetics": serialize(first.get("TOTAL DIABETICS")),
        "total_non_diabetic": serialize(first.get("TOTAL NON-DIABETIC")),
        "total_high_ad_risk": serialize(first.get("TOTAL HIGH AD RISK")),
        "total_low_ad_risk": serialize(first.get("TOTAL LOW AD RISK")),
    }


def parse_active_ids(path: Path) -> list[str]:
    df = pd.read_excel(path, sheet_name="All active IDs", header=None)
    ids = []
    for value in df.iloc[:, 0]:
        pid = normalize_id(value)
        if pid:
            ids.append(pid)
    return sorted(set(ids))


def parse_id_location(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "ID Location")
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get("ID"))
        if not pid:
            continue
        rows.append(
            {
                "participant_id": pid,
                "cabinet_location": serialize(row.get("Cabinet Location")),
                "binder_number": serialize(row.get("Binder #")),
                "ipad": serialize(row.get("IPAD ")),
            }
        )
    return rows


def parse_screening(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "Screening Visits")
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get("ID"))
        if not pid:
            continue
        rows.append(
            {
                "participant_id": pid,
                "prescreen_date": serialize(row.get("P SV Date")),
                "staff_initials": serialize(row.get("Staff Initials")),
                "icf_signed": serialize(row.get("SV ICF Signed? (Y/N) ")),
                "screening_date": serialize(row.get("SV Date")),
                "pfaq": serialize(row.get("PFAQ")),
                "phq9": serialize(row.get("PHQ-9")),
                "mind": serialize(row.get("MIND")),
                "ecog12": serialize(row.get("ECog12")),
                "blood_draw_completed": serialize(row.get("Blood Draw Completed?")),
                "notes": serialize(row.get("Notes")),
            }
        )
    return rows


def parse_dropouts_tracker(path: Path) -> dict[str, list[dict[str, Any]]]:
    baseline_df = read_sheet(path, "DropoutsNo longer eligible")
    baseline = []
    for _, row in baseline_df.iterrows():
        pid = normalize_id(row.get("ID"))
        if not pid:
            continue
        baseline.append({k: serialize(v) for k, v in row.items()} | {"participant_id": pid})

    yearly_df = read_sheet(path, "Dropouts Y1-4")
    yearly = []
    for _, row in yearly_df.iterrows():
        pid = normalize_id(row.get("ID"))
        if not pid:
            continue
        yearly.append({k: serialize(v) for k, v in row.items()} | {"participant_id": pid})
    return {"baseline_ineligible": baseline, "year_1_to_4": yearly}


def first_column_id(row: pd.Series) -> str | None:
    for value in row:
        pid = normalize_id(value)
        if pid:
            return pid
    return None


def parse_study_partner(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "Study Partner")
    id_col = df.columns[0]
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get(id_col))
        if not pid:
            continue
        rows.append(
            {
                "participant_id": pid,
                "survey_link_sent": serialize(row.get("Survey Link Sent/Mailed")),
                "completed": serialize(row.get("Completed") or row.get("Completed ")),
                "notes": serialize(row.get("Notes") or row.get("Notes ")),
                "record_id": serialize(row.get("Record ID/location")),
            }
        )
    return rows


def parse_stool_samples(path: Path) -> dict[str, Any]:
    df = read_sheet(path, "Stool Samples")
    id_col = next((c for c in df.columns if "eligible" in str(c).lower()), df.columns[0])
    summary = {
        "complete": serialize(df.iloc[0].get("Status")),
        "total_consented": serialize(df.iloc[0].get("Total Consented")),
    }
    rows = []
    for _, row in df.iloc[1:].iterrows():
        pid = normalize_id(row.get(id_col))
        if not pid:
            continue
        rows.append(
            {
                "participant_id": pid,
                "status": serialize(row.get("Status")),
                "date_given_or_mailed": serialize(row.get("Date Given or Mailed")),
                "tracking_number": serialize(row.get("Tracking Number if mailed")),
                "kit_number": serialize(row.get("KIT Number")),
                "date_received": serialize(row.get("Date Received")),
                "follow_up_notes": serialize(row.get("Follow up notes")),
                "notes": serialize(row.get("Column") or row.get("Column_1")),
            }
        )
    return {"summary": summary, "participants": rows}


def parse_labs(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "LabsMRNs")
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get("DBS-ID"))
        if not pid:
            continue
        rows.append(
            {
                "participant_id": pid,
                "year": serialize(row.get("Unnamed: 1")),
                "mrn": serialize(row.get("MRN")),
                "lab_date": serialize(row.get("LAB DATE ")),
                "labs_year": serialize(row.get("LABS YEAR ")),
                "emailed_texted": serialize(row.get("LABS EMAILED/Texted")),
                "printed": serialize(row.get("LABS PRINTED ")),
                "mri_hct": serialize(row.get("MRI HCT ")),
                "notes": serialize(row.get("notes")),
            }
        )
    return rows


def parse_consent(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "Consent to Share ")
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get("ID"))
        if not pid:
            continue
        rows.append({"participant_id": pid, "consent_icf": serialize(row.get("Consent ICF "))})
    return rows


def parse_requests(path: Path, sheet: str, id_col: str) -> list[dict[str, Any]]:
    df = read_sheet(path, sheet)
    resolved_id_col = next(
        (c for c in df.columns if str(c).strip().lower() == id_col.strip().lower()),
        id_col,
    )
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get(resolved_id_col))
        if not pid:
            pid = first_column_id(row)
        if not pid:
            continue
        rows.append({k: serialize(v) for k, v in row.items()} | {"participant_id": pid})
    return rows


def parse_sample_dropoff(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "Sample Drop Off", header=1)
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get("ID Number:"))
        if not pid:
            continue
        rows.append({k: serialize(v) for k, v in row.items()} | {"participant_id": pid})
    return rows


def parse_synopsis(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "Synopsis Forms")
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get("ID"))
        if not pid:
            continue
        rows.append({k: serialize(v) for k, v in row.items()} | {"participant_id": pid})
    return rows


def parse_scheduling_notes(path: Path, sheet: str, header: int = 1) -> list[dict[str, Any]]:
    df = pd.read_excel(path, sheet_name=sheet, header=None)
    header_row = 0
    for idx in range(min(3, len(df))):
        row_text = " ".join(str(v).lower() for v in df.iloc[idx].tolist() if pd.notna(v))
        if "id" in row_text:
            header_row = idx
            break

    headers = clean_columns(
        [serialize(v) if pd.notna(v) else f"col_{i}" for i, v in enumerate(df.iloc[header_row])]
    )
    id_col_idx = next((i for i, h in enumerate(headers) if h.lower() == "id"), None)

    rows = []
    for idx in range(header_row + 1, len(df)):
        row = df.iloc[idx]
        pid = None
        if id_col_idx is not None:
            pid = normalize_id(row.iloc[id_col_idx])
        if not pid:
            pid = first_column_id(row)
        if not pid:
            for value in row:
                pid = normalize_id(value)
                if pid:
                    break
        if not pid:
            continue
        record = {headers[i]: serialize(row.iloc[i]) for i in range(len(headers))}
        record["participant_id"] = pid
        rows.append(record)
    return rows


def parse_np_results_requests(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "NP Results Requests")
    id_col = next((c for c in df.columns if str(c).strip().lower() == "id"), df.columns[0])
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get(id_col))
        if not pid:
            continue
        rows.append({k: serialize(v) for k, v in row.items()} | {"participant_id": pid})
    return rows


def parse_uds_id_200(path: Path) -> list[dict[str, Any]]:
    df = pd.read_excel(path, sheet_name="uds_id_200", header=None)
    rows = []
    for idx in range(1, len(df)):
        uds = df.iloc[idx, 1]
        if pd.isna(uds):
            continue
        pid = normalize_id(uds)
        rows.append(
            {
                "participant_id": pid,
                "row_index": serialize(df.iloc[idx, 0]),
                "uds_id_200": serialize(uds),
                "not_enrolled": serialize(df.iloc[idx, 3]) if df.shape[1] > 3 else None,
                "dq": serialize(df.iloc[idx, 5]) if df.shape[1] > 5 else None,
            }
        )
    return rows


def parse_dq_pts(path: Path) -> list[dict[str, Any]]:
    sheet = resolve_sheet(path, "16 DQ pts", "DQ pts", required=False)
    if not sheet:
        return []
    df = pd.read_excel(path, sheet_name=sheet, header=None)
    rows = []
    for idx in range(1, len(df)):
        val = df.iloc[idx, 0]
        if pd.isna(val):
            continue
        rows.append({"uds_id": serialize(val), "participant_id": normalize_id(val)})
    return rows


def parse_mri_outcomes(path: Path) -> list[dict[str, Any]]:
    df = pd.read_excel(path, sheet_name="Sheet 9", header=None)
    headers = clean_columns([serialize(v) if pd.notna(v) else f"col_{i}" for i, v in enumerate(df.iloc[0])])
    rows = []
    for idx in range(1, len(df)):
        row = df.iloc[idx]
        if row.isna().all():
            continue
        rows.append({headers[i]: serialize(row.iloc[i]) for i in range(len(headers))})
    return rows


def parse_raw_sheet(path: Path, sheet: str) -> list[dict[str, Any]]:
    df = pd.read_excel(path, sheet_name=sheet, header=None)
    rows = []
    for idx in range(len(df)):
        row = df.iloc[idx]
        if row.isna().all():
            continue
        record = {f"col_{i}": serialize(row.iloc[i]) for i in range(df.shape[1])}
        record["row_number"] = idx + 1
        pid = first_column_id(row)
        if pid:
            record["participant_id"] = pid
        rows.append(record)
    return rows


def parse_all_raw_sheets() -> dict[str, list[dict[str, Any]]]:
    raw: dict[str, list[dict[str, Any]]] = {}
    for path, prefix in [(TRACKER, "tracker"), (MEETING, "meeting")]:
        for sheet in pd.ExcelFile(path).sheet_names:
            key = f"{prefix}::{sheet.strip()}"
            raw[key] = parse_raw_sheet(path, sheet)
    return raw


def parse_mri_dates(path: Path) -> list[dict[str, Any]]:
    df = pd.read_excel(path, sheet_name="MRI Dates Available", header=None)
    rows = []
    current_month = None
    for idx in range(1, len(df)):
        month = serialize(df.iloc[idx, 0])
        if month:
            current_month = month
        rows.append(
            {
                "month": current_month,
                "mri_date": serialize(df.iloc[idx, 1]),
                "mri_time": serialize(df.iloc[idx, 2]),
                "notes": serialize(df.iloc[idx, 3]),
            }
        )
    return [r for r in rows if r["mri_date"]]


def parse_clincard(path: Path) -> list[dict[str, Any]]:
    df = pd.read_excel(path, sheet_name="ClinCard", header=1)
    df.columns = clean_columns(list(df.columns))
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get("ID"))
        if not pid:
            continue
        rows.append({k: serialize(v) for k, v in row.items()} | {"participant_id": pid})
    return rows


def parse_pt_status(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "Pt Status")
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get("uds_id"))
        if not pid:
            continue
        rows.append(
            {
                "participant_id": pid,
                "dropout_date": serialize(row.get("dropout_date")),
                "dropout_year": serialize(row.get("dropout_yr")),
                "dropout_visit": serialize(row.get("dropout_viist")),
                "reason_dropout": serialize(row.get("reason_dropout")),
            }
        )
    return rows


def parse_meeting_dropouts(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "DropOuts")
    rows = []
    for _, row in df.iterrows():
        pid = normalize_id(row.get("Baseline Drop Outs"))
        if not pid:
            continue
        rows.append(
            {
                "participant_id": pid,
                "reason": serialize(row.get("Outreach Efforts / Reason")),
            }
        )
    return rows


def parse_second_year_monthly(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "2nd year")
    rows = []
    for _, row in df.iterrows():
        visit_type = serialize(row.get("Second Year Visits"))
        if visit_type not in {"Clinician Visit", "NP"}:
            continue
        entry = {"visit_type": visit_type, "total_completed": serialize(row.get("Total Completed"))}
        for col in df.columns:
            if col in {"Second Year Visits", "Total Completed"}:
                continue
            entry[col] = serialize(row.get(col))
        rows.append(entry)
    return rows


def parse_visit_trends(path: Path, sheet: str) -> list[dict[str, Any]]:
    df = read_sheet(path, sheet)
    rows = []
    for _, row in df.iterrows():
        if pd.isna(row.get("Date")):
            continue
        rows.append({k: serialize(v) for k, v in row.items()})
    return rows


def parse_enrolled_mapping(path: Path) -> list[dict[str, Any]]:
    df = read_sheet(path, "200 Enrolled pts")
    rows = []
    for _, row in df.iterrows():
        uds = row.get("uds_id_200")
        if pd.isna(uds):
            continue
        pid = normalize_id(uds)
        rows.append(
            {
                "participant_id": pid,
                "uds_id": serialize(uds),
                "not_enrolled": serialize(row.get("Not enrolled")),
                "dq": serialize(row.get("DQ")),
            }
        )
    return rows


def build_participant_index(data: dict[str, Any]) -> list[dict[str, Any]]:
    participants: dict[str, dict[str, Any]] = {}

    def ensure(pid: str) -> dict[str, Any]:
        if pid not in participants:
            participants[pid] = {"participant_id": pid}
        return participants[pid]

    for pid in data["active_participants"]:
        ensure(pid)["status"] = "Active"

    for row in data["pt_status"]:
        pid = row["participant_id"]
        p = ensure(pid)
        p["pt_status"] = row
        if row.get("dropout_date") or row.get("reason_dropout"):
            p["status"] = "Dropped Out"

    for row in data["id_location"]:
        ensure(row["participant_id"]).update(
            {
                "cabinet_location": row.get("cabinet_location"),
                "binder_number": row.get("binder_number"),
                "ipad": row.get("ipad"),
            }
        )

    for row in data["genotype"]:
        ensure(row["participant_id"])["genotype"] = row

    for row in data["consent"]:
        ensure(row["participant_id"])["consent"] = row.get("consent_icf")

    for row in data["study_partner"]:
        ensure(row["participant_id"])["study_partner"] = row

    for row in data["stool_samples"]["participants"]:
        ensure(row["participant_id"])["stool_sample"] = row

    for row in data["labs"]:
        ensure(row["participant_id"]).setdefault("labs", []).append(row)

    for row in data["screening"]:
        ensure(row["participant_id"]).setdefault("screening", []).append(row)

    for section, rows in data["completed_visits"].items():
        for row in rows:
            ensure(row["participant_id"]).setdefault("completed_visits", {}).setdefault(section, row)

    for row in data["dropouts"]["baseline_ineligible"] + data["dropouts"]["year_1_to_4"]:
        p = ensure(row["participant_id"])
        p.setdefault("dropout_records", []).append(row)

    for row in data["meeting_dropouts"]:
        ensure(row["participant_id"])["meeting_dropout_reason"] = row.get("reason")

    for row in data.get("baseline_scheduling", []):
        ensure(row["participant_id"]).setdefault("baseline_scheduling", row)

    for row in data.get("second_year_scheduling", []):
        ensure(row["participant_id"]).setdefault("second_year_scheduling", row)

    for row in data.get("third_year_scheduling", []):
        ensure(row["participant_id"]).setdefault("third_year_scheduling", row)

    for row in data.get("fourth_year_scheduling", []):
        ensure(row["participant_id"]).setdefault("fourth_year_scheduling", row)

    for row in data.get("clincard", []):
        ensure(row["participant_id"]).setdefault("clincard", row)

    for row in data.get("np_results_requests", []):
        ensure(row["participant_id"]).setdefault("np_results_requests", row)

    for row in data.get("uds_id_200", []):
        if row.get("participant_id"):
            ensure(row["participant_id"])["uds_id_200"] = row

    return sorted(participants.values(), key=lambda x: x["participant_id"])


def sheet_catalog() -> list[dict[str, str]]:
    catalog = []
    for path, label in [(TRACKER, "DBS Tracker"), (MEETING, "Monthly Meeting Updates")]:
        xl = pd.ExcelFile(path)
        for sheet in xl.sheet_names:
            catalog.append({"workbook": label, "sheet": sheet, "file": path.name})
    return catalog


def main() -> None:
    if not TRACKER.exists() or not MEETING.exists():
        raise FileNotFoundError("Expected both Excel workbooks in project root.")

    payload: dict[str, Any] = {
        "meta": {
            "title": "Diabetes Brain Study (DBS) Operations Dashboard",
            "subtitle": "Unified view of participant tracking, recruitment, and visit progress",
            "generated_at": datetime.now().isoformat(),
            "sources": ["DBS Tracker.xlsx", "Monthly Meeting Updates.xlsx"],
        },
        "study_progress_tracker": parse_study_progress(TRACKER),
        "study_progress_meeting": parse_study_progress(MEETING),
        "recruitment": parse_recruitment(MEETING),
        "active_participants": parse_active_ids(TRACKER),
        "id_location": parse_id_location(TRACKER),
        "completed_visits": parse_completed_visits(TRACKER),
        "genotype": parse_genotype(TRACKER),
        "genotype_summary": parse_genotype_summary(TRACKER),
        "screening": parse_screening(TRACKER),
        "dropouts": parse_dropouts_tracker(TRACKER),
        "study_partner": parse_study_partner(TRACKER),
        "stool_samples": parse_stool_samples(TRACKER),
        "labs": parse_labs(TRACKER),
        "consent": parse_consent(TRACKER),
        "lab_results_requests": parse_requests(TRACKER, "Lab Results Requests", "ID"),
        "np_results_requests": parse_np_results_requests(TRACKER),
        "mri_cd_requests": parse_requests(TRACKER, "MRI CD Requests", "ID"),
        "sample_drop_off": parse_sample_dropoff(TRACKER),
        "synopsis_forms": parse_synopsis(TRACKER),
        "clincard": parse_clincard(TRACKER),
        "mri_dates_available": parse_mri_dates(TRACKER),
        "baseline_scheduling": parse_scheduling_notes(TRACKER, "Baseline Scheduling Notes"),
        "second_year_scheduling": parse_scheduling_notes(TRACKER, "2nd Yr Scheduling Notes"),
        "third_year_scheduling": parse_scheduling_notes(TRACKER, "3rd Yr Scheduling Notes "),
        "fourth_year_scheduling": parse_scheduling_notes(TRACKER, "4th Yr Scheduling Notes"),
        "pt_status": parse_pt_status(MEETING),
        "meeting_dropouts": parse_meeting_dropouts(MEETING),
        "second_year_monthly": parse_second_year_monthly(MEETING),
        "visit_trends_sheet5": parse_visit_trends(MEETING, "Sheet5"),
        "visit_trends_sheet6": parse_visit_trends(MEETING, "Sheet6"),
        "enrolled_mapping": parse_enrolled_mapping(MEETING),
        "uds_id_200": parse_uds_id_200(TRACKER),
        "dq_pts": parse_dq_pts(MEETING),
        "mri_outcomes": parse_mri_outcomes(MEETING),
        "raw_sheets": parse_all_raw_sheets(),
        "sheet_catalog": sheet_catalog(),
    }

    payload["participants"] = build_participant_index(payload)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2, ensure_ascii=False)

    print(f"Wrote {OUTPUT} ({OUTPUT.stat().st_size // 1024} KB)")
    print(f"Participants indexed: {len(payload['participants'])}")


if __name__ == "__main__":
    main()
