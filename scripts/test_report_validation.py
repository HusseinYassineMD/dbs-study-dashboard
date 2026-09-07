#!/usr/bin/env python3
"""Unit tests for monthly report metrics (August 2026 validated against team PDF)."""

from __future__ import annotations

import json
import subprocess
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "scripts" / "fixtures" / "report_expectations.json"
DATA = ROOT / "dashboard" / "data.json"
REPORT_JS = ROOT / "dashboard" / "report.js"


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
    if path in {"scheduling.y3InProgress", "scheduling.y3ToContact", "scheduling.y4ToContact"}:
        groups = get_path(metrics, path)
        return sum(g["count"] for g in groups)
    return get_path(metrics, path)


def load_fixture(name: str = "august_2026") -> dict:
    fixtures = json.loads(FIXTURES.read_text())
    return fixtures[name]


class August2026GoldenTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if not DATA.exists() or not REPORT_JS.exists():
            raise unittest.SkipTest("Missing dashboard/data.json or dashboard/report.js")
        cls.fixture = load_fixture("august_2026")
        cls.metrics = load_report_metrics(cls.fixture["year"], cls.fixture["month"])

    def test_report_label(self) -> None:
        self.assertEqual(self.metrics["label"], self.fixture["label"])

    def test_pdf_golden_metrics(self) -> None:
        failures: list[str] = []
        for label, path, expected in self.fixture["metrics"]:
            try:
                actual = resolve_value(self.metrics, path)
            except (KeyError, TypeError):
                actual = None
            if actual is None or int(actual) != int(expected):
                failures.append(f"{label}: expected {expected}, got {actual}")
        if failures:
            self.fail("\n".join(failures))

    def test_dropout_counts(self) -> None:
        dropouts = self.fixture["dropouts"]
        self.assertEqual(self.metrics["dropouts"]["baseline"], dropouts["Baseline"])
        self.assertEqual(self.metrics["dropouts"]["year2"], dropouts["2"])
        self.assertEqual(self.metrics["dropouts"]["year3"], dropouts["3"])
        self.assertEqual(self.metrics["dropouts"]["year4"], dropouts["4"])

    def test_narrative_fields(self) -> None:
        narrative = self.metrics["narrative"]
        expected = self.fixture["narrative"]
        self.assertEqual(narrative["y3CaughtUpThrough"], expected["y3CaughtUpThrough"])
        self.assertEqual(narrative["y3CompletionDeadline"], expected["y3CompletionDeadline"])
        self.assertEqual(narrative["y3RemainingExpected"], expected["y3RemainingExpected"])
        self.assertEqual(narrative["advancedToY3"], expected["advancedToY3"])
        self.assertGreaterEqual(narrative["y4Rollovers"]["count"], expected["y4RolloverCountMin"])
        self.assertIn("2027", narrative["y4Rollovers"]["rangeLabel"])
        self.assertTrue(narrative["y3StudyPeriod"])
        self.assertTrue(narrative["expectedStudyCompletion"])


class ReportMetricInvariantTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if not DATA.exists() or not REPORT_JS.exists():
            raise unittest.SkipTest("Missing dashboard/data.json or dashboard/report.js")
        cls.fixture = load_fixture("august_2026")
        cls.august = load_report_metrics(2026, 8)

    def test_y3_visit_total_matches_components(self) -> None:
        m = self.august
        self.assertEqual(
            m["y3TotalMonthVisits"],
            m["y3Visits"]["mri"]
            + m["y3Visits"]["bloodDraw"]
            + m["y3Visits"]["np"]
            + m["y3Visits"]["cv"],
        )

    def test_y4_visit_total_matches_components(self) -> None:
        m = self.august
        self.assertEqual(
            m["y4TotalMonthVisits"],
            m["y4Visits"]["bloodDraw"] + m["y4Visits"]["np"] + m["y4Visits"]["cv"],
        )

    def test_dropout_total_matches_sum(self) -> None:
        d = self.august["dropouts"]
        self.assertEqual(
            d["total"],
            d["baseline"] + d["year2"] + d["year3"] + d["year4"],
        )

    def test_study_details_monthly_matches_chart_fixture(self) -> None:
        data = json.loads(DATA.read_text())
        monthly = data["study_details_visit_totals"]["year3"]["monthly"]
        for key, expected in self.fixture["chart_monthly"]["year3"].items():
            row = monthly[key]
            self.assertEqual(row["mri"], expected["mri"], key)
            self.assertEqual(row["bd"], expected["bd"], key)
            self.assertEqual(row["np"], expected["np"], key)
            self.assertEqual(row["cv"], expected["cv"], key)


class AvailableMonthsSmokeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if not DATA.exists() or not REPORT_JS.exists():
            raise unittest.SkipTest("Missing dashboard/data.json or dashboard/report.js")

    def test_recent_months_collect_metrics(self) -> None:
        for year, month in [(2026, 8), (2026, 7), (2026, 6), (2025, 12)]:
            with self.subTest(year=year, month=month):
                metrics = load_report_metrics(year, month)
                self.assertIn("scheduling", metrics)
                self.assertIn("narrative", metrics)
                self.assertIsNotNone(metrics["enrolled"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
