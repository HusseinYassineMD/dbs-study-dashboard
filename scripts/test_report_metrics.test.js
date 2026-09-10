#!/usr/bin/env node
/** Unit tests for monthly report metric collection (validated through current sheet data). */

const { describe, test } = require("node:test");
const assert = require("node:assert/strict");
const {
  loadReportContext,
  resolveMetricValue,
  loadExpectations,
} = require("./lib/load_report_metrics");

const ctx = loadReportContext();
const august = loadExpectations("2026-08");

function metricsFor(year, month) {
  return ctx.collectReportMetrics(year, month);
}

describe("August 2026 — team PDF golden metrics", () => {
  const m = metricsFor(august.year, august.month);
  const tolerance = august.tolerance || {};

  for (const [label, path, expected] of august.metrics) {
    test(label, () => {
      const actual = resolveMetricValue(m, path);
      assert.notEqual(actual, null, `${label} should be populated`);
      assert.notEqual(actual, undefined, `${label} should be populated`);
      const tol = Number(tolerance[path] || 0);
      assert.ok(
        Math.abs(Number(actual) - expected) <= tol,
        `${label}: expected ${expected} (±${tol}), got ${actual}`
      );
    });
  }

  test("report label", () => {
    assert.equal(m.label, august.label);
  });
});

describe("August 2026 — study progress dropouts", () => {
  test("dropout counts match meeting PDF snapshot", () => {
    const m = metricsFor(august.year, august.month);
    assert.equal(m.dropouts.baseline, 13);
    assert.equal(m.dropouts.year2, 15);
    assert.equal(m.dropouts.year3, 15);
    assert.equal(m.dropouts.year4, 0);
    assert.equal(m.dropouts.total, 43);
  });
});

describe("August 2026 — dynamic narrative", () => {
  const m = metricsFor(august.year, august.month);
  const n = m.narrative;

  test("caught-up month", () => {
    assert.equal(n.y3CaughtUpThrough, august.narrative.y3CaughtUpThrough);
  });

  test("Y3 completion deadline", () => {
    assert.equal(n.y3CompletionDeadline, august.narrative.y3CompletionDeadline);
  });

  test("Y3 remaining expected", () => {
    assert.equal(n.y3RemainingExpected, august.narrative.y3RemainingExpected);
  });

  test("advanced to Year 3 count", () => {
    assert.equal(n.advancedToY3, august.narrative.advancedToY3);
    assert.match(n.advancedToY3Text, /advanced to Year 3/);
  });

  test("Y4 rollover contacts populated from sheets", () => {
    assert.ok(n.y4Rollovers);
    assert.ok(n.y4Rollovers.count >= 65);
    assert.match(n.y4Rollovers.rangeLabel, /2027/);
  });

  test("study period and expected completion strings present", () => {
    assert.ok(n.y3StudyPeriod);
    assert.ok(n.expectedStudyCompletion);
  });
});

describe("August 2026 — chart source data", () => {
  const y3Chart = {
    "2026-07": { mri: 1, bd: 2, np: 1, cv: 2 },
    "2026-08": { mri: 5, bd: 5, np: 7, cv: 7 },
    "2026-09": { mri: 8, bd: 8, np: 10, cv: 9 },
    "2026-10": { mri: 3, bd: 3, np: 2, cv: 2 },
  };
  const y4Chart = {
    "2026-07": { bd: 4, np: 4, cv: 5 },
    "2026-08": { bd: 10, np: 10, cv: 9 },
    "2026-09": { bd: 2, np: 2, cv: 2 },
  };

  test("Y3 monthly visit series matches Study Details", () => {
    const { labels, series } = ctx.monthlyVisitSeries(2026, 8, [-1, 0, 1, 2], "year3", [
      { key: "mri", field: "mri" },
      { key: "bd", field: "bd" },
      { key: "np", field: "np" },
      { key: "cv", field: "cv" },
    ]);
    assert.equal(labels.length, 4);
    Object.keys(y3Chart).forEach((key, idx) => {
      const expected = y3Chart[key];
      assert.equal(series.mri[idx], expected.mri, `Y3 MRI ${key}`);
      assert.equal(series.bd[idx], expected.bd, `Y3 BD ${key}`);
      assert.equal(series.np[idx], expected.np, `Y3 NP ${key}`);
      assert.equal(series.cv[idx], expected.cv, `Y3 CV ${key}`);
    });
  });

  test("Y4 monthly visit series matches Study Details", () => {
    const { series } = ctx.monthlyVisitSeries(2026, 8, [-1, 0, 1], "year4", [
      { key: "bd", field: "bd" },
      { key: "np", field: "np" },
      { key: "cv", field: "cv" },
    ]);
    Object.keys(y4Chart).forEach((key, idx) => {
      const expected = y4Chart[key];
      assert.equal(series.bd[idx], expected.bd, `Y4 BD ${key}`);
      assert.equal(series.np[idx], expected.np, `Y4 NP ${key}`);
      assert.equal(series.cv[idx], expected.cv, `Y4 CV ${key}`);
    });
  });

  test("study progress chart percentages", () => {
    const m = metricsFor(2026, 8);
    const p = ctx.studyProgressChartPercents(m);
    const y2Eligible = m.baseline.completed - m.dropouts.year2;
    const y3Eligible = y2Eligible - m.dropouts.year3;
    assert.equal(p.baseline, 100);
    assert.ok(Math.abs(p.y2 - (m.y2.completed / y2Eligible) * 100) < 0.01);
    assert.ok(Math.abs(p.y3 - (m.y3.completed / y3Eligible) * 100) < 0.01);
    assert.ok(Math.abs(p.y4 - (m.y4.completed / y3Eligible) * 100) < 0.01);
  });
});

describe("Report metric invariants (any month)", () => {
  const months = ctx.availableReportMonths().slice(0, 6);

  for (const { year, month, label } of months) {
    test(`${label} — core invariants`, () => {
      const m = metricsFor(year, month);

      assert.ok(m.label);
      assert.ok(m.scheduling);
      assert.ok(m.narrative);
      assert.equal(
        m.y3TotalMonthVisits,
        m.y3Visits.mri + m.y3Visits.bloodDraw + m.y3Visits.np + m.y3Visits.cv
      );
      assert.equal(
        m.y4TotalMonthVisits,
        m.y4Visits.bloodDraw + m.y4Visits.np + m.y4Visits.cv
      );
      assert.equal(
        m.dropouts.total,
        m.dropouts.baseline + m.dropouts.year2 + m.dropouts.year3 + m.dropouts.year4
      );
      assert.ok(m.enrolled >= m.baseline.completed);
      assert.ok(m.scheduling.y3InProgress.every((g) => g.count >= 0));
      assert.ok(m.scheduling.y3ToContact.every((g) => g.count >= 0));
      assert.ok(m.scheduling.y4ToContact.every((g) => g.count >= 0));
    });
  }
});

describe("July vs August 2026 — Study Details cumulative sums", () => {
  test("each month summed from Study Details monthly columns", () => {
    const july = metricsFor(2026, 7);
    const august = metricsFor(2026, 8);
    assert.equal(july.y3.completed, 114);
    assert.equal(july.y4.completed, 18);
    assert.equal(july.scheduling.y3Components["MRI Date"], 114);
    assert.equal(july.scheduling.y3Components["BD Date"], 117);
    assert.notEqual(july.y3.completed, august.y3.completed);
    assert.notEqual(july.y3TotalMonthVisits, august.y3TotalMonthVisits);
  });
});

describe("September 2026 — live calculation (no PDF yet)", () => {
  test("metrics compute from sheet data", () => {
    const m = metricsFor(2026, 9);
    assert.equal(m.label, "September 2026");
    assert.ok(m.y3TotalMonthVisits > 0);
    assert.ok(m.y4TotalMonthVisits > 0);
    assert.ok(m.scheduling);
    assert.ok(m.narrative);
  });
});

describe("Available report months", () => {
  test("includes August 2026", () => {
    const months = ctx.availableReportMonths();
    assert.ok(months.some((m) => m.year === 2026 && m.month === 8));
  });

  test("includes September 2026", () => {
    const months = ctx.availableReportMonths();
    assert.ok(months.some((m) => m.year === 2026 && m.month === 9));
  });

  test("end month follows sheet data", () => {
    const months = ctx.availableReportMonths();
    const endKey = ctx.reportEndKey();
    assert.ok(months.length > 0);
    assert.equal(months[0].key, endKey);
  });

  test("every listed month produces metrics without throwing", () => {
    for (const { year, month } of ctx.availableReportMonths()) {
      assert.doesNotThrow(() => metricsFor(year, month));
    }
  });
});
