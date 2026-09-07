/** Load dashboard/report.js in Node for metric collection tests. */
const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..", "..");

function loadReportContext() {
  global.DATA = JSON.parse(fs.readFileSync(path.join(ROOT, "dashboard", "data.json"), "utf8"));
  const reportJs = fs.readFileSync(path.join(ROOT, "dashboard", "report.js"), "utf8");
  eval(reportJs.replace(/^async function generateMonthlyReportPPT[\s\S]*/, ""));

  return {
    DATA: global.DATA,
    collectReportMetrics,
    availableReportMonths,
    reportEndKey,
    studyDetailsMonthKey,
    studyProgressChartPercents,
    monthlyVisitSeries,
  };
}

function getPath(obj, dottedPath) {
  return dottedPath.split(".").reduce((cur, part) => {
    if (cur == null) return undefined;
    return cur[part];
  }, obj);
}

function resolveMetricValue(metrics, path) {
  if (path === "scheduling.y3InProgress" || path === "scheduling.y3ToContact" || path === "scheduling.y4ToContact") {
    const groups = getPath(metrics, path) || [];
    return groups.reduce((sum, group) => sum + (group.count || 0), 0);
  }
  return getPath(metrics, path);
}

function loadExpectations(name = "august_2026") {
  const fixturePath = path.join(ROOT, "scripts", "fixtures", "report_expectations.json");
  const fixtures = JSON.parse(fs.readFileSync(fixturePath, "utf8"));
  if (!fixtures[name]) {
    throw new Error(`Missing report expectations fixture: ${name}`);
  }
  return fixtures[name];
}

module.exports = {
  ROOT,
  loadReportContext,
  getPath,
  resolveMetricValue,
  loadExpectations,
};
