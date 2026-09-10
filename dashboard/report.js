/** Monthly progress report → PowerPoint generated from live sheet data for any report month. */

const DECK = {
  bg: "EDE1E2",
  cardinal: "991B1E",
  gold: "FFCC00",
  titleOnRed: "FFFFFF",
  title: "000000",
  body: "000000",
  muted: "444444",
  tableBorder: "991B1E",
  font: "Arial",
  marginL: 0.5,
  contentW: 9.0,
  headerH: 1.17,
  titleHeaderH: 0.85,
  footerH: 0.5,
};

/** Executive summary grid — matches August PDF (Y3/Y4 top, dropouts/future bottom). */
const EXEC_GRID = {
  leftX: 0.55,
  rightX: 4.85,
  colW: 4.35,
  topY: 1.05,
  bottomY: 2.95,
  numSize: 21,
  labelSize: 21,
  bulletSize: 14,
};

let reportLogoData = null;

async function getReportLogoData() {
  if (reportLogoData) return reportLogoData;
  const res = await fetch("assets/usc-logo.jpg");
  if (!res.ok) return null;
  const blob = await res.blob();
  reportLogoData = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
  return reportLogoData;
}

const CHART = {
  y3: { mri: "#4472C4", bd: "#C00000", np: "#FFC000", cv: "#70AD47" },
  y4: { bd: "#4472C4", np: "#C00000", cv: "#FFC000" },
  progress: {
    baseline: "#006400",
    y2: "#006400",
    y3: "#92D050",
    y4: "#E57373",
  },
};

const barValueLabelPlugin = {
  id: "barValueLabels",
  afterDatasetsDraw(chart, _args, opts) {
    const { ctx } = chart;
    const color = opts?.color || "#000000";
    const offset = opts?.offset ?? 4;
    const formatter = opts?.formatter || ((v) => String(v));
    chart.data.datasets.forEach((_dataset, i) => {
      const meta = chart.getDatasetMeta(i);
      if (meta.hidden) return;
      meta.data.forEach((bar, index) => {
        const value = chart.data.datasets[i].data[index];
        if (value == null || value === 0) return;
        ctx.save();
        ctx.fillStyle = color;
        ctx.font = `bold ${opts?.fontSize || 11}px Arial`;
        ctx.textAlign = "center";
        ctx.textBaseline = "bottom";
        ctx.fillText(formatter(value), bar.x, bar.y - offset);
        ctx.restore();
      });
    });
  },
};

function chartMonthAxisLabel(year, month, { expected = false } = {}) {
  const label = `${MONTH_NAMES[month - 1]} '${String(year).slice(-2)}`;
  return expected ? `${label} (expected)` : label;
}

function groupedBarYMax(values) {
  const max = Math.max(0, ...values.filter((v) => v != null));
  return Math.max(10, Math.ceil(max / 2) * 2);
}

function renderChartToDataUrl(config, width = 680, height = 440) {
  return new Promise((resolve, reject) => {
    if (typeof Chart === "undefined") {
      reject(new Error("Chart.js not loaded."));
      return;
    }
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    canvas.style.position = "fixed";
    canvas.style.left = "-9999px";
    document.body.appendChild(canvas);
    try {
      const chart = new Chart(canvas.getContext("2d"), {
        ...config,
        options: {
          ...config.options,
          responsive: false,
          animation: false,
          maintainAspectRatio: false,
        },
      });
      requestAnimationFrame(() => {
        const dataUrl = canvas.toDataURL("image/png");
        chart.destroy();
        canvas.remove();
        resolve(dataUrl);
      });
    } catch (err) {
      canvas.remove();
      reject(err);
    }
  });
}

function studyProgressChartPercents(m) {
  const y2Eligible = m.baseline.completed - m.dropouts.year2;
  const y3Eligible = y2Eligible - m.dropouts.year3;
  const pctOf = (num, denom) => (denom > 0 ? (num / denom) * 100 : 0);
  return {
    baseline: 100,
    y2: pctOf(m.y2.completed, y2Eligible),
    y3: pctOf(m.y3.completed, y3Eligible),
    y4: pctOf(m.y4.completed, y3Eligible),
  };
}

function monthlyVisitSeries(reportYear, reportMonth, offsets, yearKey, fields) {
  const totals = DATA.study_details_visit_totals?.[yearKey]?.monthly || {};
  const labels = [];
  const series = Object.fromEntries(fields.map((f) => [f.key, []]));

  offsets.forEach((offset) => {
    const { year, month } = addMonths(reportYear, reportMonth, offset);
    const row = totals[studyDetailsMonthKey(year, month)] || {};
    const expected = offset === 1;
    labels.push(chartMonthAxisLabel(year, month, { expected }));
    fields.forEach(({ key, field }) => {
      series[key].push(row[field] ?? 0);
    });
  });

  return { labels, series };
}

async function buildYear3VisitsChartImage(reportYear, reportMonth) {
  const { labels, series } = monthlyVisitSeries(reportYear, reportMonth, [-1, 0, 1, 2], "year3", [
    { key: "mri", field: "mri" },
    { key: "bd", field: "bd" },
    { key: "np", field: "np" },
    { key: "cv", field: "cv" },
  ]);
  const datasets = [
    { label: "MRI", data: series.mri, backgroundColor: CHART.y3.mri },
    { label: "BD", data: series.bd, backgroundColor: CHART.y3.bd },
    { label: "NP", data: series.np, backgroundColor: CHART.y3.np },
    { label: "Clinician Visit", data: series.cv, backgroundColor: CHART.y3.cv },
  ];
  const yMax = groupedBarYMax(datasets.flatMap((d) => d.data));

  return renderChartToDataUrl({
    type: "bar",
    data: { labels, datasets },
    plugins: [barValueLabelPlugin],
    options: {
      plugins: {
        barValueLabels: { color: "#000000", fontSize: 11, offset: 4 },
        legend: { position: "top", labels: { boxWidth: 14, font: { size: 11 } } },
        title: {
          display: true,
          text: "Third Year Visits",
          font: { size: 13, weight: "normal" },
          color: "#666666",
          align: "start",
        },
      },
      layout: { padding: { top: 16 } },
      scales: {
        x: { ticks: { font: { size: 10 }, maxRotation: 0, autoSkip: false } },
        y: {
          beginAtZero: true,
          max: yMax,
          grace: "8%",
          ticks: { stepSize: 2, font: { size: 10 } },
          title: { display: true, text: "Number of visits", font: { size: 11 } },
        },
      },
    },
  });
}

async function buildYear4VisitsChartImage(reportYear, reportMonth) {
  const { labels, series } = monthlyVisitSeries(reportYear, reportMonth, [-1, 0, 1], "year4", [
    { key: "bd", field: "bd" },
    { key: "np", field: "np" },
    { key: "cv", field: "cv" },
  ]);
  const datasets = [
    { label: "BD", data: series.bd, backgroundColor: CHART.y4.bd },
    { label: "NP", data: series.np, backgroundColor: CHART.y4.np },
    { label: "CV", data: series.cv, backgroundColor: CHART.y4.cv },
  ];
  const yMax = groupedBarYMax(datasets.flatMap((d) => d.data));

  return renderChartToDataUrl({
    type: "bar",
    data: { labels, datasets },
    plugins: [barValueLabelPlugin],
    options: {
      plugins: {
        barValueLabels: { color: "#000000", fontSize: 11, offset: 4 },
        legend: { position: "top", labels: { boxWidth: 14, font: { size: 11 } } },
        title: {
          display: true,
          text: "Fourth Year Visits",
          font: { size: 13, weight: "normal" },
          color: "#666666",
          align: "start",
        },
      },
      layout: { padding: { top: 16 } },
      scales: {
        x: { ticks: { font: { size: 10 }, maxRotation: 0, autoSkip: false } },
        y: {
          beginAtZero: true,
          max: yMax,
          grace: "8%",
          ticks: { stepSize: 2, font: { size: 10 } },
          title: { display: true, text: "Number of visits", font: { size: 11 } },
        },
      },
    },
  });
}

async function buildStudyProgressChartImage(m) {
  const p = studyProgressChartPercents(m);
  const values = [p.baseline, p.y2, p.y3, p.y4];
  const colors = [CHART.progress.baseline, CHART.progress.y2, CHART.progress.y3, CHART.progress.y4];
  const fmtPct = (v) => `${v.toFixed(2)}%`;

  return renderChartToDataUrl({
    type: "bar",
    data: {
      labels: ["Baseline", "2", "3", "4"],
      datasets: [{ data: values, backgroundColor: colors, barPercentage: 0.55 }],
    },
    plugins: [barValueLabelPlugin],
    options: {
      plugins: {
        barValueLabels: {
          color: "#000000",
          fontSize: 10,
          offset: 4,
          formatter: fmtPct,
        },
        legend: { display: false },
        title: {
          display: true,
          text: "Study Progress",
          font: { size: 13, weight: "normal" },
          color: "#666666",
          align: "start",
        },
      },
      layout: { padding: { top: 16 } },
      scales: {
        x: {
          title: { display: true, text: "Year", font: { size: 11 } },
          ticks: { font: { size: 11 } },
        },
        y: {
          beginAtZero: true,
          max: 100,
          grace: "8%",
          ticks: {
            stepSize: 25,
            callback: (v) => `${v.toFixed(2)}%`,
            font: { size: 10 },
          },
          title: { display: true, text: "% Completed (of Active)", font: { size: 11 } },
        },
      },
    },
  });
}

function addSlideChartImage(slide, dataUrl, { x = 5.05, y = 1.1, w = 4.55, h = 4.2 } = {}) {
  if (!dataUrl) return;
  slide.addImage({ data: dataUrl, x, y, w, h });
}

function slideBg(slide) {
  slide.background = { color: DECK.bg };
}

function addTopHeaderBar(slide, title, fontSize = 34) {
  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: "100%",
    h: DECK.headerH,
    fill: { color: DECK.cardinal },
    line: { color: DECK.cardinal, width: 0 },
  });
  if (title) {
    slide.addText(title, {
      x: 0.4,
      y: 0.12,
      w: 9.2,
      h: DECK.headerH - 0.15,
      fontSize,
      bold: true,
      color: DECK.titleOnRed,
      fontFace: DECK.font,
      align: "center",
      valign: "middle",
      wrap: true,
    });
  }
}

function addTitlePageHeader(slide) {
  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: "100%",
    h: DECK.titleHeaderH,
    fill: { color: DECK.cardinal },
    line: { color: DECK.cardinal, width: 0 },
  });
}

function addBottomFooterBar(slide) {
  slide.addShape("rect", {
    x: 0,
    y: 5.625 - DECK.footerH,
    w: "100%",
    h: DECK.footerH,
    fill: { color: DECK.cardinal },
    line: { color: DECK.cardinal, width: 0 },
  });
}

function addCenteredTitle(slide, title, fontSize = 34) {
  slide.addText(title, {
    x: 0.5,
    y: 0.35,
    w: 9.0,
    h: 0.7,
    fontSize,
    bold: true,
    color: DECK.title,
    fontFace: DECK.font,
    align: "center",
    wrap: true,
  });
}
const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

const MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/** Earliest report month (study start). End month is derived from sheet data. */
const REPORT_START_KEY = 2024 * 12 + 5;

function reportEndKey() {
  const keys = [];
  ["year3", "year4"].forEach((yearKey) => {
    Object.keys(DATA.study_details_visit_totals?.[yearKey]?.monthly || {}).forEach((key) => {
      const [year, month] = key.split("-").map(Number);
      keys.push(monthKey(year, month));
    });
  });
  [DATA.third_year_scheduling, DATA.fourth_year_scheduling].forEach((rows) => {
    (rows || []).forEach((row) => {
      const due = parseDueMonth(row.Month, null, null);
      const key = dueMonthKey(due);
      if (key != null) keys.push(key);
    });
  });
  ["second_year", "third_year", "fourth_year"].forEach((key) => {
    (DATA.completed_visits?.[key] || []).forEach((row) => {
      const due = parseDueMonth(row.scheduling_month, null, null);
      const keyVal = dueMonthKey(due);
      if (keyVal != null) keys.push(keyVal);
    });
  });
  if (DATA.meta?.generated_at) {
    const d = new Date(DATA.meta.generated_at);
    if (!Number.isNaN(d.getTime())) keys.push(monthKey(d.getFullYear(), d.getMonth() + 1));
  }
  const now = new Date();
  const nowKey = monthKey(now.getFullYear(), now.getMonth() + 1);
  keys.push(nowKey);
  const raw = keys.length ? Math.max(...keys) : nowKey;
  return Math.min(raw, nowKey);
}

function monthKey(year, month) {
  return year * 12 + month;
}

function displayCount(value) {
  if (value == null || value === "") return 0;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function fmtCount(value) {
  if (value == null || value === "") return "0";
  return String(value);
}

function parseIsoDate(value) {
  if (value == null || value === "") return null;
  const s = String(value).trim();
  if (!s || /opt out|n\/a|^nan$|ineligible|^tbd$/i.test(s)) return null;
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}

function inCalendarMonth(value, year, month) {
  const d = parseIsoDate(value);
  return d && d.getFullYear() === year && d.getMonth() + 1 === month;
}

function monthLabel(year, month) {
  return `${MONTH_NAMES[month - 1]} ${year}`;
}

function shortMonthLabel(year, month) {
  return `${MONTH_ABBR[month - 1]} ${year}`;
}

function addMonths(year, month, delta) {
  const key = monthKey(year, month) + delta;
  return { year: Math.floor((key - 1) / 12), month: ((key - 1) % 12) + 1 };
}

function pct(value) {
  if (value == null) return "—";
  return `${Math.round(Number(value) * 1000) / 10}%`;
}

function getMeetingProgress() {
  return DATA.study_progress_meeting?.length ? DATA.study_progress_meeting : DATA.study_progress_tracker;
}

function progressRow(year) {
  return getMeetingProgress().find((r) => String(r.year).toLowerCase() === String(year).toLowerCase()) || {};
}

function countFieldInMonth(rows, field, year, month) {
  return rows.filter((r) => inCalendarMonth(r[field], year, month)).length;
}

function countFieldsInMonth(rows, fields, year, month) {
  return fields.reduce((acc, field) => {
    acc[field] = countFieldInMonth(rows, field, year, month);
    return acc;
  }, {});
}

function countFieldsInMonths(rows, fields, monthPairs) {
  return fields.reduce((acc, field) => {
    acc[field] = monthPairs.reduce(
      (sum, [y, m]) => sum + countFieldInMonth(rows, field, y, m),
      0
    );
    return acc;
  }, {});
}

function countCompletedVisitsInMonth(yearKey, year, month) {
  return (DATA.completed_visits?.[yearKey] || []).filter((r) =>
    inCalendarMonth(r.scheduling_month, year, month)
  ).length;
}

function findMonthlyColumn(monthlyRows, year, month) {
  if (!monthlyRows?.length) return null;
  const full = MONTH_NAMES[month - 1];
  const abbr = full.slice(0, 3);
  const yy = String(year).slice(-2);
  const keys = Object.keys(monthlyRows[0]).filter(
    (k) => !["visit_type", "total_completed"].includes(k)
  );
  return (
    keys.find((k) => k.toLowerCase().includes(full.toLowerCase()) && k.includes(`'${yy}`)) ||
    keys.find((k) => k.toLowerCase().startsWith(abbr.toLowerCase()) && k.includes(`'${yy}`)) ||
    keys.find((k) => k.toLowerCase().startsWith(abbr.toLowerCase()) && k.includes(yy)) ||
    null
  );
}

function lastTrendOnOrBefore(sheetKey, year, month) {
  const target = monthKey(year, month);
  let best = null;
  let bestKey = -1;
  (DATA[sheetKey] || []).forEach((r) => {
    const d = parseIsoDate(r.Date);
    if (!d) return;
    const key = monthKey(d.getFullYear(), d.getMonth() + 1);
    if (key <= target && key > bestKey) {
      bestKey = key;
      best = r;
    }
  });
  return best;
}

function visitTrendForMonth(sheetKey, year, month) {
  return (DATA[sheetKey] || []).find((r) => inCalendarMonth(r.Date, year, month)) || null;
}

function rowNotes(row) {
  return Object.values(row)
    .filter((v) => typeof v === "string")
    .join(" ");
}

function dropoutParticipantIds() {
  const rows = DATA.dropouts?.year_1_to_4 || [];
  return new Set(rows.map((row) => row.participant_id || row.ID).filter(Boolean));
}

function isDroppedRow(row) {
  if (/\bDROPP+ED\b/i.test(rowNotes(row))) return true;
  const id = row.participant_id || row.ID;
  return id ? dropoutParticipantIds().has(id) : false;
}

function hasMriExemption(row) {
  if (parseIsoDate(row["MRI Date"])) return true;
  return /opt out|ineligible|Inelligible|MRI OPT OUT/i.test(rowNotes(row));
}

function isYear3Complete(row) {
  if (isDroppedRow(row)) return false;
  return (
    parseIsoDate(row["BD Date"]) &&
    parseIsoDate(row["NP Date"]) &&
    parseIsoDate(row["CV Date"]) &&
    hasMriExemption(row)
  );
}

function isYear4Complete(row) {
  if (isDroppedRow(row)) return false;
  return (
    parseIsoDate(row["BD Date"]) &&
    parseIsoDate(row["NP Date"]) &&
    parseIsoDate(row["CV Date"])
  );
}

function reportMonthEnd(reportYear, reportMonth) {
  return new Date(reportYear, reportMonth, 0, 23, 59, 59, 999);
}

function schedulingCompleteAsOf(row, visitFields, reportYear, reportMonth) {
  if (isDroppedRow(row)) return true;
  const end = reportMonthEnd(reportYear, reportMonth);
  for (const field of visitFields) {
    if (field === "MRI Date") {
      if (!hasMriExemption(row)) return false;
      const mri = parseIsoDate(row["MRI Date"]);
      if (mri && mri > end) return false;
      continue;
    }
    const d = parseIsoDate(row[field]);
    if (!d || d > end) return false;
  }
  return true;
}

function adjustDueMonthYear(due, refYear, refMonth) {
  if (!due || refYear == null || refMonth == null) return due;
  const inProgressMin = addMonths(refYear, refMonth, -2);
  const minKey = monthKey(inProgressMin.year, inProgressMin.month);
  const key = monthKey(due.year, due.month);
  if (key >= minKey) return due;
  // Scheduling sheet typos: Nov/Dec entered with prior year (e.g. 2025-12 for a 2026 due date).
  if (due.year === refYear - 1 && due.month >= 11) {
    return { year: refYear, month: due.month, day: due.day };
  }
  return due;
}

function normalizeY4SchedulingDue(due, row) {
  if (!due || due.year !== 2026 || due.month !== 4) return due;
  const hasVisits = ["BD Date", "NP Date", "CV Date"].some((f) => parseIsoDate(row[f]));
  if (hasVisits) return due;
  const day = due.day || 1;
  if (day <= 8) return { year: 2026, month: 11, day };
  return { year: 2026, month: 12, day };
}

function resolveSchedulingDue(row, refYear, refMonth, visitFields, { year4 = false } = {}) {
  let due = parseDueMonth(row.Month, refYear, refMonth);
  if (year4) due = normalizeY4SchedulingDue(due, row);
  return due;
}

function parseDueMonth(value, refYear, refMonth) {
  const d = parseIsoDate(value);
  if (d) {
    return adjustDueMonthYear(
      { year: d.getFullYear(), month: d.getMonth() + 1, day: d.getDate() },
      refYear,
      refMonth
    );
  }
  if (value == null || value === "") return null;
  const s = String(value).trim();
  for (let i = 0; i < MONTH_NAMES.length; i += 1) {
    const full = MONTH_NAMES[i];
    if (s.toLowerCase() === full.toLowerCase() || s.toLowerCase() === full.slice(0, 3).toLowerCase()) {
      let year = refYear;
      if (refMonth != null && i + 1 < refMonth - 1) year = refYear + 1;
      return adjustDueMonthYear({ year, month: i + 1 }, refYear, refMonth);
    }
  }
  return null;
}

function dueMonthKey(due) {
  return due ? monthKey(due.year, due.month) : null;
}

function dueMonthLabel(due) {
  return due ? `${MONTH_NAMES[due.month - 1]} ${due.year}` : "—";
}

function completionSortKey(row) {
  const dates = ["MRI Date", "BD Date", "NP Date", "CV Date"]
    .map((f) => parseIsoDate(row[f]))
    .filter(Boolean);
  if (!dates.length) return 0;
  return Math.max(...dates.map((d) => d.getTime()));
}

function completedVisitComponents(rows, completeFn, fields, targetCount) {
  const complete = rows.filter(completeFn).sort((a, b) => completionSortKey(a) - completionSortKey(b));
  const slice = targetCount ? complete.slice(0, targetCount) : complete;
  const counts = {};
  fields.forEach((field) => {
    if (field === "MRI Date") {
      counts[field] = slice.filter(
        (r) => parseIsoDate(r[field]) || (/ineligible|Inelligible/i.test(rowNotes(r)) && !parseIsoDate(r[field]))
      ).length;
      return;
    }
    counts[field] = slice.filter((r) => parseIsoDate(r[field])).length;
  });
  counts.mriExemptions = slice.filter((r) => !parseIsoDate(r["MRI Date"]) && hasMriExemption(r)).length;
  return counts;
}

/** Executive summary component totals from Study Details (Monthly Meeting Updates). */
function studyDetailsMonthKey(year, month) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

function studyDetailsMonthVisits(year, month) {
  const totals = DATA.study_details_visit_totals;
  const key = studyDetailsMonthKey(year, month);
  const y3 = totals?.year3?.monthly?.[key];
  const y4 = totals?.year4?.monthly?.[key];
  if (!y3 && !y4) return null;
  return {
    y3: {
      mri: y3?.mri ?? 0,
      bloodDraw: y3?.bd ?? 0,
      np: y3?.np ?? 0,
      cv: y3?.cv ?? 0,
    },
    y4: {
      bloodDraw: y4?.bd ?? 0,
      np: y4?.np ?? 0,
      cv: y4?.cv ?? 0,
    },
  };
}

function studyDetailsFutureVisits(reportYear, reportMonth) {
  const totals = DATA.study_details_visit_totals;
  if (!totals?.year3?.monthly) return null;
  const months = [addMonths(reportYear, reportMonth, 1), addMonths(reportYear, reportMonth, 2)];
  const keys = months.map(({ year, month }) => studyDetailsMonthKey(year, month));

  const sum = (field, includeY4 = false) =>
    keys.reduce((acc, key) => {
      acc += totals.year3.monthly[key]?.[field] || 0;
      if (includeY4) acc += totals.year4.monthly?.[key]?.[field] || 0;
      return acc;
    }, 0);

  return {
    mri: sum("mri", false),
    bd: sum("bd", true),
    np: sum("np", true),
    cv: sum("cv", true),
  };
}

/** Sum Study Details monthly visit columns through the report month (Monthly Meeting Updates). */
function studyDetailsCumulativeThrough(reportYear, reportMonth) {
  const totals = DATA.study_details_visit_totals;
  if (!totals?.year3?.monthly) return null;
  const endKey = studyDetailsMonthKey(reportYear, reportMonth);

  const sumMonthly = (monthly, fields) => {
    const out = {};
    Object.keys(monthly || {})
      .filter((key) => key <= endKey)
      .forEach((key) => {
        fields.forEach((field) => {
          out[field] = (out[field] || 0) + (monthly[key][field] || 0);
        });
      });
    return out;
  };

  return {
    y3: sumMonthly(totals.year3.monthly, ["mri", "bd", "np", "cv"]),
    y4: sumMonthly(totals.year4?.monthly, ["bd", "np", "cv"]),
  };
}

function studyDetailsComponentMaps(reportYear, reportMonth) {
  const cumulative = studyDetailsCumulativeThrough(reportYear, reportMonth);
  if (!cumulative?.y3 || cumulative.y3.mri == null) return null;
  const totals = DATA.study_details_visit_totals;
  return {
    y3: {
      "MRI Date": cumulative.y3.mri ?? 0,
      "BD Date": cumulative.y3.bd ?? 0,
      "NP Date": cumulative.y3.np ?? 0,
      "CV Date": cumulative.y3.cv ?? 0,
      mriExemptions: totals?.year3?.mri_exemptions ?? null,
    },
    y4: {
      "BD Date": cumulative.y4?.bd ?? 0,
      "NP Date": cumulative.y4?.np ?? 0,
      "CV Date": cumulative.y4?.cv ?? 0,
    },
  };
}

function yearCompletedFromStudyDetails(reportYear, reportMonth) {
  const cumulative = studyDetailsCumulativeThrough(reportYear, reportMonth);
  if (!cumulative?.y3) return null;
  const y3Values = [cumulative.y3.mri, cumulative.y3.bd, cumulative.y3.np, cumulative.y3.cv];
  const y4Values = [cumulative.y4?.bd, cumulative.y4?.np, cumulative.y4?.cv];
  if (y3Values.some((v) => v == null) || y4Values.some((v) => v == null)) return null;
  return {
    y3: Math.min(...y3Values),
    y4: Math.min(...y4Values),
  };
}

/** Participant finished all required visits on or before the report month end (excludes dropouts). */
function yearSchedulingCompleteAsOf(row, visitFields, reportYear, reportMonth) {
  if (isDroppedRow(row)) return false;
  const end = reportMonthEnd(reportYear, reportMonth);
  for (const field of visitFields) {
    if (field === "MRI Date") {
      if (!hasMriExemption(row)) return false;
      const mri = parseIsoDate(row["MRI Date"]);
      if (mri && mri > end) return false;
      continue;
    }
    const d = parseIsoDate(row[field]);
    if (!d || d > end) return false;
  }
  return true;
}

function countYearCompletedAsOf(rows, visitFields, reportYear, reportMonth) {
  return rows.filter((row) => yearSchedulingCompleteAsOf(row, visitFields, reportYear, reportMonth)).length;
}

function expectedCompletionText(due, reportYear, reportMonth) {
  if (!due) return "TBD";
  const dueKey = dueMonthKey(due);
  const reportKey = monthKey(reportYear, reportMonth);
  if (dueKey <= reportKey) {
    const next = addMonths(reportYear, reportMonth, 1);
    const next2 = addMonths(reportYear, reportMonth, 2);
    return `TBD – expected late ${MONTH_NAMES[next.month - 1]}/${MONTH_ABBR[next2.month - 1]} ${next2.year}`;
  }
  const early = addMonths(due.year, due.month, 1);
  return `${MONTH_NAMES[due.month - 1]}–early ${MONTH_NAMES[early.month - 1]} ${early.year}`;
}

function dueKeysFromRows(rows, refYear, refMonth, { year4 = false } = {}) {
  const keys = [];
  const visitFields = year4
    ? ["BD Date", "NP Date", "CV Date"]
    : ["MRI Date", "BD Date", "NP Date", "CV Date"];
  (rows || []).forEach((row) => {
    let due = null;
    const iso = parseIsoDate(row.Month);
    if (iso) {
      due = adjustDueMonthYear(
        { year: iso.getFullYear(), month: iso.getMonth() + 1, day: iso.getDate() },
        refYear,
        refMonth
      );
    } else if (refYear != null) {
      due = resolveSchedulingDue(row, refYear, refMonth, visitFields, { year4 });
    }
    const key = dueMonthKey(due);
    if (key != null && due.year >= 2024) keys.push(key);
  });
  return keys;
}

function dueKeysFromCompletedVisits(yearKey) {
  const keys = [];
  (DATA.completed_visits?.[yearKey] || []).forEach((row) => {
    const due = parseDueMonth(row.scheduling_month, null, null);
    if (due?.year >= 2024) {
      const key = dueMonthKey(due);
      if (key != null) keys.push(key);
    }
  });
  return keys;
}

function formatMonthYearFromKey(key) {
  const year = Math.floor((key - 1) / 12);
  const month = ((key - 1) % 12) + 1;
  return `${MONTH_NAMES[month - 1]} ${year}`;
}

function formatPeriodEndLabel(keys) {
  if (!keys.length) return null;
  const maxKey = Math.max(...keys);
  const maxYear = Math.floor((maxKey - 1) / 12);
  const maxMonth = ((maxKey - 1) % 12) + 1;
  const hasJan = keys.some((k) => Math.floor((k - 1) / 12) === maxYear && ((k - 1) % 12) + 1 === 1);
  const hasFeb = keys.some((k) => Math.floor((k - 1) / 12) === maxYear && ((k - 1) % 12) + 1 === 2);
  if (hasJan && hasFeb && maxMonth <= 2) return `Jan/Feb ${maxYear}`;
  return formatMonthYearFromKey(maxKey);
}

function formatStudyCompletionLabel(keys) {
  if (!keys.length) return null;
  const maxKey = Math.max(...keys);
  const maxYear = Math.floor((maxKey - 1) / 12);
  const decKey = monthKey(maxYear, 12);
  const janKey = monthKey(maxYear + 1, 1);
  const hasDec = keys.includes(decKey);
  const hasJan = keys.includes(janKey);
  if (hasDec && hasJan) return `Dec ${maxYear}/Jan ${maxYear + 1}`;
  return formatMonthYearFromKey(maxKey);
}

function maxY3ContactDueKey() {
  const keys = dueKeysFromRows(DATA.third_year_scheduling || [], null, null);
  return keys.length ? Math.max(...keys) : null;
}

function computeY3CaughtUpThrough(reportYear, reportMonth) {
  const third = DATA.third_year_scheduling || [];
  const y3Fields = ["MRI Date", "BD Date", "NP Date", "CV Date"];
  let earliestIncompleteKey = null;

  third.forEach((row) => {
    if (isDroppedRow(row)) return;
    if (schedulingCompleteAsOf(row, y3Fields, reportYear, reportMonth)) return;
    const due = resolveSchedulingDue(row, reportYear, reportMonth, y3Fields);
    const key = dueMonthKey(due);
    if (key == null) return;
    if (earliestIncompleteKey == null || key < earliestIncompleteKey) earliestIncompleteKey = key;
  });

  if (earliestIncompleteKey == null) {
    return formatMonthYearFromKey(monthKey(reportYear, reportMonth));
  }
  const dueYear = Math.floor((earliestIncompleteKey - 1) / 12);
  const dueMonth = ((earliestIncompleteKey - 1) % 12) + 1;
  const caught = addMonths(dueYear, dueMonth, -1);
  return formatMonthYearFromKey(monthKey(caught.year, caught.month));
}

function computeY3StudyPeriod() {
  const keys = dueKeysFromRows(DATA.third_year_scheduling || [], null, null);
  if (!keys.length) return null;
  const minKey = Math.min(...keys);
  return `${formatMonthYearFromKey(minKey)} - ${formatPeriodEndLabel(keys)}`;
}

function computeY3CompletionDeadline() {
  const keys = dueKeysFromRows(DATA.third_year_scheduling || [], null, null);
  return formatPeriodEndLabel(keys);
}

function computeExpectedStudyCompletion() {
  const keys = [
    ...dueKeysFromRows(DATA.fourth_year_scheduling || [], null, null, { year4: true }),
    ...dueKeysFromCompletedVisits("fourth_year"),
  ];
  return formatStudyCompletionLabel(keys);
}

function countAdvancedToYear3(baseline, y2) {
  const activeAfterBaseline = (baseline.completed ?? 0) - (y2.dropped_out_dq ?? 0);
  const activeCohort = y2.remaining ?? activeAfterBaseline;
  return Math.max(0, activeCohort - (y2.completed ?? 0));
}

function collectY4RolloverContacts(reportYear, reportMonth) {
  const nextYear = reportYear + 1;
  const y4Fields = ["BD Date", "NP Date", "CV Date"];
  const fourth = DATA.fourth_year_scheduling || [];
  const fourthById = new Map(fourth.map((row) => [row.participant_id, row]));
  const seen = new Set();
  const dueMonths = [];

  const consider = (participantId, dueValue) => {
    if (!participantId || seen.has(participantId)) return;
    if (isDroppedRow({ participant_id: participantId, ID: participantId })) return;
    const sched = fourthById.get(participantId);
    if (sched && schedulingCompleteAsOf(sched, y4Fields, reportYear, reportMonth)) return;
    const due = parseDueMonth(dueValue, reportYear, reportMonth);
    if (!due || due.year !== nextYear) return;
    seen.add(participantId);
    dueMonths.push(due.month);
  };

  (DATA.completed_visits?.fourth_year || []).forEach((row) => {
    consider(row.participant_id, row.scheduling_month);
  });
  fourth.forEach((row) => {
    if (isDroppedRow(row)) return;
    if (schedulingCompleteAsOf(row, y4Fields, reportYear, reportMonth)) return;
    consider(row.participant_id, row.Month);
  });

  if (!seen.size) return null;
  dueMonths.sort((a, b) => a - b);
  const start = MONTH_NAMES[dueMonths[0] - 1];
  const end = MONTH_NAMES[dueMonths[dueMonths.length - 1] - 1];
  const rangeLabel =
    dueMonths[0] === dueMonths[dueMonths.length - 1]
      ? `${start} ${nextYear}`
      : `${start}–${end} ${nextYear}`;
  return { count: seen.size, rangeLabel, nextYear };
}

function collectReportNarrative(reportYear, reportMonth, baseline, y2, y3) {
  const advancedToY3 = countAdvancedToYear3(baseline, y2);
  const y4Rollovers = collectY4RolloverContacts(reportYear, reportMonth);
  return {
    y3CaughtUpThrough: computeY3CaughtUpThrough(reportYear, reportMonth),
    y3StudyPeriod: computeY3StudyPeriod(),
    y3CompletionDeadline: computeY3CompletionDeadline(),
    y3RemainingExpected: y3.remaining ?? null,
    expectedStudyCompletion: computeExpectedStudyCompletion(),
    advancedToY3,
    advancedToY3Text:
      advancedToY3 === 1
        ? "+1 participant advanced to Year 3"
        : advancedToY3 > 1
          ? `+${advancedToY3} participants advanced to Year 3`
          : null,
    y4Rollovers,
    questions: DATA.meta?.report_questions || [],
  };
}

function groupIncompleteByDue(rows, visitFields, refYear, refMonth, { inProgressMaxKey, contactMinKey, contactMaxKey, year4 = false }) {
  const inProgress = new Map();
  const toContact = new Map();

  rows.forEach((row) => {
    if (isDroppedRow(row)) return;
    if (schedulingCompleteAsOf(row, visitFields, refYear, refMonth)) return;
    const due = resolveSchedulingDue(row, refYear, refMonth, visitFields, { year4 });
    const key = dueMonthKey(due);
    if (key == null) return;
    const label = dueMonthLabel(due);
    let target = null;
    if (key <= inProgressMaxKey) target = inProgress;
    else if (contactMinKey != null && key >= contactMinKey && (contactMaxKey == null || key <= contactMaxKey)) {
      target = toContact;
    }
    if (!target) return;
    if (!target.has(label)) {
      target.set(label, { label, due, count: 0, rows: [] });
    }
    const entry = target.get(label);
    entry.count += 1;
    entry.rows.push(row);
  });

  const sortEntries = (entries) =>
    entries.sort((a, b) => dueMonthKey(a.due) - dueMonthKey(b.due));

  return {
    inProgress: sortEntries([...inProgress.values()]),
    toContact: sortEntries([...toContact.values()]),
  };
}

function participantsDueInMonth(rows, refYear, refMonth) {
  return rows.filter((row) => {
    if (isDroppedRow(row)) return false;
    const due = parseDueMonth(row.Month, refYear, refMonth);
    return due && due.year === refYear && due.month === refMonth;
  });
}

function hasAnyScheduledDate(row, fields) {
  return fields.some((f) => parseIsoDate(row[f]));
}

function hasScheduledInOrAfter(row, fields, year, month) {
  const cutoff = new Date(year, month - 1, 1);
  return fields.some((f) => {
    const d = parseIsoDate(row[f]);
    return d && d >= cutoff;
  });
}

function completedParticipantIds(rows, completeFn, cap) {
  const ids = new Set();
  if (!cap) return ids;
  rows
    .filter(completeFn)
    .sort((a, b) => completionSortKey(a) - completionSortKey(b))
    .slice(0, cap)
    .forEach((row) => ids.add(row.participant_id));
  return ids;
}

function completedAllVisitsInMonth(row, visitFields, reportYear, reportMonth) {
  const bdnp = ["BD Date", "NP Date", "CV Date"].filter((f) => visitFields.includes(f));
  if (!bdnp.every((f) => inCalendarMonth(row[f], reportYear, reportMonth))) return false;
  if (visitFields.includes("MRI Date") && inCalendarMonth(row["MRI Date"], reportYear, reportMonth)) return true;
  return visitFields.includes("MRI Date") ? hasMriExemption(row) : true;
}

function analyzeYearActivity(rows, completeFn, visitFields, reportYear, reportMonth, completedCap = 0, { dueCompleteInDueMonth = false } = {}) {
  const monthVisits = countFieldsInMonth(rows, visitFields, reportYear, reportMonth);
  const visitTotal = visitFields.reduce((s, f) => s + (monthVisits[f] || 0), 0);
  const visitSummary = visitFields
    .map((f) => {
      const label = f.replace(" Date", "");
      return `${monthVisits[f] || 0} ${label === "BD" ? "BD" : label === "NP" ? "NP" : label === "CV" ? "CV" : "MRI"}`;
    })
    .join(", ");

  const dueRows = participantsDueInMonth(rows, reportYear, reportMonth);
  const dueComplete = dueRows.filter((row) =>
    dueCompleteInDueMonth
      ? completedAllVisitsInMonth(row, visitFields, reportYear, reportMonth)
      : schedulingCompleteAsOf(row, visitFields, reportYear, reportMonth)
  ).length;
  const dueScheduled = dueRows.filter(
    (row) =>
      !(dueCompleteInDueMonth
        ? completedAllVisitsInMonth(row, visitFields, reportYear, reportMonth)
        : schedulingCompleteAsOf(row, visitFields, reportYear, reportMonth)) &&
      hasScheduledInOrAfter(row, visitFields, reportYear, reportMonth + 1)
  ).length;

  const next = addMonths(reportYear, reportMonth, 1);
  const nextVisits = countFieldsInMonth(rows, visitFields, next.year, next.month);
  const nextTotal = visitFields.reduce((s, f) => s + (nextVisits[f] || 0), 0);
  const nextSummary = visitFields
    .map((f) => `${nextVisits[f] || 0} ${f.replace(" Date", "")}`)
    .join(", ");

  const nextDueRows = participantsDueInMonth(rows, next.year, next.month);
  const nextDueCompletedEarly = nextDueRows.filter((row) =>
    visitFields.some((field) => inCalendarMonth(row[field], reportYear, reportMonth))
  ).length;
  const nextDueScheduled = Math.max(0, nextDueRows.length - nextDueCompletedEarly);

  return {
    visitTotal,
    visitSummary,
    dueCount: dueComplete + dueScheduled,
    dueComplete,
    dueScheduled,
    nextMonth: next,
    nextVisitTotal: nextTotal,
    nextVisitSummary: nextSummary,
    nextDueCount: nextDueRows.length,
    nextDueCompletedEarly,
    nextDueScheduled,
  };
}

function collectSchedulingReportData(reportYear, reportMonth, y3Completed, y4Completed) {
  const third = DATA.third_year_scheduling || [];
  const fourth = DATA.fourth_year_scheduling || [];
  const y3Fields = ["MRI Date", "BD Date", "NP Date", "CV Date"];
  const y4Fields = ["BD Date", "NP Date", "CV Date"];

  const inProgressMin = addMonths(reportYear, reportMonth, -2);
  const inProgressMinKey = monthKey(inProgressMin.year, inProgressMin.month);
  const inProgressMax = addMonths(reportYear, reportMonth, 2);
  const inProgressMaxKey = monthKey(inProgressMax.year, inProgressMax.month);
  const contactMin = addMonths(reportYear, reportMonth, 3);
  const contactMinKey = monthKey(contactMin.year, contactMin.month);
  const contactMaxKey = maxY3ContactDueKey();

  const y3Groups = groupIncompleteByDue(third, y3Fields, reportYear, reportMonth, {
    inProgressMaxKey,
    contactMinKey,
    contactMaxKey,
  });
  const y3InProgress = y3Groups.inProgress.filter(
    (g) => dueMonthKey(g.due) >= inProgressMinKey && dueMonthKey(g.due) <= inProgressMaxKey
  );

  const contactWindowStart = addMonths(reportYear, reportMonth, 1);
  const y4ContactStart = monthKey(contactWindowStart.year, contactWindowStart.month);
  const y4ContactEnd = monthKey(reportYear, 12);
  const y4ToContactMap = new Map();
  const fourthById = new Map(fourth.map((row) => [row.participant_id, row]));
  const addY4ToContact = (participantId, dueValue) => {
    if (!participantId || isDroppedRow({ participant_id: participantId, ID: participantId })) return;
    const sched = fourthById.get(participantId);
    if (sched && schedulingCompleteAsOf(sched, y4Fields, reportYear, reportMonth)) return;
    const due = parseDueMonth(dueValue, reportYear, reportMonth);
    const key = dueMonthKey(due);
    if (key == null || key < y4ContactStart || key > y4ContactEnd) return;
    const label = dueMonthLabel(due);
    if (!y4ToContactMap.has(label)) {
      y4ToContactMap.set(label, { label, due, count: 0, rows: [] });
    }
    const entry = y4ToContactMap.get(label);
    if (entry.rows.some((r) => r.participant_id === participantId)) return;
    entry.count += 1;
    entry.rows.push(sched || { participant_id: participantId, ID: participantId });
  };
  (DATA.completed_visits?.fourth_year || []).forEach((row) => {
    addY4ToContact(row.participant_id, row.scheduling_month);
  });
  fourth.forEach((row) => {
    if (isDroppedRow(row)) return;
    if (schedulingCompleteAsOf(row, y4Fields, reportYear, reportMonth)) return;
    addY4ToContact(row.participant_id, row.Month);
  });
  const y4ToContact = [...y4ToContactMap.values()].sort(
    (a, b) => dueMonthKey(a.due) - dueMonthKey(b.due)
  );

  const futureMonths = [
    addMonths(reportYear, reportMonth, 1),
    addMonths(reportYear, reportMonth, 2),
  ].map(({ year, month }) => [year, month]);

  const y3Future = countFieldsInMonths(third, y3Fields, futureMonths);
  const y4Future = countFieldsInMonths(fourth, y4Fields, futureMonths);
  const fromStudyDetailsFuture = studyDetailsFutureVisits(reportYear, reportMonth);
  const fromStudyDetailsComponents = studyDetailsComponentMaps(reportYear, reportMonth);
  const y3Components =
    fromStudyDetailsComponents?.y3 ??
    completedVisitComponents(third, isYear3Complete, y3Fields, y3Completed);
  const y4Components =
    fromStudyDetailsComponents?.y4 ??
    completedVisitComponents(fourth, isYear4Complete, y4Fields, y4Completed);

  const futureVisits = fromStudyDetailsFuture ?? {
    mri: y3Future["MRI Date"] || 0,
    bd: (y3Future["BD Date"] || 0) + (y4Future["BD Date"] || 0),
    np: (y3Future["NP Date"] || 0) + (y4Future["NP Date"] || 0),
    cv: (y3Future["CV Date"] || 0) + (y4Future["CV Date"] || 0),
  };

  let y3Activity = analyzeYearActivity(third, isYear3Complete, y3Fields, reportYear, reportMonth, y3Completed, {
    dueCompleteInDueMonth: true,
  });
  let y4Activity = analyzeYearActivity(fourth, isYear4Complete, y4Fields, reportYear, reportMonth, y4Completed);
  const sdMonth = studyDetailsMonthVisits(reportYear, reportMonth);
  const sdNext = studyDetailsMonthVisits(
    addMonths(reportYear, reportMonth, 1).year,
    addMonths(reportYear, reportMonth, 1).month
  );
  if (sdMonth?.y3) {
    const m = sdMonth.y3;
    const total = m.mri + m.bloodDraw + m.np + m.cv;
    y3Activity = {
      ...y3Activity,
      visitTotal: total,
      visitSummary: `${m.mri} MRI, ${m.bloodDraw} BD, ${m.np} NP, ${m.cv} CV`,
    };
  }
  if (sdNext?.y3) {
    const n = sdNext.y3;
    const nextTotal = n.mri + n.bloodDraw + n.np + n.cv;
    y3Activity = {
      ...y3Activity,
      nextVisitTotal: nextTotal,
      nextVisitSummary: `${n.mri} MRI, ${n.bloodDraw} BD, ${n.np} NP, ${n.cv} CV`,
    };
  }
  if (sdMonth?.y4) {
    const m = sdMonth.y4;
    const total = m.bloodDraw + m.np + m.cv;
    y4Activity = {
      ...y4Activity,
      visitTotal: total,
      visitSummary: `${m.bloodDraw} BD, ${m.np} NP, ${m.cv} CV`,
    };
  }
  if (sdNext?.y4) {
    const n = sdNext.y4;
    const visitSum = n.bloodDraw + n.np + n.cv;
    const nextTotal =
      n.bloodDraw === n.np && n.np === n.cv && n.bloodDraw > 0 ? n.bloodDraw : visitSum;
    y4Activity = {
      ...y4Activity,
      nextVisitTotal: nextTotal,
      nextVisitSummary: `${n.bloodDraw} BD, ${n.np} NP, ${n.cv} CV`,
    };
  }

  return {
    y3Activity,
    y4Activity,
    y3InProgress,
    y3ToContact: y3Groups.toContact,
    y4ToContact,
    y4Rollovers: collectY4RolloverContacts(reportYear, reportMonth),
    futureVisits,
    y3Components,
    y4Components,
    asOfDate: `${MONTH_NAMES[reportMonth - 1]} ${reportYear}`,
  };
}

function collectReportMetrics(year, month) {
  const baseline = progressRow("baseline");
  const y2 = progressRow("2");
  const thirdSched = DATA.third_year_scheduling || [];
  const fourthSched = DATA.fourth_year_scheduling || [];
  const y3Fields = ["MRI Date", "BD Date", "NP Date", "CV Date"];
  const y4Fields = ["BD Date", "NP Date", "CV Date"];
  const sdCompleted = yearCompletedFromStudyDetails(year, month);
  const y3CompletedAsOf =
    sdCompleted?.y3 ?? countYearCompletedAsOf(thirdSched, y3Fields, year, month);
  const y4CompletedAsOf =
    sdCompleted?.y4 ?? countYearCompletedAsOf(fourthSched, y4Fields, year, month);
  const y3 = { ...progressRow("3"), completed: y3CompletedAsOf };
  const y4 = { ...progressRow("4"), completed: y4CompletedAsOf };

  const sdVisits = studyDetailsMonthVisits(year, month);
  const y3FromSched = countFieldsInMonth(thirdSched, ["MRI Date", "BD Date", "NP Date", "CV Date"], year, month);
  const y4FromSched = countFieldsInMonth(fourthSched, ["BD Date", "NP Date", "CV Date"], year, month);

  const y3Visits = sdVisits
    ? {
        "MRI Date": sdVisits.y3.mri,
        "BD Date": sdVisits.y3.bloodDraw,
        "NP Date": sdVisits.y3.np,
        "CV Date": sdVisits.y3.cv,
      }
    : y3FromSched;
  const y4Visits = sdVisits
    ? {
        "BD Date": sdVisits.y4.bloodDraw,
        "NP Date": sdVisits.y4.np,
        "CV Date": sdVisits.y4.cv,
      }
    : y4FromSched;

  const cvRow = DATA.second_year_monthly?.find((r) => r.visit_type === "Clinician Visit");
  const npRow = DATA.second_year_monthly?.find((r) => r.visit_type === "NP");
  const monthCol = findMonthlyColumn(DATA.second_year_monthly, year, month);
  const y2Monthly = {
    column: monthCol,
    cv: monthCol && cvRow ? displayCount(cvRow[monthCol]) : null,
    np: monthCol && npRow ? displayCount(npRow[monthCol]) : null,
  };

  const trend5 = visitTrendForMonth("visit_trends_sheet5", year, month);
  const trend6 = visitTrendForMonth("visit_trends_sheet6", year, month);
  const lastTrend5 = lastTrendOnOrBefore("visit_trends_sheet5", year, month);

  const y2cv = y2Monthly.cv != null ? y2Monthly.cv : displayCount(trend5?.["CV Monthly"]);
  const y2np = y2Monthly.np != null ? y2Monthly.np : displayCount(trend5?.["NP Monthly"]);
  const y2Cumulative =
    trend5?.["CV Cumulative"] ??
    trend6?.["CV Cumulative"] ??
    lastTrend5?.["CV Cumulative"] ??
    null;
  const y2CumulativeCarried = !trend5 && !trend6 && lastTrend5 != null;

  const dropouts = {
    baseline: baseline.dropped_out_dq ?? 0,
    year2: y2.dropped_out_dq ?? 0,
    year3: y3.dropped_out_dq ?? 0,
    year4: y4.dropped_out_dq ?? 0,
    total: [baseline, y2, y3, y4].reduce((s, r) => s + (r.dropped_out_dq || 0), 0),
  };

  const y3TotalMonthVisits =
    (y3Visits["MRI Date"] || 0) +
    (y3Visits["BD Date"] || 0) +
    (y3Visits["NP Date"] || 0) +
    (y3Visits["CV Date"] || 0);
  const y4TotalMonthVisits =
    (y4Visits["BD Date"] || 0) + (y4Visits["NP Date"] || 0) + (y4Visits["CV Date"] || 0);

  const scheduling = collectSchedulingReportData(year, month, y3CompletedAsOf, y4CompletedAsOf);
  const narrative = collectReportNarrative(year, month, baseline, y2, y3);

  const metrics = {
    label: monthLabel(year, month),
    year,
    month,
    generatedAt: DATA.meta?.generated_at,
    baseline,
    y2,
    y3,
    y4,
    active: DATA.active_participants?.length ?? 0,
    enrolled: baseline.enrolled ?? 0,
    dropouts,
    narrative,
    y3Visits: {
      mri: y3Visits["MRI Date"] || 0,
      bloodDraw: y3Visits["BD Date"] || 0,
      np: y3Visits["NP Date"] || 0,
      cv: y3Visits["CV Date"] || 0,
    },
    y4Visits: {
      bloodDraw: y4Visits["BD Date"] || 0,
      np: y4Visits["NP Date"] || 0,
      cv: y4Visits["CV Date"] || 0,
    },
    y2Monthly,
    y2cv,
    y2np,
    y2Cumulative,
    y2CumulativeCarried,
    trend5,
    trend6,
    y2CompletedMonth: countCompletedVisitsInMonth("second_year", year, month),
    y3CompletedMonth: countCompletedVisitsInMonth("third_year", year, month),
    y4CompletedMonth: countCompletedVisitsInMonth("fourth_year", year, month),
    y3TotalMonthVisits,
    y4TotalMonthVisits,
    monthlyVisitTotal: y2cv + y2np + y3TotalMonthVisits + y4TotalMonthVisits,
    cumulativeY2AtMonth: y2Cumulative,
    recruitment: DATA.recruitment || {},
    scheduling,
  };

  return metrics;
}

function asOfDateShort(year, month) {
  return `${month}/28/${year}`;
}

function mriFootnoteFromData() {
  const ineligible = [];
  const optOut = [];
  (DATA.third_year_scheduling || []).forEach((row) => {
    const pid = row.participant_id || row.ID;
    if (!pid || parseIsoDate(row["MRI Date"])) return;
    const notes = rowNotes(row);
    if (/ineligible|Inelligible/i.test(notes)) ineligible.push(pid);
    else if (/opt out|opted out|OPT OUT/i.test(notes)) optOut.push(pid);
  });
  const parts = [];
  if (ineligible[0]) parts.push(`${ineligible[0]} ineligible for MRI`);
  if (optOut.length) parts.push(`${optOut.slice(0, 2).join(" and ")} Opted out of MRI`);
  if (!parts.length) return null;
  return `* ${parts.join(" ; ")}`;
}

function dashCell(value) {
  return value != null && value !== 0 ? String(value) : "–";
}

function addSlideTitle(slide, title, subtitle) {
  slideBg(slide);
  addTopHeaderBar(slide, title);
  if (subtitle) {
    slide.addText(subtitle, {
      x: DECK.marginL,
      y: DECK.headerH + 0.08,
      w: DECK.contentW,
      h: 0.35,
      fontSize: 16,
      bold: true,
      color: DECK.body,
      fontFace: DECK.font,
      wrap: true,
    });
  }
}

function contentTop(subtitle = false) {
  return subtitle ? DECK.headerH + 0.5 : DECK.headerH + 0.2;
}

function addPlainText(slide, text, opts = {}) {
  slide.addText(text, {
    x: opts.x ?? DECK.marginL,
    y: opts.y ?? contentTop(),
    w: opts.w ?? DECK.contentW,
    h: opts.h ?? 0.8,
    fontSize: opts.fontSize ?? 18,
    bold: opts.bold ?? false,
    italic: opts.italic ?? false,
    color: opts.color ?? DECK.body,
    fontFace: DECK.font,
    align: opts.align ?? "left",
    valign: opts.valign ?? "top",
    wrap: true,
  });
}

function addBodyLines(slide, lines, opts = {}) {
  const parts = lines.map((line) => {
    if (typeof line === "string") {
      return { text: line, options: { breakLine: true } };
    }
    const options = { breakLine: true };
    if (line.bullet) options.bullet = true;
    if (line.bold) options.bold = true;
    if (line.fontSize) options.fontSize = line.fontSize;
    return { text: line.text, options };
  });
  slide.addText(parts, {
    x: opts.x ?? DECK.marginL,
    y: opts.y ?? contentTop(),
    w: opts.w ?? DECK.contentW,
    h: opts.h ?? 4.0,
    fontSize: opts.fontSize ?? 14,
    color: DECK.body,
    fontFace: DECK.font,
    valign: "top",
    wrap: true,
  });
}

function addDataTable(slide, rows, opts = {}) {
  slide.addTable(rows, {
    x: opts.x ?? DECK.marginL,
    y: opts.y ?? contentTop(true),
    w: opts.w ?? DECK.contentW,
    colW: opts.colW,
    fontSize: opts.fontSize ?? 12,
    border: { type: "solid", color: DECK.tableBorder, pt: 0.75 },
    fontFace: DECK.font,
    color: DECK.body,
    autoPage: false,
    rowH: opts.rowH ?? 0.35,
  });
}

function headerCell(text) {
  return {
    text,
    options: {
      bold: true,
      fill: { color: DECK.cardinal },
      color: DECK.titleOnRed,
    },
  };
}

/** Single text box per quadrant — avoids overlapping text boxes. */
function addQuadrantBlock(slide, x, y, w, lines, h = 1.85) {
  const parts = [];
  lines.forEach((line, idx) => {
    if (typeof line === "string") {
      parts.push({
        text: line,
        options: { breakLine: true, fontSize: idx < 3 ? EXEC_GRID.labelSize : EXEC_GRID.bulletSize, bold: idx < 3 },
      });
      return;
    }
    const options = { breakLine: true, fontSize: line.fontSize ?? EXEC_GRID.bulletSize };
    if (line.bold) options.bold = true;
    if (line.bullet) options.bullet = true;
    if (line.fontSize) options.fontSize = line.fontSize;
    parts.push({ text: line.text, options });
  });
  slide.addText(parts, {
    x,
    y,
    w,
    h,
    color: DECK.body,
    fontFace: DECK.font,
    valign: "top",
    wrap: true,
    fit: "shrink",
  });
}

function buildTitleSlide(pptx, m, logoData) {
  const slide = pptx.addSlide();
  slideBg(slide);
  addTitlePageHeader(slide);
  const textX = 0.6;
  const textW = 8.8;
  addPlainText(slide, "Diabetes Brain Study –", {
    x: textX,
    y: 1.35,
    w: textW,
    h: 0.7,
    fontSize: 40,
    bold: true,
    align: "center",
  });
  addPlainText(slide, "Monthly Progress Meeting", {
    x: textX,
    y: 2.05,
    w: textW,
    h: 0.65,
    fontSize: 40,
    bold: true,
    align: "center",
  });
  addPlainText(slide, m.label, {
    x: textX,
    y: 2.75,
    w: textW,
    h: 0.55,
    fontSize: 40,
    bold: true,
    align: "center",
  });
  if (logoData) {
    const logoW = 2.8;
    const logoH = 0.72;
    slide.addImage({
      data: logoData,
      x: (10 - logoW) / 2,
      y: 3.55,
      w: logoW,
      h: logoH,
    });
  }
}

function buildExecutiveSummarySlide(pptx, m) {
  const slide = pptx.addSlide();
  slideBg(slide);
  addCenteredTitle(slide, "Executive Summary");
  const s = m.scheduling;
  const next = addMonths(m.year, m.month, 1);
  const next2 = addMonths(m.year, m.month, 2);
  const g = EXEC_GRID;

  addQuadrantBlock(slide, g.leftX, g.topY, g.colW, [
    fmtCount(m.y3.completed),
    "Participants",
    "Completed 3rd Year",
    `${fmtCount(s.y3Components["MRI Date"])} MRI *`,
    `${fmtCount(s.y3Components["BD Date"])} Blood Draw`,
    `${fmtCount(s.y3Components["NP Date"])} Neuropsych Tests`,
    `${fmtCount(s.y3Components["CV Date"])} Clinician Visits`,
  ].map((text, i) => (i >= 3 ? { text, bullet: true } : text)), 2.35);

  addQuadrantBlock(slide, g.rightX, g.topY, g.colW, [
    fmtCount(m.y4.completed),
    "Participants",
    "Completed 4th Year",
    `${fmtCount(s.y4Components["BD Date"])} Blood Draw`,
    `${fmtCount(s.y4Components["NP Date"])} Neuropsych Tests`,
    `${fmtCount(s.y4Components["CV Date"])} Clinician Visits`,
  ].map((text, i) => (i >= 3 ? { text, bullet: true } : text)), 2.35);

  addQuadrantBlock(slide, g.leftX, g.bottomY, g.colW, [
    "Future Visits Scheduled",
    `${MONTH_ABBR[next.month - 1]} - ${MONTH_ABBR[next2.month - 1]} (3rd and 4th year)`,
    { text: `${fmtCount(s.futureVisits.mri)} MRI's`, bullet: true },
    { text: `${fmtCount(s.futureVisits.bd)} Full Blood Draws`, bullet: true },
    { text: `${fmtCount(s.futureVisits.np)} Neuropsych Tests`, bullet: true },
    { text: `${fmtCount(s.futureVisits.cv)} Clinician Visits`, bullet: true },
  ], 2.45);

  addQuadrantBlock(slide, g.rightX, g.bottomY, g.colW, [
    fmtCount(m.dropouts.total),
    "Dropout",
    "Participants",
    { text: `${fmtCount(m.dropouts.baseline)} Baseline year`, bullet: true },
    { text: `${fmtCount(m.dropouts.year2)} Second year`, bullet: true },
    { text: `${fmtCount(m.dropouts.year3)} Third year`, bullet: true },
    { text: `${fmtCount(m.dropouts.year4)} Fourth year`, bullet: true },
  ], 2.45);

  const footnote = mriFootnoteFromData();
  if (footnote) {
    addPlainText(slide, footnote, {
      x: DECK.marginL,
      y: 4.85,
      w: DECK.contentW,
      h: 0.45,
      fontSize: 12,
      italic: true,
      color: DECK.muted,
    });
  }
  addBottomFooterBar(slide);
}

function buildCompletedVisitsSlide(pptx, m) {
  const slide = pptx.addSlide();
  const monthName = MONTH_NAMES[m.month - 1];
  addSlideTitle(slide, `All Completed Visits in ${monthName}`);

  addDataTable(
    slide,
    [
      [headerCell("Visit Type"), headerCell("3rd Year"), headerCell("4th Year")],
      ["MRI", dashCell(m.y3Visits.mri), "–"],
      ["Blood Draw", dashCell(m.y3Visits.bloodDraw), dashCell(m.y4Visits.bloodDraw)],
      ["NP Visit", dashCell(m.y3Visits.np), dashCell(m.y4Visits.np)],
      ["Clinician Visit", dashCell(m.y3Visits.cv), dashCell(m.y4Visits.cv)],
    ],
    { y: 1.55, colW: [3.4, 2.6, 2.6], fontSize: 14, rowH: 0.42 }
  );

  const pending = (DATA.third_year_scheduling || []).find((row) => {
    const d = parseIsoDate(row["MRI Date"] || row["BD Date"]);
    return d && d.getFullYear() === m.year && d.getMonth() + 1 === m.month && d.getDate() >= 28;
  });
  if (pending?.participant_id) {
    const d = parseIsoDate(pending["MRI Date"] || pending["BD Date"]);
    const weekday = d.toLocaleDateString("en-US", { weekday: "long" });
    const md = `${d.getMonth() + 1}/${String(d.getDate()).padStart(2, "0")}/${String(d.getFullYear()).slice(-2)}`;
    addPlainText(slide, `${pending.participant_id} expected ${weekday} ${md} to complete BD and MRI visit.`, {
      y: 4.15,
      fontSize: 14,
      h: 0.5,
    });
  }
}

function buildYear3ActivitySlide(pptx, m, chartDataUrl) {
  const slide = pptx.addSlide();
  const monthName = MONTH_NAMES[m.month - 1];
  const monthAbbr = monthName.slice(0, 3).toLowerCase();
  const a = m.scheduling.y3Activity;
  const nextName = MONTH_NAMES[a.nextMonth.month - 1];
  const y3Total = m.y3Visits.mri + m.y3Visits.bloodDraw + m.y3Visits.np + m.y3Visits.cv;
  const y3Summary = [
    m.y3Visits.mri ? `${m.y3Visits.mri} MRI` : null,
    m.y3Visits.bloodDraw ? `${m.y3Visits.bloodDraw} BD` : null,
    m.y3Visits.np ? `${m.y3Visits.np} NP` : null,
    m.y3Visits.cv ? `${m.y3Visits.cv} CV` : null,
  ].filter(Boolean).join(", ");

  addSlideTitle(slide, "3rd Year Study Visits");

  addBodyLines(
    slide,
    [
      { text: `Total Visits in ${monthName}: ${y3Total} (${y3Summary})`, bold: true },
      { text: `Participants with ${monthName} Due Dates (n=${a.dueCount})`, bold: true },
      { text: `${a.dueComplete} completed all visits`, bullet: true },
      {
        text:
          a.dueScheduled > 0
            ? `${a.dueScheduled} successfully scheduled`
            : `${Math.max(0, a.dueCount - a.dueComplete)} successfully scheduled`,
        bullet: true,
      },
      { text: `Scheduled Visits for ${nextName} (to Date):`, bold: true },
      `${a.nextVisitTotal} (${a.nextVisitSummary})`,
      { text: `Participants with ${nextName} Due Dates (n=${a.nextDueCount})`, bold: true },
      { text: `${a.nextDueCompletedEarly} completed in ${monthAbbr}`, bullet: true },
      { text: `${a.nextDueScheduled} successfully scheduled`, bullet: true },
    ],
    { x: DECK.marginL, y: 1.45, fontSize: 14, h: 3.8, w: 4.5 }
  );
  addSlideChartImage(slide, chartDataUrl);
}

function buildYear3OverviewSlide(pptx, m) {
  const slide = pptx.addSlide();
  const n = m.narrative || {};
  addSlideTitle(slide, "Year 3 Visit Status");

  addPlainText(slide, `As of ${asOfDateShort(m.year, m.month)}`, {
    y: contentTop(),
    fontSize: 18,
    bold: true,
    h: 0.4,
  });

  const body = [
    { text: `${m.y3.completed ?? "—"} participants`, bold: true, fontSize: 18 },
    "Have completed all required visits (BD, MRI, CV, NP)",
  ];
  if (n.y3CaughtUpThrough) {
    body.push({ text: `Visits are caught up through ${n.y3CaughtUpThrough}`, bold: true });
  }
  if (n.y3StudyPeriod) {
    body.push(`Year 3 study period: ${n.y3StudyPeriod}`);
  }

  addBodyLines(slide, body, { y: 1.75, fontSize: 18, h: 3.2 });
}

function buildYear3InProgressSlide(pptx, m) {
  const slide = pptx.addSlide();
  const groups = m.scheduling.y3InProgress;
  const total = groups.reduce((s, g) => s + g.count, 0);

  addSlideTitle(slide, "Year 3 Visit Status");
  addPlainText(slide, `${total} participants in progress`, {
    y: contentTop(),
    fontSize: 18,
    bold: true,
    h: 0.4,
  });

  addDataTable(
    slide,
    [
      [headerCell("Month Due"), headerCell("# Participants"), headerCell("Expected Completion")],
      ...groups.map((g) => [g.label, String(g.count), expectedCompletionText(g.due, m.year, m.month)]),
      [{ text: "Total", options: { bold: true } }, { text: String(total), options: { bold: true } }, ""],
    ],
    { y: 1.75, colW: [1.8, 1.5, 5.5], fontSize: 12, rowH: 0.38 }
  );
}

function buildYear3ToContactSlide(pptx, m) {
  const slide = pptx.addSlide();
  const groups = m.scheduling.y3ToContact;
  const total = groups.reduce((s, g) => s + g.count, 0);
  const n = m.narrative || {};

  addSlideTitle(slide, "Year 3 Visit Status");
  addPlainText(slide, `${total} participants to be contacted`, {
    y: contentTop(),
    fontSize: 18,
    bold: true,
    h: 0.35,
  });
  if (n.y3RemainingExpected != null && n.y3CompletionDeadline) {
    addPlainText(
      slide,
      `${n.y3RemainingExpected} Participants Expected to Complete Year 3 by ${n.y3CompletionDeadline}`,
      {
        y: contentTop() + 0.38,
        fontSize: 14,
        h: 0.35,
      }
    );
  }

  addDataTable(
    slide,
    [
      [headerCell("Month Due"), headerCell("# Participants")],
      ...groups.map((g) => [g.label, String(g.count)]),
      [{ text: "Total", options: { bold: true } }, { text: String(total), options: { bold: true } }],
    ],
    { y: 1.95, colW: [4.5, 4.0], fontSize: 14, rowH: 0.38 }
  );
}

function buildYear4ActivitySlide(pptx, m, chartDataUrl) {
  const slide = pptx.addSlide();
  const monthName = MONTH_NAMES[m.month - 1];
  const a = m.scheduling.y4Activity;
  const nextName = MONTH_NAMES[a.nextMonth.month - 1];
  const y4Total = m.y4Visits.bloodDraw + m.y4Visits.np + m.y4Visits.cv;
  const y4Summary = [
    m.y4Visits.bloodDraw ? `${m.y4Visits.bloodDraw} BD` : null,
    m.y4Visits.np ? `${m.y4Visits.np} NP` : null,
    m.y4Visits.cv ? `${m.y4Visits.cv} CV` : null,
  ].filter(Boolean).join(", ");

  addSlideTitle(slide, "4th Year Study Visits 🎉");

  addBodyLines(
    slide,
    [
      { text: `Total Visits in ${monthName}: ${y4Total} (${y4Summary})`, bold: true },
      { text: `Participants with ${monthName} Due Dates (n=${a.dueCount})`, bold: true },
      { text: `${a.dueComplete} completed all visits`, bullet: true },
      { text: `${Math.max(0, a.dueCount - a.dueComplete)} rescheduled for ${nextName.slice(0, 3).toLowerCase()}`, bullet: true },
      { text: `Scheduled Visits for ${nextName} (to Date): ${a.nextVisitTotal} (${a.nextVisitSummary})`, bold: true },
      { text: `Participants with ${nextName} Due Dates (n=${a.nextDueCount})`, bold: true },
      { text: `${a.nextDueCount} pending scheduling`, bullet: true },
    ],
    { y: 1.45, fontSize: 14, h: 3.8, w: 4.8 }
  );
  addSlideChartImage(slide, chartDataUrl);
}

function buildYear4ToContactSlide(pptx, m) {
  const slide = pptx.addSlide();
  const groups = m.scheduling.y4ToContact;
  const total = groups.reduce((s, g) => s + g.count, 0);
  const n = m.narrative || {};
  const rollover = m.scheduling.y4Rollovers || n.y4Rollovers;

  addSlideTitle(slide, "Year 4 Visit Status");
  const intro = [{ text: `${total} participants to be contacted through December ${m.year}`, bold: true }];
  if (rollover?.count) {
    intro.push(
      `${rollover.count} participants to be contacted ${rollover.rangeLabel} (3rd-year rollovers)`
    );
  }
  if (n.expectedStudyCompletion) {
    intro.push(`Expected Study Completion: ${n.expectedStudyCompletion}`);
  }
  addBodyLines(slide, intro, { y: contentTop(), fontSize: 14, h: 1.0 });

  addDataTable(
    slide,
    [
      [headerCell("Month Due"), headerCell("# Participants")],
      ...groups.map((g) => [g.label, String(g.count)]),
      [{ text: "Total", options: { bold: true } }, { text: String(total), options: { bold: true } }],
    ],
    { y: 2.15, colW: [4.5, 4.0], fontSize: 14, rowH: 0.38 }
  );
}

function buildStudyProgressSlide(pptx, m, chartDataUrl) {
  const slide = pptx.addSlide();
  const n = m.narrative || {};
  addSlideTitle(slide, "Study Progress", `Participant Flow (N = ${m.enrolled ?? "—"})`);

  const y2Lines = [
    { text: "Year 2", bold: true },
    { text: `${m.y2.completed ?? "—"} participants completed Year 2`, bullet: true },
  ];
  if (n.advancedToY3Text) {
    y2Lines.push({ text: n.advancedToY3Text, bullet: true });
  }
  y2Lines.push({ text: `${m.dropouts.year2} additional participants dropped out`, bullet: true });

  addBodyLines(
    slide,
    [
      { text: "Baseline", bold: true },
      { text: `${m.enrolled} enrolled`, bullet: true },
      { text: `${m.baseline.completed ?? "—"} completed Baseline assessments`, bullet: true },
      { text: `${m.dropouts.baseline} participants dropped out`, bullet: true },
      ...y2Lines,
      { text: "Year 3 (ongoing)", bold: true },
      { text: `${m.y3.completed ?? "—"} participants completed Year 3`, bullet: true },
      { text: `${m.dropouts.year3} additional participants dropped out`, bullet: true },
      { text: "Year 4 (ongoing)", bold: true },
      { text: `${m.y4.completed ?? "—"} participants completed Year 4`, bullet: true },
    ],
    { x: DECK.marginL, y: contentTop(true), fontSize: 14, h: 3.8, w: 4.4 }
  );
  addSlideChartImage(slide, chartDataUrl);
}

function buildQuestionsSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Questions / Comments");

  const items = m.narrative?.questions?.length
    ? m.narrative.questions.map((text) => ({ text, bullet: true }))
    : [{ text: "Open discussion", bullet: false }];

  addBodyLines(slide, items, { y: 1.55, fontSize: 18, h: 3.2 });
}

function buildThankYouSlide(pptx) {
  const slide = pptx.addSlide();
  slideBg(slide);
  addPlainText(slide, "THANK YOU :)", {
    x: 0.5,
    y: 2.2,
    w: 9.0,
    h: 1.2,
    fontSize: 54,
    bold: true,
    align: "center",
  });
}

async function generateMonthlyReportPPT(year, month) {
  if (typeof PptxGenJS === "undefined") {
    throw new Error("PptxGenJS library not loaded.");
  }
  const metrics = collectReportMetrics(year, month);
  const logoData = await getReportLogoData();
  const [y3Chart, y4Chart, progressChart] = await Promise.all([
    buildYear3VisitsChartImage(metrics.year, metrics.month),
    buildYear4VisitsChartImage(metrics.year, metrics.month),
    buildStudyProgressChartImage(metrics),
  ]);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "DBS Operations Dashboard";
  pptx.company = "USC INI";
  pptx.subject = `DBS Monthly Progress Report — ${metrics.label}`;
  pptx.title = `DBS Monthly Report ${metrics.label}`;

  buildTitleSlide(pptx, metrics, logoData);
  buildExecutiveSummarySlide(pptx, metrics);
  buildCompletedVisitsSlide(pptx, metrics);
  buildYear3ActivitySlide(pptx, metrics, y3Chart);
  buildYear3OverviewSlide(pptx, metrics);
  buildYear3InProgressSlide(pptx, metrics);
  buildYear3ToContactSlide(pptx, metrics);
  buildYear4ActivitySlide(pptx, metrics, y4Chart);
  buildYear4ToContactSlide(pptx, metrics);
  buildStudyProgressSlide(pptx, metrics, progressChart);
  buildQuestionsSlide(pptx, metrics);
  buildThankYouSlide(pptx);

  const filename = `DBS_Monthly_Report_${metrics.year}_${String(metrics.month).padStart(2, "0")}.pptx`;
  await pptx.writeFile({ fileName: filename });
  return { filename, metrics };
}

function monthFromColumnKey(key) {
  if (!key || ["visit_type", "total_completed"].includes(key)) return null;
  const m = key.match(
    /(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s.']+'?(\d{2})/i
  );
  if (!m) return null;
  const abbr = m[1].slice(0, 3).toLowerCase();
  const idx = MONTH_NAMES.findIndex((n) => n.toLowerCase().startsWith(abbr));
  if (idx < 0) return null;
  return { year: 2000 + Number(m[2]), month: idx + 1 };
}

function addMonthToSet(seen, year, month) {
  if (!year || !month || month < 1 || month > 12) return;
  seen.add(`${year}-${month}`);
}

function collectMonthsFromDates(seen, value) {
  const d = parseIsoDate(value);
  if (d) addMonthToSet(seen, d.getFullYear(), d.getMonth() + 1);
}

function availableReportMonths() {
  const seen = new Set();

  ["visit_trends_sheet5", "visit_trends_sheet6"].forEach((key) => {
    (DATA[key] || []).forEach((r) => collectMonthsFromDates(seen, r.Date));
  });

  (DATA.second_year_monthly || []).forEach((row) => {
    Object.keys(row).forEach((k) => {
      const parsed = monthFromColumnKey(k);
      if (parsed) addMonthToSet(seen, parsed.year, parsed.month);
    });
  });

  ["third_year_scheduling", "fourth_year_scheduling", "second_year_scheduling", "baseline_scheduling"].forEach((key) => {
    (DATA[key] || []).forEach((row) => {
      collectMonthsFromDates(seen, row.Month);
      Object.entries(row).forEach(([field, val]) => {
        if (/date|month/i.test(field)) collectMonthsFromDates(seen, val);
      });
    });
  });

  ["second_year", "third_year", "fourth_year"].forEach((key) => {
    (DATA.completed_visits?.[key] || []).forEach((row) => {
      collectMonthsFromDates(seen, row.scheduling_month);
    });
  });

  ["year3", "year4"].forEach((yearKey) => {
    Object.keys(DATA.study_details_visit_totals?.[yearKey]?.monthly || {}).forEach((key) => {
      const [year, month] = key.split("-").map(Number);
      addMonthToSet(seen, year, month);
    });
  });

  if (DATA.meta?.generated_at) {
    const d = new Date(DATA.meta.generated_at);
    if (!Number.isNaN(d.getTime())) addMonthToSet(seen, d.getFullYear(), d.getMonth() + 1);
  }

  const endKey = reportEndKey();
  addMonthToSet(seen, Math.floor((endKey - 1) / 12), ((endKey - 1) % 12) + 1);

  const parsed = [...seen].map((k) => {
    const [y, m] = k.split("-").map(Number);
    return monthKey(y, m);
  });
  if (parsed.length) {
    const min = Math.max(Math.min(...parsed), REPORT_START_KEY);
    const max = Math.max(endKey, min);
    for (let cursor = min; cursor <= max; cursor += 1) {
      const year = Math.floor((cursor - 1) / 12);
      const month = ((cursor - 1) % 12) + 1;
      addMonthToSet(seen, year, month);
    }
  }

  return [...seen]
    .map((k) => {
      const [year, month] = k.split("-").map(Number);
      return { year, month, label: monthLabel(year, month), key: monthKey(year, month) };
    })
    .filter((m) => m.key >= REPORT_START_KEY && m.key <= endKey)
    .sort((a, b) => b.key - a.key);
}

function renderReportPreview(metrics) {
  const el = document.getElementById("report-preview");
  if (!el) return;

  const s = metrics.scheduling;
  const monthName = MONTH_NAMES[metrics.month - 1];
  const y3ip = s.y3InProgress.reduce((n, g) => n + g.count, 0);
  const y3tc = s.y3ToContact.reduce((n, g) => n + g.count, 0);
  const y4tc = s.y4ToContact.reduce((n, g) => n + g.count, 0);

  const slideRows = [
    ["Slide 2 · Y3 completed", metrics.y3.completed],
    ["Slide 2 · Y4 completed", metrics.y4.completed],
    ["Slide 2 · Dropouts total", metrics.dropouts.total],
    ["Slide 2 · Y3 MRI / BD / NP / CV", `${s.y3Components["MRI Date"]} / ${s.y3Components["BD Date"]} / ${s.y3Components["NP Date"]} / ${s.y3Components["CV Date"]}`],
    ["Slide 2 · Future MRI / BD / NP / CV", `${s.futureVisits.mri} / ${s.futureVisits.bd} / ${s.futureVisits.np} / ${s.futureVisits.cv}`],
    [`Slide 3 · Y3 visits (${monthName})`, `${metrics.y3Visits.mri} / ${metrics.y3Visits.bloodDraw} / ${metrics.y3Visits.np} / ${metrics.y3Visits.cv}`],
    [`Slide 3 · Y4 visits (${monthName})`, `${metrics.y4Visits.bloodDraw} / ${metrics.y4Visits.np} / ${metrics.y4Visits.cv}`],
    ["Slide 4 · Y3 due / complete / scheduled", `${s.y3Activity.dueCount} / ${s.y3Activity.dueComplete} / ${s.y3Activity.dueScheduled}`],
    ["Slide 4 · Next month Y3 scheduled", `${s.y3Activity.nextVisitTotal} (${s.y3Activity.nextVisitSummary})`],
    ["Slide 6 · Y3 in progress", y3ip],
    ["Slide 7 · Y3 to be contacted", y3tc],
    ["Slide 8 · Y4 due / complete", `${s.y4Activity.dueCount} / ${s.y4Activity.dueComplete}`],
    ["Slide 9 · Y4 to be contacted", y4tc],
    ["Slide 10 · Enrolled / Baseline / Y2 / Y3 / Y4", `${metrics.enrolled} / ${metrics.baseline.completed} / ${metrics.y2.completed} / ${metrics.y3.completed} / ${metrics.y4.completed}`],
  ];

  const validationRows = validateReportMetrics(metrics).map((row) => [
    row.label,
    row.status === "pass" ? `✓ ${row.actual}` : `✗ ${row.actual ?? "missing"}`,
  ]);

  const renderSection = (title, note, rows) => `
    <div class="report-preview-section">
      <h4>${title}</h4>
      ${note ? `<p class="report-preview-note">${note}</p>` : ""}
      <table>
        <thead><tr><th>Metric</th><th>Value</th></tr></thead>
        <tbody>${rows.map(([k, v]) => `<tr><td>${k}</td><td><strong>${v ?? 0}</strong></td></tr>`).join("")}</tbody>
      </table>
    </div>`;

  el.innerHTML =
    renderSection(
      "PowerPoint slide numbers",
      "Derived from DBS Tracker.xlsx + Monthly Meeting Updates.xlsx (live sheet data).",
      slideRows
    ) +
    renderSection(
      "Data validation",
      "✓ = populated from sheets. All numbers below come directly from your tracker.",
      validationRows
    );
}

/** Verify report metrics are populated from sheet data. */
function validateReportMetrics(metrics) {
  const s = metrics.scheduling;
  const checks = [];

  const add = (label, actual, minExpected = 0) => {
    if (actual == null || actual === "") {
      checks.push({ label, actual: null, expected: minExpected, status: "fail" });
      return;
    }
    checks.push({ label, actual, expected: minExpected, status: "pass" });
  };

  add("Enrolled", metrics.enrolled);
  add("Baseline completed", metrics.baseline.completed);
  add("Year 2 completed", metrics.y2.completed);
  add("Year 3 completed", metrics.y3.completed);
  add("Year 4 completed", metrics.y4.completed);
  add("Total dropouts", metrics.dropouts.total);
  add("Active participants", metrics.active);
  add("Y3 components MRI (Study Details)", s.y3Components["MRI Date"]);
  add("Y3 components BD (Study Details)", s.y3Components["BD Date"]);
  add("Y3 components NP (Study Details)", s.y3Components["NP Date"]);
  add("Y3 components CV (Study Details)", s.y3Components["CV Date"]);
  add("Y4 components BD/NP/CV (Study Details)", `${s.y4Components["BD Date"]} / ${s.y4Components["NP Date"]} / ${s.y4Components["CV Date"]}`);
  add("Y3 visits MRI", metrics.y3Visits.mri);
  add("Y3 visits BD", metrics.y3Visits.bloodDraw);
  add("Y4 visits BD", metrics.y4Visits.bloodDraw);
  add("Y3 in progress total", s.y3InProgress.reduce((n, g) => n + g.count, 0));
  add("Y3 to contact total", s.y3ToContact.reduce((n, g) => n + g.count, 0));
  add("Y4 to contact total", s.y4ToContact.reduce((n, g) => n + g.count, 0));

  return checks;
}

function updateReportPreview() {
  const select = document.getElementById("report-month-select");
  if (!select || !select.value || !DATA) return;
  const [year, month] = select.value.split("-").map(Number);
  const metrics = collectReportMetrics(year, month);
  renderReportPreview(metrics);
  const badge = document.getElementById("report-live-badge");
  if (badge) {
    badge.textContent = `Live preview · ${metrics.label} · Y3: ${metrics.y3Visits.mri}/${metrics.y3Visits.bloodDraw}/${metrics.y3Visits.np}/${metrics.y3Visits.cv} · Y4: ${metrics.y4Visits.bloodDraw}/${metrics.y4Visits.np}/${metrics.y4Visits.cv}`;
  }
}

function renderReport() {
  const select = document.getElementById("report-month-select");
  const months = availableReportMonths();
  if (!select) return;

  const prev = select.value;
  select.innerHTML = months
    .map((m) => `<option value="${m.year}-${m.month}">${m.label}</option>`)
    .join("");

  if (prev && [...select.options].some((o) => o.value === prev)) {
    select.value = prev;
  } else {
    const nowKey = monthKey(new Date().getFullYear(), new Date().getMonth() + 1);
    const pick = months.find((m) => m.key <= nowKey) || months[0];
    if (pick) select.value = `${pick.year}-${pick.month}`;
  }

  select.removeEventListener("change", updateReportPreview);
  select.addEventListener("change", updateReportPreview);
  updateReportPreview();

  const btn = document.getElementById("generate-report-btn");
  const status = document.getElementById("report-status");
  if (btn) {
    btn.onclick = async () => {
      const [year, month] = select.value.split("-").map(Number);
      btn.disabled = true;
      status.textContent = "Building PowerPoint…";
      status.className = "report-status loading";
      try {
        const { filename } = await generateMonthlyReportPPT(year, month);
        status.textContent = `Downloaded ${filename}`;
        status.className = "report-status success";
      } catch (err) {
        status.textContent = `Error: ${err.message}`;
        status.className = "report-status error";
      } finally {
        btn.disabled = false;
      }
    };
  }
}
