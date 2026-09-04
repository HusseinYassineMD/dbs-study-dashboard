/** Monthly progress report → USC-branded PowerPoint matching the monthly meeting deck. */

const USC = {
  cardinal: "990000",
  gold: "FFCC00",
  white: "FFFFFF",
  black: "000000",
  gray: "555555",
  lightGray: "F5F5F5",
};

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

const MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/** Report range: May 2024 through September 2026 */
const REPORT_START_KEY = 2024 * 12 + 5;
const REPORT_END_KEY = 2026 * 12 + 9;

function monthKey(year, month) {
  return year * 12 + month;
}

function displayCount(value) {
  if (value == null || value === "") return 0;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function parseIsoDate(value) {
  if (value == null || value === "") return null;
  const s = String(value).trim();
  if (!s || /opt out|n\/a|^nan$|ineligible/i.test(s)) return null;
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
      (sum, [y, m]) => sum + rows.filter((r) => inCalendarMonth(r[field], y, m)).length,
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

function isDroppedRow(row) {
  return /\bDROPPED\b/i.test(rowNotes(row));
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

function parseDueMonth(value, refYear, refMonth) {
  const d = parseIsoDate(value);
  if (d) return { year: d.getFullYear(), month: d.getMonth() + 1 };
  if (value == null || value === "") return null;
  const s = String(value).trim();
  for (let i = 0; i < MONTH_NAMES.length; i += 1) {
    const full = MONTH_NAMES[i];
    if (s.toLowerCase() === full.toLowerCase() || s.toLowerCase() === full.slice(0, 3).toLowerCase()) {
      let year = refYear;
      if (refMonth != null && i + 1 < refMonth - 1) year = refYear + 1;
      return { year, month: i + 1 };
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
    counts[field] = slice.filter((r) => parseIsoDate(r[field])).length;
  });
  counts.mriExemptions = slice.filter((r) => !parseIsoDate(r["MRI Date"]) && hasMriExemption(r)).length;
  return counts;
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

function groupIncompleteByDue(rows, completeFn, refYear, refMonth, { inProgressMaxKey, contactMinKey, completedCap }) {
  const completedIds = new Set();
  if (completedCap) {
    rows
      .filter(completeFn)
      .sort((a, b) => completionSortKey(a) - completionSortKey(b))
      .slice(0, completedCap)
      .forEach((row) => completedIds.add(row.participant_id));
  }

  const inProgress = new Map();
  const toContact = new Map();

  rows.forEach((row) => {
    if (isDroppedRow(row)) return;
    if (completedCap ? completedIds.has(row.participant_id) : completeFn(row)) return;
    const due = parseDueMonth(row.Month, refYear, refMonth);
    const key = dueMonthKey(due);
    if (key == null) return;
    const label = dueMonthLabel(due);
    const target = key <= inProgressMaxKey ? inProgress : key >= contactMinKey ? toContact : toContact;
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

function analyzeYearActivity(rows, completeFn, visitFields, reportYear, reportMonth) {
  const monthVisits = countFieldsInMonth(rows, visitFields, reportYear, reportMonth);
  const visitTotal = visitFields.reduce((s, f) => s + (monthVisits[f] || 0), 0);
  const visitSummary = visitFields
    .map((f) => {
      const label = f.replace(" Date", "");
      return `${monthVisits[f] || 0} ${label === "BD" ? "BD" : label === "NP" ? "NP" : label === "CV" ? "CV" : "MRI"}`;
    })
    .join(", ");

  const dueRows = participantsDueInMonth(rows, reportYear, reportMonth);
  const dueComplete = dueRows.filter(completeFn).length;
  const dueScheduled = dueRows.filter(
    (r) => !completeFn(r) && hasScheduledInOrAfter(r, visitFields, reportYear, reportMonth + 1)
  ).length;

  const next = addMonths(reportYear, reportMonth, 1);
  const nextVisits = countFieldsInMonth(rows, visitFields, next.year, next.month);
  const nextTotal = visitFields.reduce((s, f) => s + (nextVisits[f] || 0), 0);
  const nextSummary = visitFields
    .map((f) => `${nextVisits[f] || 0} ${f.replace(" Date", "")}`)
    .join(", ");

  const nextDueRows = participantsDueInMonth(rows, next.year, next.month);

  return {
    visitTotal,
    visitSummary,
    dueCount: dueRows.length,
    dueComplete,
    dueScheduled,
    nextMonth: next,
    nextVisitTotal: nextTotal,
    nextVisitSummary: nextSummary,
    nextDueCount: nextDueRows.length,
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

  const y3Groups = groupIncompleteByDue(third, isYear3Complete, reportYear, reportMonth, {
    inProgressMaxKey,
    contactMinKey,
    completedCap: y3Completed,
  });
  const y3InProgress = y3Groups.inProgress.filter(
    (g) => dueMonthKey(g.due) >= inProgressMinKey && dueMonthKey(g.due) <= inProgressMaxKey
  );

  const y4ContactStart = monthKey(reportYear, reportMonth + 1);
  const y4ContactEnd = monthKey(reportYear, 12);
  const y4CompletedIds = new Set(
    fourth
      .filter(isYear4Complete)
      .sort((a, b) => completionSortKey(a) - completionSortKey(b))
      .slice(0, y4Completed || 0)
      .map((r) => r.participant_id)
  );
  const y4ToContactMap = new Map();
  fourth.forEach((row) => {
    if (isDroppedRow(row) || y4CompletedIds.has(row.participant_id)) return;
    const due = parseDueMonth(row.Month, reportYear, reportMonth);
    const key = dueMonthKey(due);
    if (key == null || key < y4ContactStart || key > y4ContactEnd) return;
    const label = dueMonthLabel(due);
    if (!y4ToContactMap.has(label)) y4ToContactMap.set(label, { label, due, count: 0 });
    y4ToContactMap.get(label).count += 1;
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
  const y3Components = completedVisitComponents(third, isYear3Complete, y3Fields, y3Completed);
  const y4Components = completedVisitComponents(fourth, isYear4Complete, y4Fields, y4Completed);

  return {
    y3Activity: analyzeYearActivity(third, isYear3Complete, y3Fields, reportYear, reportMonth),
    y4Activity: analyzeYearActivity(fourth, isYear4Complete, y4Fields, reportYear, reportMonth),
    y3InProgress,
    y3ToContact: y3Groups.toContact,
    y4ToContact,
    futureVisits: {
      mri: y3Future["MRI Date"] || 0,
      bd: (y3Future["BD Date"] || 0) + (y4Future["BD Date"] || 0),
      np: (y3Future["NP Date"] || 0) + (y4Future["NP Date"] || 0),
      cv: (y3Future["CV Date"] || 0) + (y4Future["CV Date"] || 0),
    },
    y3Components,
    y4Components,
    asOfDate: `${MONTH_NAMES[reportMonth - 1]} ${reportYear}`,
  };
}

function collectReportMetrics(year, month) {
  const baseline = progressRow("baseline");
  const y2 = progressRow("2");
  const y3 = progressRow("3");
  const y4 = progressRow("4");

  const thirdSched = DATA.third_year_scheduling || [];
  const fourthSched = DATA.fourth_year_scheduling || [];

  const y3Visits = countFieldsInMonth(thirdSched, ["MRI Date", "BD Date", "NP Date", "CV Date"], year, month);
  const y4Visits = countFieldsInMonth(fourthSched, ["BD Date", "NP Date", "CV Date"], year, month);

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

  const scheduling = collectSchedulingReportData(year, month, y3.completed ?? 0, y4.completed ?? 0);

  return {
    label: monthLabel(year, month),
    year,
    month,
    generatedAt: DATA.meta?.generated_at,
    baseline,
    y2,
    y3,
    y4,
    active: DATA.active_participants?.length ?? 0,
    enrolled: baseline.enrolled ?? 200,
    dropouts,
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
}

function addCardinalBar(slide) {
  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: "100%",
    h: 0.12,
    fill: { color: USC.cardinal },
    line: { color: USC.cardinal, width: 0 },
  });
}

function addGoldAccent(slide) {
  slide.addShape("rect", {
    x: 0,
    y: 0.12,
    w: "100%",
    h: 0.04,
    fill: { color: USC.gold },
    line: { color: USC.gold, width: 0 },
  });
}

function addSlideTitle(slide, title, subtitle) {
  addCardinalBar(slide);
  addGoldAccent(slide);
  slide.addText(title, {
    x: 0.55,
    y: 0.45,
    w: 8.9,
    h: 0.7,
    fontSize: 28,
    bold: true,
    color: USC.cardinal,
    fontFace: "Arial",
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.55,
      y: 1.05,
      w: 8.9,
      h: 0.35,
      fontSize: 14,
      color: USC.gray,
      fontFace: "Arial",
    });
  }
}

function addBullets(slide, items, opts = {}) {
  const text = items.map((item) => ({ text: item, options: { bullet: true, breakLine: true } }));
  slide.addText(text, {
    x: opts.x ?? 0.65,
    y: opts.y ?? 1.55,
    w: opts.w ?? 8.7,
    h: opts.h ?? 3.8,
    fontSize: opts.fontSize ?? 16,
    color: USC.black,
    fontFace: "Arial",
    valign: "top",
  });
}

function addDataTable(slide, rows, opts = {}) {
  slide.addTable(rows, {
    x: opts.x ?? 0.55,
    y: opts.y ?? 1.55,
    w: opts.w ?? 8.9,
    colW: opts.colW,
    fontSize: opts.fontSize ?? 13,
    border: { type: "solid", color: USC.cardinal, pt: 0.75 },
    fontFace: "Arial",
  });
}

function headerCell(text) {
  return { text, options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } };
}

function buildTitleSlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: USC.cardinal };
  slide.addShape("rect", {
    x: 0,
    y: 4.85,
    w: "100%",
    h: 0.775,
    fill: { color: USC.gold },
    line: { color: USC.gold, width: 0 },
  });
  slide.addText("Diabetes Brain Study –", {
    x: 0.7,
    y: 1.45,
    w: 8.6,
    h: 0.9,
    fontSize: 38,
    bold: true,
    color: USC.white,
    fontFace: "Arial",
  });
  slide.addText("Monthly Progress Meeting", {
    x: 0.7,
    y: 2.25,
    w: 8.6,
    h: 0.55,
    fontSize: 24,
    color: USC.gold,
    fontFace: "Arial",
  });
  slide.addText(m.label, {
    x: 0.7,
    y: 3.05,
    w: 8.6,
    h: 0.5,
    fontSize: 22,
    color: USC.white,
    fontFace: "Arial",
  });
}

function buildExecutiveSummarySlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Executive Summary");

  const s = m.scheduling;
  const next = addMonths(m.year, m.month, 1);
  const next2 = addMonths(m.year, m.month, 2);
  const futureLabel = `${MONTH_NAMES[next.month - 1].slice(0, 3)} – ${MONTH_NAMES[next2.month - 1].slice(0, 3)}`;

  slide.addText(`${m.dropouts.total} Dropout Participants`, {
    x: 0.55,
    y: 1.45,
    w: 2.8,
    h: 0.35,
    fontSize: 16,
    bold: true,
    color: USC.cardinal,
    fontFace: "Arial",
  });
  addBullets(
    slide,
    [
      `${m.dropouts.baseline} Baseline year`,
      `${m.dropouts.year2} Second year`,
      `${m.dropouts.year3} Third year`,
      `${m.dropouts.year4} Fourth year`,
    ],
    { x: 0.65, y: 1.85, w: 2.6, h: 1.5, fontSize: 13 }
  );

  slide.addText("Future Visits Scheduled", {
    x: 3.55,
    y: 1.45,
    w: 2.8,
    h: 0.35,
    fontSize: 16,
    bold: true,
    color: USC.cardinal,
    fontFace: "Arial",
  });
  slide.addText(`${futureLabel} (3rd and 4th year)`, {
    x: 3.55,
    y: 1.78,
    w: 2.8,
    h: 0.3,
    fontSize: 11,
    color: USC.gray,
    fontFace: "Arial",
  });
  addBullets(
    slide,
    [
      `${s.futureVisits.mri} MRI's`,
      `${s.futureVisits.bd} Full Blood Draws`,
      `${s.futureVisits.np} Neuropsych Tests`,
      `${s.futureVisits.cv} Clinician Visits`,
    ],
    { x: 3.65, y: 2.1, w: 2.6, h: 1.5, fontSize: 13 }
  );

  slide.addText(`${m.y3.completed ?? "—"} Participants`, {
    x: 0.55,
    y: 3.45,
    w: 4.2,
    h: 0.35,
    fontSize: 16,
    bold: true,
    color: USC.cardinal,
    fontFace: "Arial",
  });
  slide.addText("Completed 3rd Year", {
    x: 0.55,
    y: 3.78,
    w: 4.2,
    h: 0.3,
    fontSize: 12,
    color: USC.gray,
    fontFace: "Arial",
  });
  addBullets(
    slide,
    [
      `${s.y3Components["MRI Date"] ?? "—"} MRI *`,
      `${s.y3Components["BD Date"] ?? "—"} Blood Draw`,
      `${s.y3Components["NP Date"] ?? "—"} Neuropsych Tests`,
      `${s.y3Components["CV Date"] ?? "—"} Clinician Visits`,
    ],
    { x: 0.65, y: 4.1, w: 4, h: 1.2, fontSize: 13 }
  );

  slide.addText(`${m.y4.completed ?? "—"} Participants`, {
    x: 5.1,
    y: 3.45,
    w: 4.2,
    h: 0.35,
    fontSize: 16,
    bold: true,
    color: USC.cardinal,
    fontFace: "Arial",
  });
  slide.addText("Completed 4th Year", {
    x: 5.1,
    y: 3.78,
    w: 4.2,
    h: 0.3,
    fontSize: 12,
    color: USC.gray,
    fontFace: "Arial",
  });
  addBullets(
    slide,
    [
      `${s.y4Components["BD Date"] ?? "—"} Blood Draw`,
      `${s.y4Components["NP Date"] ?? "—"} Neuropsych Tests`,
      `${s.y4Components["CV Date"] ?? "—"} Clinician Visits`,
    ],
    { x: 5.2, y: 4.1, w: 4, h: 1.2, fontSize: 13 }
  );

  if (s.y3Components.mriExemptions) {
    slide.addText("* Some participants ineligible or opted out of MRI", {
      x: 0.55,
      y: 5.15,
      w: 8.5,
      h: 0.25,
      fontSize: 10,
      italic: true,
      color: USC.gray,
      fontFace: "Arial",
    });
  }
}

function buildCompletedVisitsSlide(pptx, m) {
  const slide = pptx.addSlide();
  const monthName = MONTH_NAMES[m.month - 1];
  addSlideTitle(slide, `All Completed Visits in ${monthName}`);

  const dash = "–";
  const rows = [
    [headerCell("Visit Type"), headerCell("3rd Year"), headerCell("4th Year")],
    ["MRI", String(m.y3Visits.mri || dash), dash],
    ["Blood Draw", String(m.y3Visits.bloodDraw || dash), String(m.y4Visits.bloodDraw || dash)],
    ["NP Visit", String(m.y3Visits.np || dash), String(m.y4Visits.np || dash)],
    ["Clinician Visit", String(m.y3Visits.cv || dash), String(m.y4Visits.cv || dash)],
  ];

  addDataTable(slide, rows, { y: 1.65, colW: [2.8, 2.8, 2.8] });

  const total = m.y3TotalMonthVisits;
  const parts = [
    m.y3Visits.mri ? `${m.y3Visits.mri} MRI` : null,
    m.y3Visits.bloodDraw ? `${m.y3Visits.bloodDraw} BD` : null,
    m.y3Visits.np ? `${m.y3Visits.np} NP` : null,
    m.y3Visits.cv ? `${m.y3Visits.cv} CV` : null,
  ].filter(Boolean);
  slide.addText(`Total Visits in ${monthName}: ${total} (${parts.join(", ")})`, {
    x: 0.55,
    y: 4.35,
    w: 8.9,
    h: 0.35,
    fontSize: 13,
    color: USC.black,
    fontFace: "Arial",
  });
}

function buildYear3ActivitySlide(pptx, m) {
  const slide = pptx.addSlide();
  const monthName = MONTH_NAMES[m.month - 1];
  const a = m.scheduling.y3Activity;
  const nextName = MONTH_NAMES[a.nextMonth.month - 1];
  addSlideTitle(slide, "3rd Year Study Visits");

  addBullets(slide, [
    `Total Visits in ${monthName}: ${a.visitTotal} (${a.visitSummary})`,
    `Participants with ${monthName} Due Dates (n=${a.dueCount})`,
    `● ${a.dueComplete} completed all visits`,
    `● ${Math.max(0, a.dueCount - a.dueComplete - a.dueScheduled)} successfully scheduled (${a.dueScheduled} in ${nextName}+)`,
    `Scheduled Visits for ${nextName} (to Date): ${a.nextVisitTotal} (${a.nextVisitSummary})`,
    `Participants with ${nextName} Due Dates (n=${a.nextDueCount})`,
  ], { y: 1.45, fontSize: 15 });
}

function buildYear3OverviewSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Year 3 Visit Status", `As of ${m.scheduling.asOfDate}`);

  addBullets(slide, [
    `${m.y3.completed ?? "—"} participants have completed all required visits (BD, MRI, CV, NP)`,
    "Visits are caught up through May 2026",
    "Year 3 study period: May 2025 – Jan/Feb 2027",
    `${m.y3.remaining ?? m.active ?? "—"} participants expected to complete Year 3 by Jan/Feb 2027`,
  ], { y: 1.55, fontSize: 17 });
}

function buildYear3InProgressSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Year 3 Visit Status");

  const groups = m.scheduling.y3InProgress;
  const total = groups.reduce((s, g) => s + g.count, 0);

  slide.addText(`${total} participants in progress`, {
    x: 0.55,
    y: 1.45,
    w: 8.9,
    h: 0.4,
    fontSize: 18,
    bold: true,
    color: USC.cardinal,
    fontFace: "Arial",
  });

  const rows = [
    [headerCell("Month Due"), headerCell("# Participants"), headerCell("Expected Completion")],
    ...groups.map((g) => [
      g.label,
      String(g.count),
      expectedCompletionText(g.due, m.year, m.month),
    ]),
    [{ text: "Total", options: { bold: true } }, { text: String(total), options: { bold: true } }, ""],
  ];

  addDataTable(slide, rows, { y: 2.0, colW: [2.4, 2.2, 4.0], fontSize: 12 });
}

function buildYear3ToContactSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Year 3 Visit Status");

  const groups = m.scheduling.y3ToContact;
  const total = groups.reduce((s, g) => s + g.count, 0);

  slide.addText(`${total} participants to be contacted`, {
    x: 0.55,
    y: 1.35,
    w: 8.9,
    h: 0.35,
    fontSize: 18,
    bold: true,
    color: USC.cardinal,
    fontFace: "Arial",
  });
  slide.addText(`${m.y3.remaining ?? m.active ?? "—"} Participants Expected to Complete Year 3 by Jan/Feb 2027`, {
    x: 0.55,
    y: 1.72,
    w: 8.9,
    h: 0.3,
    fontSize: 12,
    color: USC.gray,
    fontFace: "Arial",
  });

  const rows = [
    [headerCell("Month Due"), headerCell("# Participants")],
    ...groups.map((g) => [g.label, String(g.count)]),
    [{ text: "Total", options: { bold: true } }, { text: String(total), options: { bold: true } }],
  ];

  addDataTable(slide, rows, { y: 2.15, colW: [4.0, 4.0], fontSize: 13 });
}

function buildYear4ActivitySlide(pptx, m) {
  const slide = pptx.addSlide();
  const monthName = MONTH_NAMES[m.month - 1];
  const a = m.scheduling.y4Activity;
  const nextName = MONTH_NAMES[a.nextMonth.month - 1];
  addSlideTitle(slide, "4th Year Study Visits 🎉");

  addBullets(slide, [
    `Total Visits in ${monthName}: ${a.visitTotal} (${a.visitSummary})`,
    `Participants with ${monthName} Due Dates (n=${a.dueCount})`,
    `● ${a.dueComplete} completed all visits`,
    `● ${Math.max(0, a.dueCount - a.dueComplete)} rescheduled or pending`,
    `Scheduled Visits for ${nextName} (to Date): ${a.nextVisitTotal} (${a.nextVisitSummary})`,
    `Participants with ${nextName} Due Dates (n=${a.nextDueCount})`,
    `● ${a.nextDueCount} pending scheduling`,
  ], { y: 1.45, fontSize: 15 });
}

function buildYear4ToContactSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Year 4 Visit Status");

  const groups = m.scheduling.y4ToContact;
  const total = groups.reduce((s, g) => s + g.count, 0);
  const yearEnd = m.year;

  slide.addText(`${total} participants to be contacted through December ${yearEnd}`, {
    x: 0.55,
    y: 1.35,
    w: 8.9,
    h: 0.35,
    fontSize: 16,
    bold: true,
    color: USC.cardinal,
    fontFace: "Arial",
  });
  slide.addText("65 participants to be contacted January–August 2027 (3rd-year rollovers)", {
    x: 0.55,
    y: 1.72,
    w: 8.9,
    h: 0.3,
    fontSize: 12,
    color: USC.gray,
    fontFace: "Arial",
  });
  slide.addText("Expected Study Completion: Dec 2027 / Jan 2028", {
    x: 0.55,
    y: 2.0,
    w: 8.9,
    h: 0.3,
    fontSize: 12,
    color: USC.gray,
    fontFace: "Arial",
  });

  const rows = [
    [headerCell("Month Due"), headerCell("# Participants")],
    ...groups.map((g) => [g.label, String(g.count)]),
    [{ text: "Total", options: { bold: true } }, { text: String(total), options: { bold: true } }],
  ];

  addDataTable(slide, rows, { y: 2.45, colW: [4.0, 4.0], fontSize: 13 });
}

function buildStudyProgressSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Study Progress", `Participant Flow (N = ${m.enrolled})`);

  addBullets(slide, [
    `Baseline`,
    `● ${m.enrolled} enrolled`,
    `● ${m.baseline.completed ?? "—"} completed Baseline assessments`,
    `● ${m.dropouts.baseline} participants dropped out`,
    `Year 2`,
    `● ${m.y2.completed ?? "—"} participants completed Year 2`,
    `● +1 participant advanced to Year 3`,
    `● ${m.dropouts.year2} additional participants dropped out`,
    `Year 3 (ongoing)`,
    `● ${m.y3.completed ?? "—"} participants completed Year 3`,
    `● ${m.dropouts.year3} additional participants dropped out`,
    `Year 4 (ongoing)`,
    `● ${m.y4.completed ?? "—"} participants completed Year 4`,
  ], { y: 1.45, fontSize: 15 });
}

function buildQuestionsSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Questions / Comments");

  addBullets(slide, [
    "Funding account end date: review and update CIA scheduler for MRI bookings beyond current period",
    "Review scheduling notes in DBS Tracker for any participant-specific updates before presenting",
    `Report generated from dashboard data · ${m.label}`,
  ], { y: 1.55, fontSize: 16 });
}

function buildThankYouSlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: USC.cardinal };
  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: "100%",
    h: 0.08,
    fill: { color: USC.gold },
    line: { color: USC.gold, width: 0 },
  });
  slide.addText("THANK YOU :)", {
    x: 0.7,
    y: 2.1,
    w: 8.6,
    h: 0.9,
    fontSize: 44,
    bold: true,
    color: USC.white,
    align: "center",
    fontFace: "Arial",
  });
  slide.addText(`Diabetes Brain Study · ${m.label}`, {
    x: 0.7,
    y: 4.85,
    w: 8.6,
    h: 0.35,
    fontSize: 12,
    color: USC.white,
    align: "center",
    fontFace: "Arial",
  });
}

async function generateMonthlyReportPPT(year, month) {
  if (typeof PptxGenJS === "undefined") {
    throw new Error("PptxGenJS library not loaded.");
  }
  const metrics = collectReportMetrics(year, month);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "DBS Operations Dashboard";
  pptx.company = "USC";
  pptx.subject = `DBS Monthly Progress Report — ${metrics.label}`;
  pptx.title = `DBS Monthly Report ${metrics.label}`;

  buildTitleSlide(pptx, metrics);
  buildExecutiveSummarySlide(pptx, metrics);
  buildCompletedVisitsSlide(pptx, metrics);
  buildYear3ActivitySlide(pptx, metrics);
  buildYear3OverviewSlide(pptx, metrics);
  buildYear3InProgressSlide(pptx, metrics);
  buildYear3ToContactSlide(pptx, metrics);
  buildYear4ActivitySlide(pptx, metrics);
  buildYear4ToContactSlide(pptx, metrics);
  buildStudyProgressSlide(pptx, metrics);
  buildQuestionsSlide(pptx, metrics);
  buildThankYouSlide(pptx, metrics);

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

  if (DATA.meta?.generated_at) {
    const d = new Date(DATA.meta.generated_at);
    if (!Number.isNaN(d.getTime())) addMonthToSet(seen, d.getFullYear(), d.getMonth() + 1);
  }

  const parsed = [...seen].map((k) => {
    const [y, m] = k.split("-").map(Number);
    return monthKey(y, m);
  });
  if (parsed.length) {
    const min = Math.max(Math.min(...parsed), REPORT_START_KEY);
    const max = Math.min(Math.max(...parsed), REPORT_END_KEY);
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
    .filter((m) => m.key >= REPORT_START_KEY && m.key <= REPORT_END_KEY)
    .sort((a, b) => b.key - a.key);
}

function renderReportPreview(metrics) {
  const el = document.getElementById("report-preview");
  if (!el) return;

  const s = metrics.scheduling;
  const monthName = MONTH_NAMES[metrics.month - 1];

  const slideSections = [
    {
      title: "Slide 2 — Executive Summary",
      rows: [
        ["Total Dropouts", metrics.dropouts.total],
        ["Future MRI (next 2 mo)", s.futureVisits.mri],
        ["Future BD (next 2 mo)", s.futureVisits.bd],
        ["Future NP (next 2 mo)", s.futureVisits.np],
        ["Future CV (next 2 mo)", s.futureVisits.cv],
        ["Year 3 Completed", metrics.y3.completed],
        ["Y3 MRI / BD / NP / CV", `${s.y3Components["MRI Date"]} / ${s.y3Components["BD Date"]} / ${s.y3Components["NP Date"]} / ${s.y3Components["CV Date"]}`],
        ["Year 4 Completed", metrics.y4.completed],
      ],
    },
    {
      title: `Slide 3 — All Completed Visits in ${monthName}`,
      rows: [
        ["3rd Year MRI", metrics.y3Visits.mri],
        ["3rd Year BD / NP / CV", `${metrics.y3Visits.bloodDraw} / ${metrics.y3Visits.np} / ${metrics.y3Visits.cv}`],
        ["4th Year BD / NP / CV", `${metrics.y4Visits.bloodDraw} / ${metrics.y4Visits.np} / ${metrics.y4Visits.cv}`],
      ],
    },
    {
      title: "Slide 4 — 3rd Year Study Visits",
      rows: [
        [`${monthName} visit total`, s.y3Activity.visitTotal],
        [`${monthName} due participants`, s.y3Activity.dueCount],
        [`Next month scheduled`, s.y3Activity.nextVisitTotal],
      ],
    },
    {
      title: "Slides 6–7 — Year 3 Scheduling",
      rows: [
        ["In progress", s.y3InProgress.reduce((n, g) => n + g.count, 0)],
        ["To be contacted", s.y3ToContact.reduce((n, g) => n + g.count, 0)],
        ...s.y3InProgress.map((g) => [`  Due ${g.label}`, g.count]),
        ...s.y3ToContact.map((g) => [`  Contact ${g.label}`, g.count]),
      ],
    },
    {
      title: "Slide 10 — Study Progress",
      rows: [
        ["Enrolled", metrics.enrolled],
        ["Baseline completed", metrics.baseline.completed],
        ["Year 2 completed", metrics.y2.completed],
        ["Year 3 completed", metrics.y3.completed],
        ["Year 4 completed", metrics.y4.completed],
        ["Active participants", metrics.active],
      ],
    },
  ];

  el.innerHTML = slideSections
    .map(
      (section) => `
    <div class="report-preview-section">
      <h4>${section.title}</h4>
      <table>
        <thead><tr><th>Metric</th><th>Value</th></tr></thead>
        <tbody>${section.rows.map(([k, v]) => `<tr><td>${k}</td><td><strong>${v ?? 0}</strong></td></tr>`).join("")}</tbody>
      </table>
    </div>`
    )
    .join("");
}

function updateReportPreview() {
  const select = document.getElementById("report-month-select");
  if (!select || !select.value || !DATA) return;
  const [year, month] = select.value.split("-").map(Number);
  const metrics = collectReportMetrics(year, month);
  renderReportPreview(metrics);
  const badge = document.getElementById("report-live-badge");
  if (badge) {
    badge.textContent = `Live preview · ${metrics.label} · matches August 2026 deck format · Y3: ${metrics.y3Visits.mri}/${metrics.y3Visits.bloodDraw}/${metrics.y3Visits.np}/${metrics.y3Visits.cv} · Y4: ${metrics.y4Visits.bloodDraw}/${metrics.y4Visits.np}/${metrics.y4Visits.cv}`;
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
