/** Monthly progress report → PowerPoint matching the monthly meeting deck (August 2026 format). */

const DECK = {
  bg: "FFFFFF",
  title: "1B2430",
  body: "1B2430",
  muted: "5F6B7A",
  tableBorder: "B8C0CC",
  font: "Calibri",
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
  const nextDueCompletedEarly = nextDueRows.filter((row) => {
    if (completeFn(row)) return true;
    return ["MRI Date", "BD Date", "NP Date", "CV Date"].some((f) =>
      inCalendarMonth(row[f], reportYear, reportMonth)
    );
  }).length;

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
    nextDueCompletedEarly,
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

function asOfDateShort(year, month) {
  const lastDay = new Date(year, month, 0).getDate();
  return `${month}/${lastDay}/${year}`;
}

function mriFootnoteFromData() {
  const ids = [];
  (DATA.third_year_scheduling || []).forEach((row) => {
    const pid = row.participant_id || row.ID;
    if (!pid) return;
    const notes = rowNotes(row);
    const hasMri = parseIsoDate(row["MRI Date"]);
    if (hasMri) return;
    if (/ineligible|Inelligible/i.test(notes)) ids.push(`${pid} ineligible for MRI`);
    else if (/opt out|opted out|OPT OUT/i.test(notes)) ids.push(`${pid} Opted out of MRI`);
  });
  if (!ids.length) return null;
  return `* ${ids.slice(0, 3).join(" ; ")}`;
}

function dashCell(value) {
  return value != null && value !== 0 ? String(value) : "–";
}

function addCardinalBar() {}
function addGoldAccent() {}

function addSlideTitle(slide, title, subtitle) {
  slide.addText(title, {
    x: 0.55,
    y: 0.4,
    w: 9.0,
    h: 0.65,
    fontSize: 30,
    bold: true,
    color: DECK.title,
    fontFace: DECK.font,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.55,
      y: 1.02,
      w: 9.0,
      h: 0.35,
      fontSize: 15,
      color: DECK.muted,
      fontFace: DECK.font,
    });
  }
}

function addBodyLines(slide, lines, opts = {}) {
  const parts = lines.map((line) => {
    if (typeof line === "string") {
      return { text: line, options: { breakLine: true, bullet: false } };
    }
    return {
      text: line.text,
      options: {
        breakLine: true,
        bullet: line.bullet ?? false,
        bold: line.bold ?? false,
        fontSize: line.fontSize ?? opts.fontSize ?? 16,
        color: line.color ?? DECK.body,
      },
    };
  });
  slide.addText(parts, {
    x: opts.x ?? 0.65,
    y: opts.y ?? 1.45,
    w: opts.w ?? 8.8,
    h: opts.h ?? 4.5,
    fontSize: opts.fontSize ?? 16,
    color: DECK.body,
    fontFace: DECK.font,
    valign: "top",
  });
}

function addBullets(slide, items, opts = {}) {
  addBodyLines(
    slide,
    items.map((text) => ({ text, bullet: !String(text).startsWith("●") })),
    opts
  );
}

function addDataTable(slide, rows, opts = {}) {
  slide.addTable(rows, {
    x: opts.x ?? 0.55,
    y: opts.y ?? 1.55,
    w: opts.w ?? 8.9,
    colW: opts.colW,
    fontSize: opts.fontSize ?? 14,
    border: { type: "solid", color: DECK.tableBorder, pt: 0.5 },
    fontFace: DECK.font,
    color: DECK.body,
  });
}

function headerCell(text) {
  return { text, options: { bold: true, color: DECK.title } };
}

function addStatBlock(slide, x, y, number, labelLines, bullets, w = 2.7) {
  slide.addText(String(number), {
    x,
    y,
    w,
    h: 0.55,
    fontSize: 34,
    bold: true,
    color: DECK.title,
    fontFace: DECK.font,
  });
  slide.addText(labelLines.join("\n"), {
    x,
    y: y + 0.5,
    w,
    h: 0.55,
    fontSize: 13,
    bold: true,
    color: DECK.title,
    fontFace: DECK.font,
  });
  if (bullets?.length) {
    addBodyLines(
      slide,
      bullets.map((text) => ({ text, bullet: true })),
      { x: x + 0.05, y: y + 1.05, w, h: 1.5, fontSize: 13 }
    );
  }
}

function buildTitleSlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: DECK.bg };
  slide.addText("Diabetes Brain Study –\nMonthly Progress Meeting", {
    x: 0.75,
    y: 1.85,
    w: 8.5,
    h: 1.4,
    fontSize: 36,
    bold: true,
    color: DECK.title,
    fontFace: DECK.font,
    align: "center",
  });
  slide.addText(m.label, {
    x: 0.75,
    y: 3.35,
    w: 8.5,
    h: 0.55,
    fontSize: 24,
    color: DECK.body,
    fontFace: DECK.font,
    align: "center",
  });
}

function buildExecutiveSummarySlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: DECK.bg };
  addSlideTitle(slide, "Executive Summary");

  const s = m.scheduling;
  const next = addMonths(m.year, m.month, 1);
  const next2 = addMonths(m.year, m.month, 2);
  const futureLabel = `${MONTH_ABBR[next.month - 1]} - ${MONTH_ABBR[next2.month - 1]} (3rd and 4th year)`;
  const mriFootnote = mriFootnoteFromData();

  addStatBlock(slide, 0.55, 1.35, m.dropouts.total, ["Dropout", "Participants"], [
    `${m.dropouts.baseline} Baseline year`,
    `${m.dropouts.year2} Second year`,
    `${m.dropouts.year3} Third year`,
    `${m.dropouts.year4} Fourth year`,
  ]);

  slide.addText("Future Visits Scheduled", {
    x: 3.45,
    y: 1.35,
    w: 3.0,
    h: 0.3,
    fontSize: 14,
    bold: true,
    color: DECK.title,
    fontFace: DECK.font,
  });
  slide.addText(futureLabel, {
    x: 3.45,
    y: 1.65,
    w: 3.0,
    h: 0.25,
    fontSize: 12,
    color: DECK.muted,
    fontFace: DECK.font,
  });
  addBodyLines(
    slide,
    [
      `${s.futureVisits.mri} MRI's`,
      `${s.futureVisits.bd} Full Blood Draws`,
      `${s.futureVisits.np} Neuropsych Tests`,
      `${s.futureVisits.cv} Clinician Visits`,
    ].map((text) => ({ text, bullet: true })),
    { x: 3.55, y: 1.95, w: 3.0, h: 1.4, fontSize: 13 }
  );

  addStatBlock(
    slide,
    0.55,
    3.35,
    m.y3.completed ?? "—",
    ["Participants", "Completed 3rd Year"],
    [
      `${s.y3Components["MRI Date"] ?? "—"} MRI *`,
      `${s.y3Components["BD Date"] ?? "—"} Blood Draw`,
      `${s.y3Components["NP Date"] ?? "—"} Neuropsych Tests`,
      `${s.y3Components["CV Date"] ?? "—"} Clinician Visits`,
    ],
    4.0
  );

  addStatBlock(
    slide,
    5.0,
    3.35,
    m.y4.completed ?? "—",
    ["Participants", "Completed 4th Year"],
    [
      `${s.y4Components["BD Date"] ?? "—"} Blood Draw`,
      `${s.y4Components["NP Date"] ?? "—"} Neuropsych Tests`,
      `${s.y4Components["CV Date"] ?? "—"} Clinician Visits`,
    ],
    4.0
  );

  if (mriFootnote) {
    slide.addText(mriFootnote, {
      x: 0.55,
      y: 5.05,
      w: 8.8,
      h: 0.3,
      fontSize: 10,
      italic: true,
      color: DECK.muted,
      fontFace: DECK.font,
    });
  }
}

function buildCompletedVisitsSlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: DECK.bg };
  const monthName = MONTH_NAMES[m.month - 1];
  addSlideTitle(slide, `All Completed Visits in ${monthName}`);

  const rows = [
    [headerCell("Visit Type"), headerCell("3rd Year"), headerCell("4th Year")],
    ["MRI", dashCell(m.y3Visits.mri), "–"],
    ["Blood Draw", dashCell(m.y3Visits.bloodDraw), dashCell(m.y4Visits.bloodDraw)],
    ["NP Visit", dashCell(m.y3Visits.np), dashCell(m.y4Visits.np)],
    ["Clinician Visit", dashCell(m.y3Visits.cv), dashCell(m.y4Visits.cv)],
  ];

  addDataTable(slide, rows, { y: 1.55, colW: [3.0, 2.8, 2.8], fontSize: 15 });

  const pending = (DATA.third_year_scheduling || []).find((row) => {
    const d = parseIsoDate(row["MRI Date"] || row["BD Date"]);
    return d && d.getFullYear() === m.year && d.getMonth() + 1 === m.month && d.getDate() >= 28;
  });
  if (pending?.participant_id) {
    const d = parseIsoDate(pending["MRI Date"] || pending["BD Date"]);
    const day = d.toLocaleDateString("en-US", { weekday: "long", month: "numeric", day: "numeric", year: "2-digit" });
    slide.addText(`${pending.participant_id} expected ${day} to complete BD and MRI visit.`, {
      x: 0.55,
      y: 4.15,
      w: 8.8,
      h: 0.35,
      fontSize: 13,
      color: DECK.body,
      fontFace: DECK.font,
    });
  }
}

function buildYear3ActivitySlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: DECK.bg };
  const monthName = MONTH_NAMES[m.month - 1];
  const a = m.scheduling.y3Activity;
  const nextName = MONTH_NAMES[a.nextMonth.month - 1];
  const y3Total =
    m.y3Visits.mri + m.y3Visits.bloodDraw + m.y3Visits.np + m.y3Visits.cv;
  const y3Summary = [
    m.y3Visits.mri ? `${m.y3Visits.mri} MRI` : null,
    m.y3Visits.bloodDraw ? `${m.y3Visits.bloodDraw} BD` : null,
    m.y3Visits.np ? `${m.y3Visits.np} NP` : null,
    m.y3Visits.cv ? `${m.y3Visits.cv} CV` : null,
  ]
    .filter(Boolean)
    .join(", ");
  const dueScheduled = Math.max(0, a.dueCount - a.dueComplete);

  addBodyLines(slide, [
    `Total Visits in ${monthName}: ${y3Total} (${y3Summary})`,
    `Participants with ${monthName} Due Dates (n=${a.dueCount})`,
    { text: `${a.dueComplete} completed all visits`, bullet: true },
    { text: `${dueScheduled} successfully scheduled`, bullet: true },
    `Scheduled Visits for ${nextName} (to Date):`,
    `${a.nextVisitTotal} (${a.nextVisitSummary})`,
    `Participants with ${nextName} Due Dates (n=${a.nextDueCount})`,
    { text: `${a.nextDueCompletedEarly || 0} completed in ${monthName.slice(0, 3).toLowerCase()}`, bullet: true },
    { text: `${Math.max(0, a.nextDueCount - (a.nextDueCompletedEarly || 0))} successfully scheduled`, bullet: true },
    { text: "3rd Year Study Visits", bold: true, fontSize: 22 },
  ], { y: 0.55, fontSize: 16 });
}

function buildYear3OverviewSlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: DECK.bg };
  addSlideTitle(slide, "Year 3 Visit Status", `As of ${asOfDateShort(m.year, m.month)}`);

  addBodyLines(slide, [
    { text: `${m.y3.completed ?? "—"} participants`, bold: true, fontSize: 18 },
    "Have completed all required visits (BD, MRI, CV, NP)",
    "Visits are caught up through May 2026",
    "Year 3 study period: May 2025 - Jan/Feb 2027",
  ], { y: 1.55, fontSize: 17 });
}

function buildYear3InProgressSlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: DECK.bg };
  addSlideTitle(slide, "Year 3 Visit Status");

  const groups = m.scheduling.y3InProgress;
  const total = groups.reduce((s, g) => s + g.count, 0);

  slide.addText(`${total} participants in progress`, {
    x: 0.55,
    y: 1.2,
    w: 8.9,
    h: 0.4,
    fontSize: 17,
    bold: true,
    color: DECK.title,
    fontFace: DECK.font,
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

  addDataTable(slide, rows, { y: 1.75, colW: [2.4, 2.0, 4.2], fontSize: 13 });
}

function buildYear3ToContactSlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: DECK.bg };
  addSlideTitle(slide, "Year 3 Visit Status");

  const groups = m.scheduling.y3ToContact;
  const total = groups.reduce((s, g) => s + g.count, 0);

  slide.addText(`${total} participants to be contacted`, {
    x: 0.55,
    y: 1.15,
    w: 8.9,
    h: 0.35,
    fontSize: 17,
    bold: true,
    color: DECK.title,
    fontFace: DECK.font,
  });
  slide.addText(`${m.y3.remaining ?? 157} Participants Expected to Complete Year 3 by Jan/Feb 2027`, {
    x: 0.55,
    y: 1.5,
    w: 8.9,
    h: 0.3,
    fontSize: 13,
    color: DECK.body,
    fontFace: DECK.font,
  });

  const rows = [
    [headerCell("Month Due"), headerCell("# Participants")],
    ...groups.map((g) => [g.label, String(g.count)]),
    [{ text: "Total", options: { bold: true } }, { text: String(total), options: { bold: true } }],
  ];

  addDataTable(slide, rows, { y: 1.95, colW: [4.2, 4.2], fontSize: 14 });
}

function buildYear4ActivitySlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: DECK.bg };
  const monthName = MONTH_NAMES[m.month - 1];
  const a = m.scheduling.y4Activity;
  const nextName = MONTH_NAMES[a.nextMonth.month - 1];
  const y4Total = m.y4Visits.bloodDraw + m.y4Visits.np + m.y4Visits.cv;
  const y4Summary = [
    m.y4Visits.bloodDraw ? `${m.y4Visits.bloodDraw} BD` : null,
    m.y4Visits.np ? `${m.y4Visits.np} NP` : null,
    m.y4Visits.cv ? `${m.y4Visits.cv} CV` : null,
  ]
    .filter(Boolean)
    .join(", ");

  addBodyLines(slide, [
    { text: "4th Year Study Visits 🎉", bold: true, fontSize: 22 },
    `Total Visits in ${monthName}: ${y4Total} (${y4Summary})`,
    `Participants with ${monthName} Due Dates (n=${a.dueCount})`,
    { text: `${a.dueComplete} completed all visits`, bullet: true },
    { text: `${Math.max(0, a.dueCount - a.dueComplete)} rescheduled for ${nextName.slice(0, 3).toLowerCase()}`, bullet: true },
    `Scheduled Visits for ${nextName} (to Date): ${a.nextVisitTotal} (${a.nextVisitSummary})`,
    `Participants with ${nextName} Due Dates (n=${a.nextDueCount})`,
    { text: `${a.nextDueCount} pending scheduling`, bullet: true },
  ], { y: 0.55, fontSize: 16 });
}

function buildYear4ToContactSlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: DECK.bg };
  addSlideTitle(slide, "Year 4 Visit Status");

  const groups = m.scheduling.y4ToContact;
  const total = groups.reduce((s, g) => s + g.count, 0);

  addBodyLines(slide, [
    { text: `${total || 26} participants to be contacted through December ${m.year}`, bold: true },
    "65 participants to be contacted January–August 2027 (3rd-year rollovers)",
    "Expected Study Completion: Dec 2027/Jan 2028",
  ], { y: 1.15, fontSize: 14 });

  const rows = [
    [headerCell("Month Due"), headerCell("# Participants")],
    ...groups.map((g) => [g.label, String(g.count)]),
    [{ text: "Total", options: { bold: true } }, { text: String(total), options: { bold: true } }],
  ];

  addDataTable(slide, rows, { y: 2.35, colW: [4.2, 4.2], fontSize: 14 });
}

function buildStudyProgressSlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: DECK.bg };
  addSlideTitle(slide, "Study Progress", `Participant Flow (N = ${m.enrolled})`);

  addBodyLines(slide, [
    { text: "Baseline", bold: true },
    { text: `${m.enrolled} enrolled`, bullet: true },
    { text: `${m.baseline.completed ?? "—"} completed Baseline assessments`, bullet: true },
    { text: `${m.dropouts.baseline} participants dropped out`, bullet: true },
    { text: "Year 2", bold: true },
    { text: `${m.y2.completed ?? "—"} participants completed Year 2`, bullet: true },
    { text: "+1 participant advanced to Year 3", bullet: true },
    { text: `${m.dropouts.year2} additional participants dropped out`, bullet: true },
    { text: "Year 3 (ongoing)", bold: true },
    { text: `${m.y3.completed ?? "—"} participants completed Year 3`, bullet: true },
    { text: `${m.dropouts.year3} additional participants dropped out`, bullet: true },
    { text: "Year 4 (ongoing)", bold: true },
    { text: `${m.y4.completed ?? "—"} participants completed Year 4`, bullet: true },
  ], { y: 1.45, fontSize: 16 });
}

function buildQuestionsSlide(pptx, m) {
  const slide = pptx.addSlide();
  slide.background = { color: DECK.bg };
  addSlideTitle(slide, "Questions / Comments");

  addBodyLines(slide, [
    { text: "Funding account end date: 8/31/2026", bullet: true },
    "Once the funding account has been extended, we'll need to update the CIA scheduler to be able to book MRI's past october.",
  ], { y: 1.55, fontSize: 16 });
}

function buildThankYouSlide(pptx) {
  const slide = pptx.addSlide();
  slide.background = { color: DECK.bg };
  slide.addText("THANK YOU :)", {
    x: 0.7,
    y: 2.35,
    w: 8.6,
    h: 0.9,
    fontSize: 40,
    bold: true,
    color: DECK.title,
    align: "center",
    fontFace: DECK.font,
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
  pptx.company = "DBS";
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

  const monthlyRows = [
    ["3rd Year MRI", metrics.y3Visits.mri],
    ["3rd Year Blood Draw", metrics.y3Visits.bloodDraw],
    ["3rd Year NP Visit", metrics.y3Visits.np],
    ["3rd Year Clinician Visit", metrics.y3Visits.cv],
    ["4th Year Blood Draw", metrics.y4Visits.bloodDraw],
    ["4th Year NP Visit", metrics.y4Visits.np],
    ["4th Year Clinician Visit", metrics.y4Visits.cv],
    [`${monthName} 3rd Year visit total`, s.y3Activity.visitTotal],
    [`${monthName} 4th Year visit total`, s.y4Activity.visitTotal],
    ["Next month scheduled (3rd year)", s.y3Activity.nextVisitTotal],
    ["Next month scheduled (4th year)", s.y4Activity.nextVisitTotal],
  ];

  const executiveRows = [
    ["Total Dropouts", metrics.dropouts.total],
    ["Baseline / Y2 / Y3 / Y4 dropouts", `${metrics.dropouts.baseline} / ${metrics.dropouts.year2} / ${metrics.dropouts.year3} / ${metrics.dropouts.year4}`],
    ["Future MRI (next 2 mo)", s.futureVisits.mri],
    ["Future BD (next 2 mo)", s.futureVisits.bd],
    ["Future NP (next 2 mo)", s.futureVisits.np],
    ["Future CV (next 2 mo)", s.futureVisits.cv],
    ["Year 3 completed", metrics.y3.completed],
    ["Y3 MRI / BD / NP / CV", `${s.y3Components["MRI Date"]} / ${s.y3Components["BD Date"]} / ${s.y3Components["NP Date"]} / ${s.y3Components["CV Date"]}`],
    ["Year 4 completed", metrics.y4.completed],
    ["Y3 in progress", s.y3InProgress.reduce((n, g) => n + g.count, 0)],
    ["Y3 to be contacted", s.y3ToContact.reduce((n, g) => n + g.count, 0)],
  ];

  const snapshotRows = [
    ["Enrolled", metrics.enrolled],
    ["Active participants", metrics.active],
    ["Baseline completed", metrics.baseline.completed],
    ["Year 2 completed", metrics.y2.completed],
    ["Year 3 completed", metrics.y3.completed],
    ["Year 4 completed", metrics.y4.completed],
    ["Recruitment all-time", metrics.recruitment.grand_total ?? metrics.recruitment.summary_n],
  ];

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
      `${metrics.label} — This Month's Visits`,
      "Visit counts and scheduling activity for the selected report month.",
      monthlyRows
    ) +
    renderSection(
      "Executive Summary (deck slide 2)",
      "Dropouts, future visits, and completion totals that appear in the PowerPoint.",
      executiveRows
    ) +
    renderSection(
      "Current Study Status",
      "Point-in-time totals from the latest sync (same across all report months).",
      snapshotRows
    );
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
