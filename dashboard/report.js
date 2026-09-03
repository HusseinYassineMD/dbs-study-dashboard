/** Monthly progress report → USC-branded PowerPoint (PptxGenJS). */

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
  if (!s || /opt out|n\/a|^nan$/i.test(s)) return null;
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

/** Sum recruitment leads from period columns overlapping the selected calendar month. */
function recruitmentLeadsForMonth(year, month) {
  const sources = DATA.recruitment?.sources || [];
  const periods = DATA.recruitment?.periods || [];
  const monthStart = new Date(year, month - 1, 1);
  const monthEnd = new Date(year, month, 0);

  const monthRe = new RegExp(
    `(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\\.?\\s+'?(\\d{2})`,
    "gi"
  );

  function parsePeriodDate(token) {
    const m = token.trim().match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+'?(\d{2})/i);
    if (!m) return null;
    const abbr = m[1].slice(0, 3).toLowerCase();
    const idx = MONTH_NAMES.findIndex((n) => n.toLowerCase().startsWith(abbr));
    if (idx < 0) return null;
    const y = 2000 + Number(m[2]);
    return new Date(y, idx, 1);
  }

  let total = 0;
  const matchedPeriods = [];

  periods.forEach((period) => {
    const inner = period.match(/\(([^)]+)\)/);
    if (!inner) return;
    const parts = inner[1].split(/–|-/);
    if (parts.length < 2) return;
    const start = parsePeriodDate(parts[0]);
    const end = parsePeriodDate(parts[1]);
    if (!start || !end) return;
    const endLast = new Date(end.getFullYear(), end.getMonth() + 1, 0);
    if (start <= monthEnd && endLast >= monthStart) {
      matchedPeriods.push(period);
      sources.forEach((s) => {
        total += Number(s.timeline?.[period]) || 0;
      });
    }
  });

  return { total, matchedPeriods };
}

function collectReportMetrics(year, month) {
  const baseline = progressRow("baseline");
  const y2 = progressRow("2");
  const y3 = progressRow("3");
  const y4 = progressRow("4");

  const thirdSched = DATA.third_year_scheduling || [];
  const fourthSched = DATA.fourth_year_scheduling || [];

  const y3Visits = {
    mri: countFieldInMonth(thirdSched, "MRI Date", year, month),
    bloodDraw: countFieldInMonth(thirdSched, "BD Date", year, month),
    np: countFieldInMonth(thirdSched, "NP Date", year, month),
    cv: countFieldInMonth(thirdSched, "CV Date", year, month),
  };
  const y4Visits = {
    bloodDraw: countFieldInMonth(fourthSched, "BD Date", year, month),
    np: countFieldInMonth(fourthSched, "NP Date", year, month),
    cv: countFieldInMonth(fourthSched, "CV Date", year, month),
  };

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
  const lastTrend6 = lastTrendOnOrBefore("visit_trends_sheet6", year, month);

  const y2cv = y2Monthly.cv != null ? y2Monthly.cv : displayCount(trend5?.["CV Monthly"]);
  const y2np = y2Monthly.np != null ? y2Monthly.np : displayCount(trend5?.["NP Monthly"]);
  const y2Cumulative =
    trend5?.["CV Cumulative"] ??
    trend6?.["CV Cumulative"] ??
    lastTrend5?.["CV Cumulative"] ??
    lastTrend6?.["CV Cumulative"] ??
    null;
  const y2CumulativeCarried = !trend5 && !trend6 && lastTrend5 != null;
  const recruitmentMonth = recruitmentLeadsForMonth(year, month);

  const dropouts = {
    baseline: baseline.dropped_out_dq ?? 0,
    year2: y2.dropped_out_dq ?? 0,
    year3: y3.dropped_out_dq ?? 0,
    year4: y4.dropped_out_dq ?? 0,
    total: [baseline, y2, y3, y4].reduce((s, r) => s + (r.dropped_out_dq || 0), 0),
  };

  const recruitment = DATA.recruitment || {};
  const genotype = DATA.genotype_summary || {};

  const y2CompletedMonth = countCompletedVisitsInMonth("second_year", year, month);
  const y3CompletedMonth = countCompletedVisitsInMonth("third_year", year, month);
  const y4CompletedMonth = countCompletedVisitsInMonth("fourth_year", year, month);

  const y3TotalMonthVisits = y3Visits.mri + y3Visits.bloodDraw + y3Visits.np + y3Visits.cv;
  const y4TotalMonthVisits = y4Visits.bloodDraw + y4Visits.np + y4Visits.cv;

  const monthlyVisitTotal = y2cv + y2np + y3TotalMonthVisits + y4TotalMonthVisits;

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
    y3Visits,
    y4Visits,
    y2Monthly,
    y2cv,
    y2np,
    y2Cumulative,
    y2CumulativeCarried,
    trend5,
    trend6,
    lastTrend5,
    recruitmentMonth,
    y2CompletedMonth,
    y3CompletedMonth,
    y4CompletedMonth,
    y3TotalMonthVisits,
    y4TotalMonthVisits,
    monthlyVisitTotal,
    cumulativeY2AtMonth: y2Cumulative,
    recruitment,
    genotype,
    meetingDropouts: DATA.meeting_dropouts?.length ?? 0,
    mriOutcomes: DATA.mri_outcomes || [],
    screening: DATA.screening?.length ?? 0,
    studyPartnersComplete: (DATA.study_partner || []).filter(
      (r) => r.completed === 1 || r.completed === "1"
    ).length,
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

function addMetricTiles(slide, tiles, startY = 1.5) {
  const w = 2.1;
  const gap = 0.25;
  tiles.forEach((tile, i) => {
    const x = 0.55 + i * (w + gap);
    slide.addShape("rect", {
      x,
      y: startY,
      w,
      h: 1.35,
      fill: { color: USC.lightGray },
      line: { color: USC.cardinal, width: 1 },
    });
    slide.addText(String(tile.value), {
      x,
      y: startY + 0.2,
      w,
      h: 0.55,
      fontSize: 30,
      bold: true,
      color: USC.cardinal,
      align: "center",
      fontFace: "Arial",
    });
    slide.addText(tile.label, {
      x,
      y: startY + 0.78,
      w,
      h: 0.45,
      fontSize: 11,
      color: USC.gray,
      align: "center",
      fontFace: "Arial",
    });
  });
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
  slide.addText("Diabetes Brain Study", {
    x: 0.7,
    y: 1.5,
    w: 8.6,
    h: 0.9,
    fontSize: 40,
    bold: true,
    color: USC.white,
    fontFace: "Arial",
  });
  slide.addText("Monthly Progress Meeting", {
    x: 0.7,
    y: 2.35,
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
  slide.addText("USC Mark and Mary Stevens Neuroimaging and Informatics Institute", {
    x: 0.7,
    y: 5.05,
    w: 8.6,
    h: 0.35,
    fontSize: 11,
    color: USC.black,
    fontFace: "Arial",
  });
}

function buildExecutiveSummarySlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Executive Summary", m.label);

  addMetricTiles(slide, [
    { label: "2nd Yr CV This Month", value: m.y2cv },
    { label: "2nd Yr NP This Month", value: m.y2np },
    { label: "Year 3 Completed (now)", value: m.y3.completed ?? "—" },
    { label: "Year 4 Completed (now)", value: m.y4.completed ?? "—" },
  ]);

  addBullets(slide, [
    `${MONTH_NAMES[m.month - 1]} activity: ${m.monthlyVisitTotal} total visit events logged across study years`,
    `${m.y3Visits.mri} MRI · ${m.y3Visits.bloodDraw} BD · ${m.y3Visits.np} NP · ${m.y3Visits.cv} CV (3rd year, ${MONTH_NAMES[m.month - 1]})`,
    `${m.y4Visits.bloodDraw} BD · ${m.y4Visits.np} NP · ${m.y4Visits.cv} CV (4th year, ${MONTH_NAMES[m.month - 1]})`,
    `Recruitment leads this month: ${m.recruitmentMonth.total}`,
    `Total dropouts to date: ${m.dropouts.total} · Active participants: ${m.active}`,
  ], { y: 3.05, fontSize: 15 });
}

function buildMonthlyVisitsSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, `Completed Visits in ${MONTH_NAMES[m.month - 1]}`, m.label);

  const y2cv = m.y2cv;
  const y2np = m.y2np;

  const rows = [
    [
      { text: "Visit Type", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "2nd Year", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "3rd Year", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "4th Year", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
    ],
    ["Clinician Visit", String(y2cv), String(m.y3Visits.cv), String(m.y4Visits.cv)],
    ["NP Visit", String(y2np), String(m.y3Visits.np), String(m.y4Visits.np)],
    ["Blood Draw", "—", String(m.y3Visits.bloodDraw), String(m.y4Visits.bloodDraw)],
    ["MRI", "—", String(m.y3Visits.mri), "—"],
    [
      { text: "Completions logged", options: { bold: true } },
      { text: String(m.y2CompletedMonth), options: { bold: true } },
      { text: String(m.y3CompletedMonth), options: { bold: true } },
      { text: String(m.y4CompletedMonth), options: { bold: true } },
    ],
  ];

  slide.addTable(rows, {
    x: 0.55,
    y: 1.55,
    w: 8.9,
    colW: [2.5, 2.1, 2.1, 2.1],
    fontSize: 13,
    border: { type: "solid", color: USC.cardinal, pt: 0.75 },
    fontFace: "Arial",
  });

  const extras = [];
  if (m.trend5 || m.y2Cumulative != null) {
    extras.push(
      `2nd year cumulative through ${MONTH_NAMES[m.month - 1]}: CV ${m.y2Cumulative ?? "—"}${m.y2CumulativeCarried ? " (latest available)" : ""}`
    );
  }
  if (m.recruitmentMonth.total) {
    extras.push(`Recruitment leads acquired: ${m.recruitmentMonth.total}`);
  }
  addBullets(slide, extras.length ? extras : ["No cumulative trend data for this month"], {
    y: 4.2,
    h: 1,
    fontSize: 13,
  });
}

function buildYear3StatusSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Year 3 Visit Status", `As of ${m.label}`);

  addMetricTiles(slide, [
    { label: "Completed Y3 (now)", value: m.y3.completed ?? "—" },
    { label: "This Month CV", value: m.y3Visits.cv },
    { label: "This Month NP", value: m.y3Visits.np },
    { label: "Completions Logged", value: m.y3CompletedMonth },
  ], 1.45);

  addBullets(slide, [
    `${m.y3.completed ?? "—"} participants have completed all required Year 3 visits (current total)`,
    `${MONTH_NAMES[m.month - 1]} ${m.year}: ${m.y3Visits.mri} MRI · ${m.y3Visits.bloodDraw} BD · ${m.y3Visits.np} NP · ${m.y3Visits.cv} CV`,
    `${m.y3.remaining ?? "—"} participants expected to complete Year 3`,
    `${m.dropouts.year3} Year 3 dropouts to date`,
  ], { y: 3.15, fontSize: 15 });
}

function buildYear4StatusSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Year 4 Visit Status", `As of ${m.label}`);

  addMetricTiles(slide, [
    { label: "Completed Y4 (now)", value: m.y4.completed ?? "—" },
    { label: "This Month BD", value: m.y4Visits.bloodDraw },
    { label: "This Month NP/CV", value: `${m.y4Visits.np}/${m.y4Visits.cv}` },
    { label: "Completions Logged", value: m.y4CompletedMonth },
  ], 1.45);

  addBullets(slide, [
    `${m.y4.completed ?? "—"} participants completed Year 4 (current total)`,
    `${MONTH_NAMES[m.month - 1]} ${m.year}: ${m.y4Visits.bloodDraw} BD · ${m.y4Visits.np} NP · ${m.y4Visits.cv} CV`,
    `${m.y4.remaining ?? "—"} participants expected to complete Year 4`,
    `Expected study completion: Dec 2027 / Jan 2028`,
  ], { y: 3.15, fontSize: 15 });
}

function buildStudyProgressSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Study Progress", `Current snapshot · N = ${m.enrolled}`);

  addBullets(slide, [
    `Baseline: ${m.enrolled} enrolled · ${m.baseline.completed ?? "—"} completed · ${m.dropouts.baseline} dropped out`,
    `Year 2: ${m.y2.completed ?? "—"} completed · ${m.dropouts.year2} additional dropouts`,
    `Year 3 (ongoing): ${m.y3.completed ?? "—"} completed · ${m.dropouts.year3} additional dropouts`,
    `Year 4 (ongoing): ${m.y4.completed ?? "—"} completed · ${m.dropouts.year4} additional dropouts`,
    m.cumulativeY2AtMonth != null
      ? `At end of ${MONTH_NAMES[m.month - 1]} ${m.year}: ~${m.cumulativeY2AtMonth} Year 2 CV visits cumulative`
      : "Historical cumulative not available for this month",
  ], { y: 1.55, fontSize: 16 });

  const rows = [
    [
      { text: "Year", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "Completed", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "Dropped", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "Remaining", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
    ],
    ["Baseline", m.baseline.completed, m.dropouts.baseline, m.baseline.remaining],
    ["Year 2", m.y2.completed, m.dropouts.year2, m.y2.remaining],
    ["Year 3", m.y3.completed, m.dropouts.year3, m.y3.remaining],
    ["Year 4", m.y4.completed, m.dropouts.year4, m.y4.remaining],
  ];

  slide.addTable(rows, {
    x: 0.55,
    y: 3.55,
    w: 8.9,
    colW: [1.8, 2.2, 2.2, 2.2],
    fontSize: 12,
    border: { type: "solid", color: USC.cardinal, pt: 0.75 },
    fontFace: "Arial",
  });
}

function buildRecruitmentSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(
    slide,
    "Recruitment Summary",
    `${MONTH_NAMES[m.month - 1]} leads: ${m.recruitmentMonth.total} · All-time: ${m.recruitment.grand_total ?? "—"}`
  );

  const sources = (m.recruitment.sources || []).slice(0, 5);
  const rows = [
    [
      { text: "Source", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "All-Time", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "This Month", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
    ],
    ...sources.map((s) => {
      const monthLeads = (m.recruitmentMonth.matchedPeriods || []).reduce(
        (sum, p) => sum + (Number(s.timeline?.[p]) || 0),
        0
      );
      return [s.source, String(s.total_leads ?? "—"), String(monthLeads)];
    }),
  ];

  slide.addTable(rows, {
    x: 0.55,
    y: 1.55,
    w: 8.9,
    colW: [3.5, 2.5, 2.5],
    fontSize: 13,
    border: { type: "solid", color: USC.cardinal, pt: 0.75 },
    fontFace: "Arial",
  });

  addBullets(
    slide,
    m.recruitmentMonth.matchedPeriods?.length
      ? [`Recruitment periods in ${MONTH_NAMES[m.month - 1]}: ${m.recruitmentMonth.matchedPeriods.length}`]
      : [`No recruitment period overlap for ${MONTH_NAMES[m.month - 1]} ${m.year}`],
    { y: 4.2, fontSize: 13 }
  );
}

function buildGenotypeSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Genotype & Clinical Summary");

  addMetricTiles(slide, [
    { label: "Diabetics", value: m.genotype.total_diabetics ?? "—" },
    { label: "Non-Diabetic", value: m.genotype.total_non_diabetic ?? "—" },
    { label: "High AD Risk", value: m.genotype.total_high_ad_risk ?? "—" },
    { label: "Low AD Risk", value: m.genotype.total_low_ad_risk ?? "—" },
  ]);

  const mriLines = (m.mriOutcomes || []).map(
    (r) =>
      `${r.Timepoint}: ${r["Did Come Out"] ?? "—"} came out · ${r["Did Not Come Out"] ?? "—"} did not (${pct(r["Failure %"])} failure)`
  );
  addBullets(slide, mriLines.length ? mriLines : ["MRI outcome data not available"], {
    y: 3.05,
    fontSize: 14,
  });
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
  slide.addText("Thank You", {
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
  slide.addText("Questions / Comments", {
    x: 0.7,
    y: 3.05,
    w: 8.6,
    h: 0.45,
    fontSize: 20,
    color: USC.gold,
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
  buildMonthlyVisitsSlide(pptx, metrics);
  buildYear3StatusSlide(pptx, metrics);
  buildYear4StatusSlide(pptx, metrics);
  buildStudyProgressSlide(pptx, metrics);
  buildRecruitmentSlide(pptx, metrics);
  buildGenotypeSlide(pptx, metrics);
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

  // Fill every month between earliest and latest (capped at Sep 2026)
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

  const y2CumulativeLabel = metrics.y2CumulativeCarried
    ? `${metrics.y2Cumulative} (latest trend data — 2nd year sheet ends Mar 2026)`
    : metrics.y2Cumulative ?? 0;

  const monthlyRows = [
    ["2nd Year CV Completed", metrics.y2cv],
    ["2nd Year NP Completed", metrics.y2np],
    ["2nd Year CV Cumulative", y2CumulativeLabel],
    ["2nd Year Completions Logged", metrics.y2CompletedMonth],
    ["3rd Year MRI", metrics.y3Visits.mri],
    ["3rd Year Blood Draw", metrics.y3Visits.bloodDraw],
    ["3rd Year NP Visit", metrics.y3Visits.np],
    ["3rd Year Clinician Visit", metrics.y3Visits.cv],
    ["3rd Year Completions Logged", metrics.y3CompletedMonth],
    ["4th Year Blood Draw", metrics.y4Visits.bloodDraw],
    ["4th Year NP Visit", metrics.y4Visits.np],
    ["4th Year Clinician Visit", metrics.y4Visits.cv],
    ["4th Year Completions Logged", metrics.y4CompletedMonth],
    ["Recruitment Leads (this month)", metrics.recruitmentMonth.total],
    ["Total Visit Events (this month)", metrics.monthlyVisitTotal],
  ];

  const snapshotRows = [
    ["Year 3 Completed (current)", metrics.y3.completed],
    ["Year 4 Completed (current)", metrics.y4.completed],
    ["Active Participants", metrics.active],
    ["Total Dropouts", metrics.dropouts.total],
    ["Recruitment All-Time", metrics.recruitment.grand_total ?? metrics.recruitment.summary_n],
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
      `${metrics.label} — This Month's Activity`,
      "These metrics change when you select a different month.",
      monthlyRows
    ) +
    renderSection(
      "Current Study Status",
      "Point-in-time totals from the latest Google Sheets sync (same across all months).",
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
    badge.textContent = `Live preview · ${metrics.label} · 2nd Yr CV: ${metrics.y2cv} · 3rd Yr: ${metrics.y3Visits.mri}/${metrics.y3Visits.bloodDraw}/${metrics.y3Visits.np}/${metrics.y3Visits.cv} · 4th Yr: ${metrics.y4Visits.bloodDraw}/${metrics.y4Visits.np}/${metrics.y4Visits.cv}`;
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
