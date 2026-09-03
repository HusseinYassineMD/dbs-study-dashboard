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

function parseIsoDate(value) {
  if (!value) return null;
  const d = new Date(value);
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
  const abbr = MONTH_NAMES[month - 1].slice(0, 3);
  const yy = String(year).slice(-2);
  const patterns = [
    `${abbr} '${String(year).slice(-2)}`,
    `${abbr} '${yy}`,
    `${abbr} ${yy}`,
    `${abbr}'${yy}`,
    `${MONTH_NAMES[month - 1]} '${yy}`,
    `${MONTH_NAMES[month - 1]} ${year}`,
    `${MONTH_NAMES[month - 1]} '${String(year).slice(-2)}`,
  ];
  if (!monthlyRows?.length) return null;
  const keys = Object.keys(monthlyRows[0]).filter(
    (k) => !["visit_type", "total_completed"].includes(k)
  );
  for (const pattern of patterns) {
    const match = keys.find((k) => k.toLowerCase().startsWith(pattern.toLowerCase()));
    if (match) return match;
  }
  return keys.find((k) => k.toLowerCase().includes(abbr.toLowerCase()) && k.includes(String(year).slice(-2))) || null;
}

function visitTrendForMonth(year, month) {
  return (DATA.visit_trends_sheet5 || []).find((r) => inCalendarMonth(r.Date, year, month)) || null;
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
    cv: monthCol && cvRow ? Number(cvRow[monthCol]) || 0 : null,
    np: monthCol && npRow ? Number(npRow[monthCol]) || 0 : null,
  };

  const trend = visitTrendForMonth(year, month);
  const dropouts = {
    baseline: baseline.dropped_out_dq ?? 0,
    year2: y2.dropped_out_dq ?? 0,
    year3: y3.dropped_out_dq ?? 0,
    year4: y4.dropped_out_dq ?? 0,
    total: [baseline, y2, y3, y4].reduce((s, r) => s + (r.dropped_out_dq || 0), 0),
  };

  const recruitment = DATA.recruitment || {};
  const genotype = DATA.genotype_summary || {};

  const y3CompletedMonth = countCompletedVisitsInMonth("third_year", year, month);
  const y4CompletedMonth = countCompletedVisitsInMonth("fourth_year", year, month);

  const y3TotalMonthVisits = y3Visits.mri + y3Visits.bloodDraw + y3Visits.np + y3Visits.cv;
  const y4TotalMonthVisits = y4Visits.bloodDraw + y4Visits.np + y4Visits.cv;

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
    trend,
    y3CompletedMonth,
    y4CompletedMonth,
    y3TotalMonthVisits,
    y4TotalMonthVisits,
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
    { label: "Total Dropouts", value: m.dropouts.total },
    { label: "Year 3 Completed", value: m.y3.completed ?? "—" },
    { label: "Year 4 Completed", value: m.y4.completed ?? "—" },
    { label: "Active Participants", value: m.active },
  ]);

  addBullets(slide, [
    `${m.dropouts.baseline} Baseline year · ${m.dropouts.year2} Second year · ${m.dropouts.year3} Third year · ${m.dropouts.year4} Fourth year dropouts`,
    `${m.y3.completed ?? "—"} participants completed Year 3 (BD, MRI, CV, NP)`,
    `${m.y4.completed ?? "—"} participants completed Year 4`,
    `${m.enrolled} enrolled · ${m.baseline.completed ?? "—"} baseline completed · ${m.y2.completed ?? "—"} Year 2 completed`,
    `Data synced ${m.generatedAt ? new Date(m.generatedAt).toLocaleString() : "from dashboard"}`,
  ], { y: 3.05, fontSize: 15 });
}

function buildMonthlyVisitsSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, `Completed Visits in ${MONTH_NAMES[m.month - 1]}`, "Visit counts from scheduling records");

  const rows = [
    [
      { text: "Visit Type", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "3rd Year", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "4th Year", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
    ],
    ["MRI", String(m.y3Visits.mri), "—"],
    ["Blood Draw", String(m.y3Visits.bloodDraw), String(m.y4Visits.bloodDraw)],
    ["NP Visit", String(m.y3Visits.np), String(m.y4Visits.np)],
    ["Clinician Visit", String(m.y3Visits.cv), String(m.y4Visits.cv)],
    [
      { text: "Total", options: { bold: true } },
      { text: String(m.y3TotalMonthVisits), options: { bold: true } },
      { text: String(m.y4TotalMonthVisits), options: { bold: true } },
    ],
  ];

  slide.addTable(rows, {
    x: 0.55,
    y: 1.55,
    w: 8.9,
    colW: [3.5, 2.7, 2.7],
    fontSize: 14,
    border: { type: "solid", color: USC.cardinal, pt: 0.75 },
    fontFace: "Arial",
  });

  const extras = [
    `3rd year participant completions logged: ${m.y3CompletedMonth}`,
    `4th year participant completions logged: ${m.y4CompletedMonth}`,
  ];
  if (m.trend) {
    extras.push(
      `2nd year trend (Sheet5): CV ${m.trend["CV Monthly"] ?? 0}, NP ${m.trend["NP Monthly"] ?? 0} this month`
    );
  }
  if (m.y2Monthly.cv != null) {
    extras.push(`2nd year monthly sheet: CV ${m.y2Monthly.cv}, NP ${m.y2Monthly.np}`);
  }
  addBullets(slide, extras, { y: 4.05, h: 1.2, fontSize: 13 });
}

function buildYear3StatusSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Year 3 Visit Status", `As of ${m.label}`);

  addMetricTiles(slide, [
    { label: "Completed Y3", value: m.y3.completed ?? "—" },
    { label: "Remaining", value: m.y3.remaining ?? "—" },
    { label: "% of Active", value: pct(m.y3.pct_completed_active) },
    { label: "Dropouts", value: m.dropouts.year3 },
  ], 1.45);

  addBullets(slide, [
    `${m.y3.completed ?? "—"} participants have completed all required Year 3 visits (BD, MRI, CV, NP)`,
    `${m.y3.remaining ?? "—"} participants expected to complete Year 3`,
    `${m.dropouts.year3} additional Year 3 dropouts to date`,
    `Year 3 study period: May 2025 – Jan/Feb 2027`,
    `${m.y3CompletedMonth} Year 3 completion entries recorded in ${MONTH_NAMES[m.month - 1]}`,
  ], { y: 3.15, fontSize: 15 });
}

function buildYear4StatusSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Year 4 Visit Status", `As of ${m.label}`);

  addMetricTiles(slide, [
    { label: "Completed Y4", value: m.y4.completed ?? "—" },
    { label: "Remaining", value: m.y4.remaining ?? "—" },
    { label: "% of Active", value: pct(m.y4.pct_completed_active) },
    { label: "Dropouts", value: m.dropouts.year4 },
  ], 1.45);

  addBullets(slide, [
    `${m.y4.completed ?? "—"} participants completed Year 4`,
    `${m.y4.remaining ?? "—"} participants expected to complete Year 4`,
    `Expected study completion: Dec 2027 / Jan 2028`,
    `${m.y4CompletedMonth} Year 4 completion entries in ${MONTH_NAMES[m.month - 1]}`,
    `${m.y4Visits.bloodDraw} BD · ${m.y4Visits.np} NP · ${m.y4Visits.cv} CV scheduled/completed in ${MONTH_NAMES[m.month - 1]}`,
  ], { y: 3.15, fontSize: 15 });
}

function buildStudyProgressSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Study Progress", `Participant Flow (N = ${m.enrolled})`);

  addBullets(slide, [
    `Baseline: ${m.enrolled} enrolled · ${m.baseline.completed ?? "—"} completed · ${m.dropouts.baseline} dropped out`,
    `Year 2: ${m.y2.completed ?? "—"} completed · ${m.dropouts.year2} additional dropouts`,
    `Year 3 (ongoing): ${m.y3.completed ?? "—"} completed · ${m.dropouts.year3} additional dropouts · ${pct(m.y3.pct_completed_total)} of enrolled`,
    `Year 4 (ongoing): ${m.y4.completed ?? "—"} completed · ${m.dropouts.year4} additional dropouts · ${pct(m.y4.pct_completed_total)} of enrolled`,
    `Active participants: ${m.active}`,
  ], { y: 1.55, fontSize: 17 });

  const rows = [
    [
      { text: "Year", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "Completed", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "Dropped", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "Remaining", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "% Active", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
    ],
    ["Baseline", m.baseline.completed, m.dropouts.baseline, m.baseline.remaining, pct(m.baseline.pct_completed_active)],
    ["Year 2", m.y2.completed, m.dropouts.year2, m.y2.remaining, pct(m.y2.pct_completed_active)],
    ["Year 3", m.y3.completed, m.dropouts.year3, m.y3.remaining, pct(m.y3.pct_completed_active)],
    ["Year 4", m.y4.completed, m.dropouts.year4, m.y4.remaining, pct(m.y4.pct_completed_active)],
  ];

  slide.addTable(rows, {
    x: 0.55,
    y: 3.55,
    w: 8.9,
    colW: [1.5, 1.7, 1.5, 1.7, 1.5],
    fontSize: 12,
    border: { type: "solid", color: USC.cardinal, pt: 0.75 },
    fontFace: "Arial",
  });
}

function buildRecruitmentSlide(pptx, m) {
  const slide = pptx.addSlide();
  addSlideTitle(slide, "Recruitment Summary", `Total leads: ${m.recruitment.grand_total ?? m.recruitment.summary_n ?? "—"}`);

  const sources = (m.recruitment.sources || []).slice(0, 8);
  const rows = [
    [
      { text: "Source", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
      { text: "Total Leads", options: { bold: true, fill: { color: USC.cardinal }, color: USC.white } },
    ],
    ...sources.map((s) => [s.source, String(s.total_leads ?? "—")]),
  ];

  slide.addTable(rows, {
    x: 0.55,
    y: 1.55,
    w: 5.5,
    colW: [3.5, 2],
    fontSize: 13,
    border: { type: "solid", color: USC.cardinal, pt: 0.75 },
    fontFace: "Arial",
  });

  addBullets(slide, [
    `Screening records: ${m.screening}`,
    `Study partner surveys completed: ${m.studyPartnersComplete}`,
    `Meeting dropout records: ${m.meetingDropouts}`,
    `Recruitment periods tracked: ${(m.recruitment.periods || []).length}`,
  ], { x: 6.3, y: 1.55, w: 3.2, h: 3, fontSize: 13 });
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
    (r) => `${r.Timepoint}: ${r["Did Come Out"] ?? "—"} came out · ${r["Did Not Come Out"] ?? "—"} did not (${pct(r["Failure %"])} failure)`
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

function availableReportMonths() {
  const seen = new Set();
  const options = [];

  const add = (year, month) => {
    const key = `${year}-${month}`;
    if (seen.has(key)) return;
    seen.add(key);
    options.push({ year, month, label: monthLabel(year, month) });
  };

  (DATA.visit_trends_sheet5 || []).forEach((r) => {
    const d = parseIsoDate(r.Date);
    if (d) add(d.getFullYear(), d.getMonth() + 1);
  });

  if (DATA.meta?.generated_at) {
    const d = new Date(DATA.meta.generated_at);
    if (!Number.isNaN(d.getTime())) add(d.getFullYear(), d.getMonth() + 1);
  }

  const now = new Date();
  add(now.getFullYear(), now.getMonth() + 1);

  return options.sort((a, b) => b.year - a.year || b.month - a.month);
}

function renderReportPreview(metrics) {
  const el = document.getElementById("report-preview");
  if (!el) return;

  const rows = [
    ["Report Month", metrics.label],
    ["Enrolled", metrics.enrolled],
    ["Active Participants", metrics.active],
    ["Baseline Completed", metrics.baseline.completed],
    ["Year 2 Completed", metrics.y2.completed],
    ["Year 3 Completed", metrics.y3.completed],
    ["Year 4 Completed", metrics.y4.completed],
    ["Total Dropouts", metrics.dropouts.total],
    ["Recruitment Total", metrics.recruitment.grand_total ?? metrics.recruitment.summary_n],
    ["Y3 Visits This Month (MRI/BD/NP/CV)", `${metrics.y3Visits.mri}/${metrics.y3Visits.bloodDraw}/${metrics.y3Visits.np}/${metrics.y3Visits.cv}`],
    ["Y4 Visits This Month (BD/NP/CV)", `${metrics.y4Visits.bloodDraw}/${metrics.y4Visits.np}/${metrics.y4Visits.cv}`],
    ["Diabetics / Non-Diabetic", `${metrics.genotype.total_diabetics ?? "—"} / ${metrics.genotype.total_non_diabetic ?? "—"}`],
  ];

  el.innerHTML = `
    <table>
      <thead><tr><th>Metric</th><th>Value</th></tr></thead>
      <tbody>${rows.map(([k, v]) => `<tr><td>${k}</td><td><strong>${v ?? "—"}</strong></td></tr>`).join("")}</tbody>
    </table>`;
}

function renderReport() {
  const select = document.getElementById("report-month-select");
  const months = availableReportMonths();
  if (!select) return;

  select.innerHTML = months
    .map((m, i) => `<option value="${m.year}-${m.month}" ${i === 0 ? "selected" : ""}>${m.label}</option>`)
    .join("");

  const updatePreview = () => {
    const [year, month] = select.value.split("-").map(Number);
    renderReportPreview(collectReportMetrics(year, month));
  };

  select.onchange = updatePreview;
  updatePreview();

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
