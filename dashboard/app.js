const DATASETS = {
  "DBS Tracker > ID Location": "id_location",
  "DBS Tracker > Completed Visits (2nd Year)": "completed_visits.second_year",
  "DBS Tracker > Completed Visits (3rd Year)": "completed_visits.third_year",
  "DBS Tracker > Completed Visits (4th Year)": "completed_visits.fourth_year",
  "DBS Tracker > Completed Visits (Completed)": "completed_visits.completed",
  "DBS Tracker > Screening Visits": "screening",
  "DBS Tracker > Genotype": "genotype",
  "DBS Tracker > Study Partner": "study_partner",
  "DBS Tracker > Stool Samples": "stool_samples.participants",
  "DBS Tracker > Labs & MRNs": "labs",
  "DBS Tracker > ClinCard": "clincard",
  "DBS Tracker > Baseline Scheduling": "baseline_scheduling",
  "DBS Tracker > 2nd Year Scheduling": "second_year_scheduling",
  "DBS Tracker > 3rd Year Scheduling": "third_year_scheduling",
  "DBS Tracker > 4th Year Scheduling": "fourth_year_scheduling",
  "DBS Tracker > Dropouts (Ineligible)": "dropouts.baseline_ineligible",
  "DBS Tracker > Dropouts (Y1-4)": "dropouts.year_1_to_4",
  "DBS Tracker > Consent to Share": "consent",
  "DBS Tracker > Lab Results Requests": "lab_results_requests",
  "DBS Tracker > MRI CD Requests": "mri_cd_requests",
  "DBS Tracker > Sample Drop Off": "sample_drop_off",
  "DBS Tracker > Synopsis Forms": "synopsis_forms",
  "DBS Tracker > MRI Dates Available": "mri_dates_available",
  "Monthly Meeting > Pt Status": "pt_status",
  "Monthly Meeting > DropOuts": "meeting_dropouts",
  "Monthly Meeting > 2nd Year Monthly": "second_year_monthly",
  "Monthly Meeting > Visit Trends (Sheet5)": "visit_trends_sheet5",
  "Monthly Meeting > Visit Trends (Sheet6)": "visit_trends_sheet6",
  "Monthly Meeting > Enrolled Mapping": "enrolled_mapping",
  "DBS Tracker > NP Results Requests": "np_results_requests",
  "DBS Tracker > uds_id_200": "uds_id_200",
  "Monthly Meeting > 16 DQ pts": "dq_pts",
  "Monthly Meeting > MRI Outcomes (Sheet 9)": "mri_outcomes",
};

let DATA = null;
const charts = {};

const sectionCopy = {
  overview: ["Overview", "Unified metrics across DBS Tracker and Monthly Meeting Updates"],
  progress: ["Study Progress", "Longitudinal completion across baseline through year 4"],
  recruitment: ["Recruitment", "Lead sources and enrollment funnel performance"],
  participants: ["Participants", "Searchable directory with linked operational records"],
  visits: ["Visit Tracking", "Scheduling, completion, and monthly visit trends"],
  clinical: ["Clinical Operations", "Screening, study partners, samples, and MRI logistics"],
  genotype: ["Genotype & Labs", "Diabetes status, AD risk groups, and lab fulfillment"],
  dropouts: ["Dropouts", "Attrition analysis across baseline and follow-up years"],
  "monthly-report": ["Monthly Report", "One-click USC-branded PowerPoint for monthly progress meetings"],
  "data-explorer": ["Data Explorer", "Browse every sheet loaded from both workbooks"],
};

function getPath(obj, path) {
  return path.split(".").reduce((acc, key) => (acc ? acc[key] : undefined), obj);
}

function getDataset(path, type = "structured") {
  if (type === "raw") {
    return DATA.raw_sheets?.[path] || [];
  }
  return getPath(DATA, path) || [];
}

function tableColumns(rows) {
  if (!rows?.length) return [];
  const keys = new Set();
  rows.forEach((row) => Object.keys(row).forEach((k) => keys.add(k)));
  const ordered = ["participant_id", "ID", ...Array.from(keys).filter((k) => k !== "participant_id" && k !== "ID")];
  return ordered.filter((k, i, arr) => arr.indexOf(k) === i);
}

function pct(value) {
  if (value == null) return "—";
  const num = Number(value);
  return Number.isFinite(num) ? `${Math.round(num * 1000) / 10}%` : value;
}

function fmtDate(value) {
  if (!value) return "—";
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? value : d.toLocaleDateString();
}

function renderTable(containerId, rows, columns, options = {}) {
  const el = document.getElementById(containerId);
  if (!rows?.length) {
    el.innerHTML = `<p class="about-copy">No records available.</p>`;
    return;
  }
  const limited = options.limit ? rows.slice(0, options.limit) : rows;
  const headers = columns || tableColumns(limited);
  el.innerHTML = `
    <table>
      <thead><tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr></thead>
      <tbody>
        ${limited
          .map((row, idx) => {
            const cells = headers
              .map((h) => {
                const val = row[h];
                const display =
                  val == null || val === ""
                    ? "—"
                    : typeof val === "object"
                      ? JSON.stringify(val)
                      : String(val);
                return `<td>${display}</td>`;
              })
              .join("");
            const clickable = options.onRowClick ? " clickable-row" : "";
            return `<tr class="${clickable}" data-index="${idx}">${cells}</tr>`;
          })
          .join("")}
      </tbody>
    </table>`;

  if (options.onRowClick) {
    el.querySelectorAll("tbody tr").forEach((tr) => {
      tr.addEventListener("click", () => options.onRowClick(limited[Number(tr.dataset.index)]));
    });
  }
}

function destroyChart(id) {
  if (charts[id]) {
    charts[id].destroy();
    delete charts[id];
  }
}

function chartDefaults() {
  Chart.defaults.font.family = '"DM Sans", system-ui, sans-serif';
  Chart.defaults.color = "#5f6b7a";
}

function getMeetingProgress() {
  return DATA.study_progress_meeting?.length ? DATA.study_progress_meeting : DATA.study_progress_tracker;
}

function renderOverview() {
  const progress = getMeetingProgress();
  const baseline = progress.find((r) => String(r.year).toLowerCase() === "baseline") || progress[0];
  const year2 = progress.find((r) => String(r.year) === "2") || {};
  const year3 = progress.find((r) => String(r.year) === "3") || {};
  const active = DATA.active_participants.length;
  const dropouts = DATA.meeting_dropouts.length + DATA.dropouts.baseline_ineligible.length;

  document.getElementById("kpi-grid").innerHTML = [
    {
      label: "Enrolled Participants",
      value: baseline.enrolled ?? 200,
      footnote: "Target cohort size",
      cls: "accent",
    },
    {
      label: "Active Participants",
      value: active,
      footnote: "From All active IDs sheet",
      cls: "teal",
    },
    {
      label: "Baseline Completed",
      value: baseline.completed ?? "—",
      footnote: pct(baseline.pct_completed_total) + " of enrolled",
      cls: "",
    },
    {
      label: "Year 3 Completed",
      value: year3.completed ?? "—",
      footnote: pct(year3.pct_completed_active) + " of active · Monthly Meeting Updates",
      cls: "",
    },
  ]
    .map(
      (k) => `
      <div class="kpi">
        <div class="kpi-label">${k.label}</div>
        <div class="kpi-value ${k.cls}">${k.value}</div>
        <div class="kpi-footnote">${k.footnote}</div>
      </div>`
    )
    .join("");

  const recruitmentLeads = DATA.recruitment.grand_total || 1735;
  const funnelLeadsEl = document.getElementById("funnel-leads-count");
  if (funnelLeadsEl) {
    funnelLeadsEl.textContent = recruitmentLeads.toLocaleString();
    funnelLeadsEl.title = "Recruitment leads";
  }

  destroyChart("funnel-chart");
  charts["funnel-chart"] = new Chart(document.getElementById("funnel-chart"), {
    type: "bar",
    data: {
      labels: ["Enrolled", "Active", "Year 2 Complete", "Year 3 Complete"],
      datasets: [
        {
          label: "Count",
          data: [
            baseline.enrolled || 200,
            active,
            year2.completed || 0,
            year3.completed || 0,
          ],
          backgroundColor: ["#c45c26", "#1f6f78", "#3d8b93", "#7eb4ba"],
          borderRadius: 10,
          barThickness: 52,
        },
      ],
    },
    options: {
      plugins: { legend: { display: false } },
      scales: {
        y: {
          beginAtZero: true,
          suggestedMax: 220,
          ticks: { stepSize: 50 },
        },
      },
    },
  });

  const years = progress.filter((r) => r.completed != null).map((r) => `Year ${r.year}`);
  const completed = progress.filter((r) => r.completed != null).map((r) => r.completed);
  destroyChart("progress-chart");
  charts["progress-chart"] = new Chart(document.getElementById("progress-chart"), {
    type: "line",
    data: {
      labels: years,
      datasets: [
        {
          label: "Completed",
          data: completed,
          borderColor: "#1f6f78",
          backgroundColor: "rgba(31,111,120,0.15)",
          fill: true,
          tension: 0.35,
        },
      ],
    },
    options: { plugins: { legend: { display: false } } },
  });
}

function renderProgress() {
  const progress = getMeetingProgress();
  const tracker = DATA.study_progress_tracker || [];
  const labels = progress.map((r) => r.year);
  destroyChart("completion-total-chart");
  charts["completion-total-chart"] = new Chart(document.getElementById("completion-total-chart"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "% Completed (Total)",
          data: progress.map((r) => (r.pct_completed_total != null ? r.pct_completed_total * 100 : null)),
          backgroundColor: "#c45c26",
          borderRadius: 8,
        },
      ],
    },
    options: { scales: { y: { max: 100, ticks: { callback: (v) => `${v}%` } } } },
  });

  destroyChart("completion-active-chart");
  charts["completion-active-chart"] = new Chart(document.getElementById("completion-active-chart"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "% Completed (Active)",
          data: progress.map((r) => (r.pct_completed_active != null ? r.pct_completed_active * 100 : null)),
          backgroundColor: "#1f6f78",
          borderRadius: 8,
        },
      ],
    },
    options: { scales: { y: { max: 100, ticks: { callback: (v) => `${v}%` } } } },
  });

  const rows = progress.map((r) => ({
    Year: r.year,
    Enrolled: r.enrolled,
    "Dropped / DQ": r.dropped_out_dq,
    Completed: r.completed,
    Remaining: r.remaining,
    "% Total": pct(r.pct_completed_total),
    "% Active": pct(r.pct_completed_active),
    Notes: r.notes || "—",
    Source: "Monthly Meeting Updates",
  }));

  renderTable("progress-table", rows);

  const compareEl = document.getElementById("progress-compare-table");
  if (compareEl) {
    const compareRows = progress
      .filter((r) => r.completed != null)
      .map((meetingRow) => {
        const trackerRow = tracker.find((t) => String(t.year) === String(meetingRow.year)) || {};
        const meetingCompleted = meetingRow.completed;
        const trackerCompleted = trackerRow.completed;
        const match = meetingCompleted === trackerCompleted ? "Match" : "Mismatch";
        return {
          Year: meetingRow.year,
          "Meeting Completed": meetingCompleted,
          "Tracker Completed": trackerCompleted ?? "—",
          Difference: trackerCompleted != null ? meetingCompleted - trackerCompleted : "—",
          Status: match,
        };
      });
    renderTable("progress-compare-table", compareRows);
  }
}

function renderRecruitment() {
  const sources = DATA.recruitment.sources.filter((s) => s.total_leads != null);
  destroyChart("recruitment-pie");
  charts["recruitment-pie"] = new Chart(document.getElementById("recruitment-pie"), {
    type: "doughnut",
    data: {
      labels: sources.map((s) => s.source),
      datasets: [
        {
          data: sources.map((s) => s.total_leads),
          backgroundColor: ["#c45c26", "#1f6f78", "#7eb4ba", "#d8a48a", "#132033", "#8aa0b5"],
        },
      ],
    },
  });

  const topSources = sources.slice(0, 3);
  const periods = DATA.recruitment.periods.slice(0, 12);
  destroyChart("recruitment-timeline");
  charts["recruitment-timeline"] = new Chart(document.getElementById("recruitment-timeline"), {
    type: "line",
    data: {
      labels: periods.map((p) => p.slice(0, 28)),
      datasets: topSources.map((source, idx) => ({
        label: source.source,
        data: periods.map((p) => source.timeline[p] || 0),
        borderColor: ["#c45c26", "#1f6f78", "#132033"][idx],
        tension: 0.25,
      })),
    },
    options: { scales: { x: { ticks: { maxRotation: 45, minRotation: 45 } } } },
  });

  renderTable(
    "recruitment-table",
    sources.map((s) => ({ Source: s.source, "Total Leads": s.total_leads })),
    ["Source", "Total Leads"]
  );
}

function renderParticipants(filter = "all", search = "") {
  let rows = DATA.participants;
  if (filter !== "all") rows = rows.filter((p) => p.status === filter);
  if (search) {
    const q = search.toUpperCase();
    rows = rows.filter((p) => p.participant_id.includes(q));
  }

  renderTable(
    "participants-table",
    rows.map((p) => ({
      ID: p.participant_id,
      Status: p.status || "—",
      Cabinet: p.cabinet_location || "—",
      Consent: p.consent || "—",
      Genotype: p.genotype?.genotype ? fmtDate(p.genotype.genotype) : "—",
      HbA1c: p.genotype?.hba1c ?? "—",
      Dropout: p.pt_status?.reason_dropout || p.meeting_dropout_reason || "—",
    })),
    ["ID", "Status", "Cabinet", "Consent", "Genotype", "HbA1c", "Dropout"],
    {
      onRowClick: (row) => {
        const participant = DATA.participants.find((p) => p.participant_id === row.ID);
        document.getElementById("participant-detail-card").hidden = false;
        document.getElementById("participant-detail-title").textContent = row.ID;
        document.getElementById("participant-detail").textContent = JSON.stringify(participant, null, 2);
      },
    }
  );
}

function renderVisits() {
  const monthly = DATA.second_year_monthly.find((r) => r.visit_type === "Clinician Visit") || {};
  const npMonthly = DATA.second_year_monthly.find((r) => r.visit_type === "NP") || {};
  const monthKeys = Object.keys(monthly).filter((k) => !["visit_type", "total_completed"].includes(k));

  destroyChart("second-year-chart");
  charts["second-year-chart"] = new Chart(document.getElementById("second-year-chart"), {
    type: "bar",
    data: {
      labels: monthKeys,
      datasets: [
        { label: "Clinician Visit", data: monthKeys.map((k) => monthly[k] || 0), backgroundColor: "#1f6f78" },
        { label: "NP", data: monthKeys.map((k) => npMonthly[k] || 0), backgroundColor: "#c45c26" },
      ],
    },
    options: { scales: { x: { ticks: { maxRotation: 45, minRotation: 45 } } } },
  });

  const trends = DATA.visit_trends_sheet5;
  destroyChart("visit-trend-chart");
  charts["visit-trend-chart"] = new Chart(document.getElementById("visit-trend-chart"), {
    type: "line",
    data: {
      labels: trends.map((r) => fmtDate(r.Date)),
      datasets: [
        { label: "CV Cumulative", data: trends.map((r) => r["CV Cumulative"]), borderColor: "#1f6f78" },
        { label: "NP Cumulative", data: trends.map((r) => r["NP Cumulative"]), borderColor: "#c45c26" },
        { label: "Max Possible", data: trends.map((r) => r["Max Possible Visits"]), borderColor: "#d8dee8" },
      ],
    },
  });

  renderTable(
    "completed-second",
    DATA.completed_visits.second_year.map((r) => ({
      ID: r.participant_id,
      Month: fmtDate(r.scheduling_month),
      Assigned: r.assigned_to || "—",
      Notes: r.notes || "—",
    }))
  );
  renderTable(
    "completed-third",
    DATA.completed_visits.third_year.map((r) => ({
      ID: r.participant_id,
      Month: fmtDate(r.scheduling_month),
      Notes: r.notes || "—",
    }))
  );
  renderTable(
    "completed-fourth",
    DATA.completed_visits.fourth_year.map((r) => ({
      ID: r.participant_id,
      Month: fmtDate(r.scheduling_month),
      Notes: r.notes || "—",
    }))
  );
  renderTable("baseline-scheduling-table", DATA.baseline_scheduling);
  renderTable("second-year-scheduling-table", DATA.second_year_scheduling);
  renderTable("third-year-scheduling-table", DATA.third_year_scheduling);
  renderTable("fourth-year-scheduling-table", DATA.fourth_year_scheduling);
}

function renderClinical() {
  const partnerComplete = DATA.study_partner.filter((r) => r.completed === 1 || r.completed === "1").length;
  const stoolComplete = DATA.stool_samples.participants.filter((r) =>
    String(r.status || "").toLowerCase().includes("complete")
  ).length;

  document.getElementById("clinical-kpis").innerHTML = [
    { label: "Screening Records", value: DATA.screening.length },
    { label: "Study Partners Completed", value: partnerComplete },
    { label: "Stool Samples Complete", value: stoolComplete },
    { label: "ClinCard Records", value: DATA.clincard.length },
  ]
    .map(
      (k) => `
      <div class="mini-kpi">
        <div class="mini-kpi-label">${k.label}</div>
        <div class="mini-kpi-value">${k.value}</div>
      </div>`
    )
    .join("");

  renderTable(
    "screening-table",
    DATA.screening.map((r) => ({
      ID: r.participant_id,
      "Screening Date": fmtDate(r.screening_date),
      ICF: r.icf_signed,
      "Blood Draw": r.blood_draw_completed,
      Notes: r.notes || "—",
    }))
  );
  renderTable(
    "study-partner-table",
    DATA.study_partner.map((r) => ({
      ID: r.participant_id,
      Completed: r.completed ?? "—",
      Sent: r.survey_link_sent || "—",
      Notes: r.notes || "—",
    }))
  );
  renderTable("stool-table", DATA.stool_samples.participants);
  renderTable("sample-dropoff-table", DATA.sample_drop_off);
  renderTable("mri-dates-table", DATA.mri_dates_available);
  renderTable("clincard-table", DATA.clincard);
  renderTable("np-results-table", DATA.np_results_requests);
  renderTable("mri-outcomes-table", DATA.mri_outcomes);
  renderTable("dq-pts-table", DATA.dq_pts);
}

function renderGenotype() {
  const summary = DATA.genotype_summary;
  document.getElementById("genotype-kpis").innerHTML = [
    { label: "Diabetics", value: summary.total_diabetics ?? "—" },
    { label: "Non-Diabetic", value: summary.total_non_diabetic ?? "—" },
    { label: "High AD Risk", value: summary.total_high_ad_risk ?? "—" },
    { label: "Low AD Risk", value: summary.total_low_ad_risk ?? "—" },
  ]
    .map(
      (k) => `
      <div class="mini-kpi">
        <div class="mini-kpi-label">${k.label}</div>
        <div class="mini-kpi-value">${k.value}</div>
      </div>`
    )
    .join("");

  destroyChart("diabetes-chart");
  charts["diabetes-chart"] = new Chart(document.getElementById("diabetes-chart"), {
    type: "doughnut",
    data: {
      labels: ["Diabetic", "Non-Diabetic"],
      datasets: [
        {
          data: [summary.total_diabetics || 0, summary.total_non_diabetic || 0],
          backgroundColor: ["#c45c26", "#1f6f78"],
        },
      ],
    },
  });

  destroyChart("ad-risk-chart");
  charts["ad-risk-chart"] = new Chart(document.getElementById("ad-risk-chart"), {
    type: "doughnut",
    data: {
      labels: ["High AD Risk", "Low AD Risk"],
      datasets: [
        {
          data: [summary.total_high_ad_risk || 0, summary.total_low_ad_risk || 0],
          backgroundColor: ["#132033", "#7eb4ba"],
        },
      ],
    },
  });

  renderTable(
    "genotype-table",
    DATA.genotype.map((r) => ({
      ID: r.participant_id,
      HbA1c: r.hba1c,
      pTau: r.ptau,
      Diabetic: r.diabetic,
      "High AD Risk": r.high_ad_risk,
      "Low AD Risk": r.low_ad_risk,
    }))
  );
  renderTable(
    "labs-table",
    DATA.labs.map((r) => ({
      ID: r.participant_id,
      Year: r.labs_year,
      MRN: r.mrn,
      "Lab Date": r.lab_date,
      Emailed: r.emailed_texted,
    }))
  );
}

function renderDropouts() {
  const dropoutDates = DATA.pt_status.filter((r) => r.dropout_date);
  const byYear = {};
  dropoutDates.forEach((r) => {
    const year = r.dropout_year || "Unknown";
    byYear[year] = (byYear[year] || 0) + 1;
  });

  destroyChart("dropout-chart");
  charts["dropout-chart"] = new Chart(document.getElementById("dropout-chart"), {
    type: "bar",
    data: {
      labels: Object.keys(byYear),
      datasets: [{ label: "Dropouts", data: Object.values(byYear), backgroundColor: "#c45c26", borderRadius: 8 }],
    },
    options: { plugins: { legend: { display: false } } },
  });

  renderTable(
    "meeting-dropouts-table",
    DATA.meeting_dropouts.map((r) => ({ ID: r.participant_id, Reason: r.reason || "—" }))
  );
  renderTable("baseline-dropouts-table", DATA.dropouts.baseline_ineligible);
  renderTable("yearly-dropouts-table", DATA.dropouts.year_1_to_4);
}

function renderCatalog() {
  const grid = document.getElementById("catalog-grid");
  const cards = [];

  Object.entries(DATASETS).forEach(([label, path]) => {
    const rows = getDataset(path);
    const count = Array.isArray(rows) ? rows.length : 0;
    cards.push({ label, path, type: "structured", count });
  });

  (DATA.sheet_catalog || []).forEach((item) => {
    const prefix = item.workbook === "DBS Tracker" ? "tracker" : "meeting";
    const rawKey = `${prefix}::${item.sheet.trim()}`;
    const label = `${item.workbook} > ${item.sheet.trim()} (Complete Raw)`;
    const rows = getDataset(rawKey, "raw");
    cards.push({ label, path: rawKey, type: "raw", count: rows.length });
  });

  grid.innerHTML = cards
    .map(
      (card) => `
      <div class="catalog-card" data-path="${card.path}" data-type="${card.type}">
        <h4>${card.label.split(" > ").slice(-1)[0]}</h4>
        <p>${card.label.includes(" > ") ? card.label.split(" > ")[0] : "Dataset"} · ${card.count} rows</p>
      </div>`
    )
    .join("");

  grid.querySelectorAll(".catalog-card").forEach((card) => {
    card.addEventListener("click", () => {
      grid.querySelectorAll(".catalog-card").forEach((c) => c.classList.remove("active"));
      card.classList.add("active");
      const rows = getDataset(card.dataset.path, card.dataset.type);
      document.getElementById("explorer-title").textContent = card.querySelector("h4").textContent;
      renderTable("explorer-table", Array.isArray(rows) ? rows : [rows]);
    });
  });
}

function activateSection(section) {
  document.querySelectorAll(".nav-item").forEach((btn) => btn.classList.toggle("active", btn.dataset.section === section));
  document.querySelectorAll(".panel").forEach((panel) => panel.classList.toggle("active", panel.id === section));
  const [title, subtitle] = sectionCopy[section];
  document.getElementById("section-title").textContent = title;
  document.getElementById("section-subtitle").textContent = subtitle;
  if (section === "monthly-report" && typeof updateReportPreview === "function") {
    updateReportPreview();
  }
}

function initNavigation() {
  document.querySelectorAll(".nav-item").forEach((btn) => {
    btn.addEventListener("click", () => activateSection(btn.dataset.section));
  });

  document.getElementById("status-filter").addEventListener("change", (e) => {
    renderParticipants(e.target.value, document.getElementById("global-search").value.trim());
  });

  document.getElementById("global-search").addEventListener("input", (e) => {
    const q = e.target.value.trim();
    if (q) {
      activateSection("participants");
      renderParticipants(document.getElementById("status-filter").value, q);
    }
  });
}

async function boot() {
  chartDefaults();
  const response = await fetch("data.json");
  DATA = await response.json();
  document.getElementById("generated-at").textContent = `Updated ${new Date(DATA.meta.generated_at).toLocaleString()}`;

  renderOverview();
  renderProgress();
  renderRecruitment();
  renderParticipants();
  renderVisits();
  renderClinical();
  renderGenotype();
  renderDropouts();
  renderReport();
  renderCatalog();
  initNavigation();
}

boot().catch((err) => {
  document.body.innerHTML = `<div style="padding:40px;font-family:sans-serif"><h2>Dashboard failed to load</h2><p>${err.message}</p><p>Run <code>python scripts/build_dashboard_data.py</code> then serve the dashboard folder.</p></div>`;
});
