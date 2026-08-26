# Diabetes Brain Study (DBS) Operations Dashboard

An interactive dashboard that unifies data from **DBS Tracker.xlsx** and **Monthly Meeting Updates.xlsx** into a single, presentation-ready research operations view.

## What it shows

- **Overview** — enrollment funnel, active participants, and year-over-year completion
- **Study Progress** — baseline through year 4 metrics from both workbooks
- **Recruitment** — lead sources, totals, and monthly recruitment timeline
- **Participants** — searchable directory linking ID location, genotype, consent, labs, study partner, and dropout status
- **Visit Tracking** — completed visits by year, 2nd-year monthly completion, cumulative visit trends
- **Clinical Operations** — screening, study partners, stool samples, sample drop-off, MRI scheduling
- **Genotype & Labs** — diabetes/AD risk distribution and lab fulfillment
- **Dropouts** — baseline ineligible, year 1–4, and meeting dropout reasons
- **Data Explorer** — browse every parsed sheet from both Excel files

## Quick start

```bash
# 1. Create environment and install dependencies
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# 2. Build dashboard data from the Excel files
python scripts/build_dashboard_data.py

# 3. Serve the dashboard locally
python -m http.server 8080 --directory dashboard
```

Open [http://localhost:8080](http://localhost:8080) in your browser.

## Updating data

Whenever the Excel files change, rerun:

```bash
python scripts/build_dashboard_data.py
```

Then refresh the browser.

## Project structure

```
DBS_Study/
├── DBS Tracker.xlsx
├── Monthly Meeting Updates.xlsx
├── scripts/
│   └── build_dashboard_data.py   # ETL: Excel → dashboard/data.json
├── dashboard/
│   ├── index.html
│   ├── styles.css
│   ├── app.js
│   └── data.json                 # generated
└── requirements.txt
```

## Notes

- Participant IDs are normalized to `DBS-####` format so records link across both workbooks.
- The dashboard is static HTML/JS — no backend server required after data generation.
- All sheets from both workbooks are included in the Data Explorer section.
