# Google Sheets Auto-Sync Setup

The dashboard can pull directly from Google Sheets and refresh automatically once this one-time setup is complete.

## What Jimena needs to do

1. Share **both Google Sheets** with:
   - `ajagadis@usc.edu` (Editor or Viewer)
   - The dashboard service account email (Ashwarya will send this after step 2 below)

2. Send Ashwarya the **Google Sheets links** for:
   - DBS Tracker
   - Monthly Meeting Updates

## What Ashwarya needs to do (one time)

### 1. Create a Google Cloud service account

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create or select a project
3. Enable **Google Drive API**
4. Go to **IAM & Admin → Service Accounts → Create service account**
5. Create a key (**JSON**) and save it as:

```bash
credentials/google_service_account.json
```

6. Copy the service account email (looks like `dbs-dashboard@your-project.iam.gserviceaccount.com`) and send it to Jimena so she can share both sheets with it.

### 2. Configure sheet IDs

Copy the example config and paste the Google Sheet IDs or full URLs:

```bash
cp config/google_sheets.example.json config/google_sheets.json
```

Edit `config/google_sheets.json`:

```json
{
  "owner_email": "ajagadis@usc.edu",
  "tracker_spreadsheet_id": "https://docs.google.com/spreadsheets/d/YOUR_TRACKER_ID/edit",
  "meeting_spreadsheet_id": "https://docs.google.com/spreadsheets/d/YOUR_MEETING_ID/edit"
}
```

### 3. Test locally

```bash
source .venv/bin/activate
pip install -r requirements.txt
python scripts/sync_from_google_sheets.py
python scripts/build_dashboard_data.py
./publish_site.sh
```

### 4. Add GitHub secrets (for automatic daily updates)

In the repo **Settings → Secrets and variables → Actions**, add:

| Secret | Value |
|---|---|
| `GOOGLE_CREDENTIALS_JSON` | Full contents of `credentials/google_service_account.json` |
| `GOOGLE_TRACKER_SHEET_ID` | Tracker spreadsheet ID or URL |
| `GOOGLE_MEETING_SHEET_ID` | Meeting updates spreadsheet ID or URL |
| `DBS_SITE_DEPLOY_TOKEN` | GitHub personal access token with `repo` access (for publishing the live site) |

After secrets are added, the workflow **Sync Google Sheets and Publish Dashboard** will:
- Run daily at 6:00 AM Pacific
- Download the latest Google Sheets
- Rebuild the dashboard
- Publish to the live website

You can also trigger it manually from **Actions → Run workflow**.

## Contact

Dashboard owner: **ajagadis@usc.edu**

Live site: https://husseinyassinemd.github.io/dbs-study-dashboard-site/
