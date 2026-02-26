# Consulting Job Board Scraper

Automated weekly scraper for management consulting job boards. Generates a McKinsey-styled Excel file with jobs from multiple sources.

## Quick Start (Local)

```bash
pip install playwright openpyxl
playwright install chromium
python scrape_jobs.py
```

## GitHub Actions (Automated Weekly)

1. Push this repo to GitHub
2. Go to **Settings > Actions > General** and enable workflows
3. The scraper runs every Monday at 8am UTC automatically
4. Download the Excel from **Actions > Latest run > Artifacts**

### Optional: Email delivery

To get the Excel emailed to you weekly:

1. Go to **Settings > Secrets and variables > Actions**
2. Add these secrets:
   - `EMAIL_USERNAME` - your Gmail address
   - `EMAIL_PASSWORD` - a Gmail App Password
   - `EMAIL_TO` - where to send the report
3. Add a repository variable: `SEND_EMAIL` = `true`

### Manual trigger

Click **Actions > Weekly Consulting Job Scrape > Run workflow** to run on demand.
