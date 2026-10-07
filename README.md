# Nymi Inc — CFO Financial Dashboard (September 2026 refresh)

GitHub Pages-ready static CFO dashboard. The dashboard uses `data/financials.json` and also contains an embedded fallback snapshot so it can still render if the JSON request is temporarily unavailable.

## Upload to GitHub

Upload the contents of this folder to the root of your repository, preserving:

- `index.html`
- `data/financials.json`
- `scripts/extract_data.py`
- `requirements.txt`
- `.github/workflows/deploy.yml`

Do **not** upload the raw XLSB workbook to GitHub.

## Refresh next month

Place the new XLSB workbook in the repository root locally and run:

```bash
pip install -r requirements.txt
python scripts/extract_data.py "Statement Of Financial MIS <month>.xlsb"
```

Commit the refreshed `data/financials.json` and push to `main`. GitHub Pages will redeploy through the included workflow.

## Current source

`Statement Of Financial MIS Sept'26 Corp.xlsb`

The September workbook contains `Summary SOP` and `Detailed SOP`. It does not contain a Balance Sheet worksheet, so the dashboard does not fabricate new Balance Sheet/Cash Flow figures from this workbook.
