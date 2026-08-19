# Carbon Measures Tracker

An independent dashboard tracking border carbon adjustments (BCAs) and domestic carbon
measures by jurisdiction — status, sector coverage, timeline, and official sources.
All content is maintained in a single spreadsheet (`BCA_tracker.xlsx`); the app reads it and renders the views.

**Independent · not affiliated with or endorsed by any government or organisation listed.**

## Files
| File | What it is |
|---|---|
| `BCA_tracker.xlsx` | The dataset. This is the only file you edit to update the tracker. |
| `app.py` | The Streamlit app that reads the spreadsheet. |
| `requirements.txt` | Python packages Streamlit Cloud installs automatically. |

## To update the tracker
Edit `BCA_tracker.xlsx`, replace the file in this repo (see steps below), and the live site
refreshes within a minute or two. You never touch the code.

## Data model (tabs in the spreadsheet)
- **Instruments** — one row per instrument (the hub). `instrument_id` links to the other tabs.
- **Sectors** — one row per instrument × sector (+ HS code). Pivots into the coverage matrix.
- **Events** — one row per instrument × dated event. Drives the timeline.
- **Sources** — one row per official document per instrument.
- **Status_legend / Category_legend / Sector_legend** — controlled vocabularies. Use these exact values.
- **Methodology** — attribution and sourcing note, shown on the Methodology tab.

Rows whose `category`/`status` are not valid legend values (or that are notes to self) are
skipped by the app and reported in the sidebar, so working notes can stay in the sheet.
