# Ticket Report Automation

This repository contains a simple Python workflow to compute ticket resolution KPIs from
an exported Excel file.

## Quick start

1. Place your export file in `data/` (or point to its location directly).
2. Run the script:

```bash
python src/reporting.py data/ExportData.xlsx --year 2025
```

### Common options

Generate a year-to-date report (last 365 days):

```bash
python src/reporting.py data/ExportData.xlsx --ytd
```

Filter the report to a specific responsible alias (e.g., `BIN`):

```bash
python src/reporting.py data/ExportData.xlsx --year 2025 --alias BIN
```

Write the HTML report to a custom location:

```bash
python src/reporting.py data/ExportData.xlsx --output-report reports/tech_support.html
```

## What it does

- Loads the ticket export from Excel.
- Parses the `Created` and `Closed` timestamps.
- Calculates raw resolution time in hours.
- Calculates business-hour resolution time, excluding weekends and defined shutdowns.
- Builds a one-page HTML report with KPIs, charts, and top-10 breakdowns.
- Supports optional YTD and alias filtering to narrow the report window and owner.

## Customizing shutdowns

Edit `DEFAULT_SHUTDOWNS` in `src/reporting.py` to match your organization’s shutdown
calendar.

## Notes on date parsing

The export uses month/day/year timestamps with AM/PM (example: `2/3/2026 9:05:02 AM`).
The script parses this format first and falls back to pandas auto-detection for any
unexpected variants.