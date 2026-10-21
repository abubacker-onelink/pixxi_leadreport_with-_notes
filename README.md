# Buy Leads Dashboard + Lead Notes Report

Two tools in one project. `npm start` opens a page with a switch between them:
- **Buy Leads Dashboard** (`/`) – described below.
- **Lead Notes Report** (`/notes`) – see the section further down.

## 1. Buy Leads Dashboard

Upload one Pixi CRM **Buy Leads** Excel export and get:

1. `leads_dashboard.html` – the interactive dashboard (single file, opens in any browser)
2. `Buy_Leads_Report_<period>.xlsx` – Excel report for active leads (live formulas)
3. `Buy_Leads_from_Lead_Pool_<period>.xlsx` – only when the export contains lead pool leads

## Setup (once)
```bash
npm install
```
Needs Node.js 18 or newer.

## Run
```bash
npm start
```
A browser page opens. Drop the Pixi export (.xlsx) on it and click **Generate reports**.
Everything is saved in `output/leads-report_<start>_to_<end>/`.

No browser? Use the command line:
```bash
node src/cli.js path/to/Buy-Leads_16-10-2026.xlsx
node src/cli.js export.xlsx --start "13-10-2026 00:00" --end "14-10-2026 09:40" --source "Meta Ads" --exclude-ids 1539972
```

## 2. Lead Notes Report
Set the Pixi API token in `PIXXI_TOKEN` before starting the app. The **Lead Notes Report** page takes one file: a Pixi leads Excel export (Buy or Portal). For every lead ID in the Excel file, the app fetches Note activity logs from Pixi's `POST /pixxiapi/v1/lead/activity-logs` endpoint and generates a separate report for that export.

For example, in macOS/Linux:
```bash
PIXXI_TOKEN="your-token" npm start
```

The token is sent to Pixi in the `X-PIXXI-TOKEN` header and is not stored in the workbook. The notes CSV remains supported by the command-line workflow below.

You get `Lead_Notes_Report_<Buy|Portal>_<date>.xlsx` with three tabs:
- **Leads with Notes** – one row per lead, Follow-up 1 = oldest note … last = newest. Bold date line then the note; "(by Name)" when the author is not the lead's agent. Grouped by created day, most notes first.
- **Leads without Notes** – agent merged once per group, biggest group first, oldest lead first.
- **Summary** – per agent: total leads, with notes, without notes, % with notes, total notes (formulas).

Note authors are Pixi user IDs mapped to names in `authors.json`. If unknown IDs appear, the page asks for their names, saves them to `authors.json` and re-runs.

CSV-based command line:
```bash
node src/notes-cli.js --notes notes.csv --leads leads.xlsx --title "Buy leads notes"
```

## Test with fake data
```bash
npm run sample        # writes sample/sample-export.xlsx
npm run sample:notes  # writes sample/sample-notes.csv and sample/sample-leads.xlsx
npm test              # both tools; full checks (formula recalculation needs LibreOffice, otherwise it is skipped)
```

## Defaults (same as the skill)
- All sources, lead pool included, Dubai time (UTC+4).
- Period = first Create Time day 00:00 to last day 23:59 (or "now" if the last day is today). Override under *Advanced* or with `--start/--end`.
- Test accounts and test lead IDs from `teams.json` are excluded. Leads with "test" in the name or note are flagged, not removed.
- Phone numbers and emails are never read, and number/email patterns typed into notes are masked.

## Files you edit
- `authors.json` – Pixi user ID to name for note authors.
- `teams.json` – teams, leaders, members, hot/lost stages, test accounts. Names must match Pixi's *Agent Name*.
- `assets/dashboard_template.html` – dashboard look and charts. The line `var D=/*DATA*/;` must stay.
- If Pixi renames a column, update `REQUIRED` in `src/readExport.js`.

## Notes
- Excel files recalculate when opened in Excel. Some quick-preview tools (email or phone previews) show empty cells until the file is opened in Excel or LibreOffice.
- **Template version:** this project bundles the template and `teams.json` from the skill. If your newest copies are in `Downloads/leads-report-out/skill-update`, copy them over `assets/dashboard_template.html` and `teams.json`.
- Lead notes: `First note` / `Last note` columns show the date and time of the oldest / newest note. Note times given as plain text are used as written (Dubai time); epoch numbers or ISO times ending in Z are treated as UTC and shifted +4h. Check one note against Pixi on your first run.
# pixxi_leadreport_with-_notes
