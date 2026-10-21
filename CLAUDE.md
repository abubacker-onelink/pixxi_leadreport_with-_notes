# Project notes for Claude Code

Node.js (CommonJS) tool: Pixi CRM Buy Leads export (.xlsx) -> dashboard HTML + Excel reports.

## Module API (use this to merge into another workflow)
```js
const { runReport } = require('./src/pipeline');
const res = await runReport('/path/export.xlsx', {
  start: '13-10-2026 00:00',   // optional, DD-MM-YYYY HH:MM, Dubai time
  end: '14-10-2026 09:40',     // optional
  export: undefined,           // optional, defaults to now (Dubai)
  source: '',                  // '' = all sources, 'Meta Ads' pre-selects Meta Ads
  excludeIds: [],              // extra lead IDs to exclude
  outRoot: './output',
  log: console.log,
});
// res = { qc, files: { dashboard, excel, excel_pool? }, outDir, period, source }
```

## Lead Notes Report (second tool)
```js
const { runNotesReport } = require('./src/notes/runNotesReport');
const r = await runNotesReport('notes.csv', 'leads.xlsx', { title, outRoot, authorsPath, newAuthors: { 9999: 'Name' }, log });
// r = { qc, file, outDir, kind }   qc has leads, with_notes, without_notes, total_notes, unknown_author_ids, lead_ids_not_found_in_leads_file, warnings
```
- `src/notes/readNotes.js` CSV + notes JSON + leads Excel readers (Buy / Portal / prepared leads.xlsx auto-detected).
- `src/notes/runNotesReport.js` builds the 3-tab workbook. Universe of leads = leads listed in the notes CSV.
- Rules: Follow-up 1 is the OLDEST note; never show user IDs (map through `authors.json`); Summary uses COUNTIF/SUMIF on "Leads with Notes" and without = total - with (merged cells break COUNTIF); no phone/email in text (`src/util.js` redact).
- Server: `/notes` page and `POST /api/notes` (fields: notes, leads, title, authors JSON).

## Layout (Buy Leads Dashboard)
- `src/readExport.js` reads only whitelisted columns (never phone/email). Splits Active vs Pool using the Status column.
- `src/prepareData.js` cleaning, QC, team mapping, builds the dashboard data object D (port of the skill's prepare_data.py).
- `src/excelReport.js` ExcelJS workbook with live formulas pointing at the `Lead Data` tab. Data bars on Daily Trend.
- `src/pipeline.js` orchestrates; `src/server.js` upload page (Express + multer, 127.0.0.1 only); `src/cli.js` command line.
- `assets/dashboard_template.html` must keep the `/*DATA*/` placeholder.

## Rules that must not change
- All times are Dubai wall-clock (UTC+4) stored as minutes since 1970 with no timezone conversion.
- Never output phone numbers or emails. Keep the column whitelist and the redact() safety net.
- Test accounts come from `teams.json` (`exclude_agents`, `exclude_ids`). Never ask about them.
- Team views count a lead for a team when Create By OR current Agent Name is on that team (Meta round robin: Create By = assigned at creation, Agent Name = holder now).
- Excel By Agent has "Leads" (holding now) and "Created by (assigned at creation)".
- Skip the pool workbook when the pool has 0 leads.

## Commands
`npm start` (both pages), `node src/cli.js file.xlsx`, `node src/notes-cli.js --notes n.csv --leads l.xlsx`, `npm run sample`, `npm run sample:notes`, `npm test`.
