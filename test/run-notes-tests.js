'use strict';
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync, spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const { runNotesReport } = require('../src/notes/runNotesReport');
const { readNotesCsv } = require('../src/notes/readNotes');
const { fetchPixiNotes } = require('../src/notes/fetchPixiNotes');
const D = require('../src/dates');

let failed = 0;
const ok = (c, m) => { console.log((c ? 'PASS ' : 'FAIL ') + m); if (!c) failed++; };
const val = (c) => { const v = c.value; return v && typeof v === 'object' && 'result' in v ? v.result : v; };

(async () => {
  const requests = [];
  const apiPages = [
    { statusCode: 0, data: { currentPage: 1, totalPages: 2, list: [{ message: 'First note', createdBy: 9509, createTime: '2026-10-18 10:00:00' }] } },
    { statusCode: 0, data: { currentPage: 2, totalPages: 2, list: [{ noteContent: 'Second note', createdById: 9528, createdAt: '2026-10-18 11:00:00' }] } },
  ];
  const apiNotes = await fetchPixiNotes([12345], {
    token: 'test-token',
    endpoint: 'https://pixi.test/activity-logs',
    fetchImpl: async (url, options) => {
      requests.push({ url, options });
      return { ok: true, status: 200, json: async () => apiPages[requests.length - 1] };
    },
  });
  ok(apiNotes.leads.get(12345).notes.length === 2, 'Pixi notes pagination combines every page');
  ok(requests.length === 2 && requests[0].options.headers['X-PIXXI-TOKEN'] === 'test-token', 'Pixi API uses the token header');
  const reqBody = JSON.parse(requests[0].options.body);
  ok(reqBody.leadId === 12345 && reqBody.type === 'Note' && reqBody.page === 1, 'Pixi API sends the lead ID and note filter');
  const success200 = await fetchPixiNotes([12345], {
    token: 'test-token',
    fetchImpl: async () => ({
      ok: true, status: 200,
      json: async () => ({ statusCode: 200, message: 'success', data: { currentPage: 1, totalPages: 1, list: [] } }),
    }),
  });
  ok(success200.leads.has(12345), 'Pixi statusCode 200 with success message is accepted');

  let apiError = '';
  try {
    await fetchPixiNotes([12345], {
      token: 'test-token',
      fetchImpl: async () => ({ ok: false, status: 401, json: async () => ({ message: 'unauthorized' }) }),
    });
  } catch (e) { apiError = e.message; }
  ok(apiError.includes('HTTP 401'), 'Pixi API errors are surfaced with the response status');

  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'notes-test-'));
  execFileSync('node', [path.join(__dirname, '..', 'sample', 'make-notes-sample.js'), tmp], { stdio: 'ignore' });
  const csv = path.join(tmp, 'sample-notes.csv'), xlsx = path.join(tmp, 'sample-leads.xlsx');
  const authorsCopy = path.join(tmp, 'authors.json');
  fs.copyFileSync(path.join(__dirname, '..', 'authors.json'), authorsCopy);
  const res = await runNotesReport(csv, xlsx, { outRoot: path.join(tmp, 'out'), authorsPath: authorsCopy });
  const q = res.qc;

  ok(q.leads === q.with_notes + q.without_notes, `leads = with + without (${q.leads} = ${q.with_notes} + ${q.without_notes})`);
  ok(q.lead_ids_not_found_in_leads_file.length === 1 && q.lead_ids_not_found_in_leads_file[0] === 1559999, 'lead missing from leads file reported');
  ok(q.unknown_author_ids.length === 1 && q.unknown_author_ids[0] === '9999', 'unknown author ID reported');

  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(res.file);
  ok(wb.worksheets.map((w) => w.name).join('|') === 'Leads with Notes|Leads without Notes|Summary', 'three sheets in skill order');
  const ws = wb.getWorksheet('Leads with Notes');

  // independent check: follow-up 1 is the OLDEST note, last follow-up the NEWEST
  const { leads: raw } = readNotesCsv(csv);
  let checked = 0, orderOk = true, countOk = true;
  for (let r = 2; r <= ws.rowCount; r++) {
    const id = ws.getCell(r, 1).value; if (!id) continue;
    const times = raw.get(id).notes.map((n) => n.time).filter((t) => t !== null).sort((a, b) => a - b);
    const count = ws.getCell(r, 6).value;
    if (count !== raw.get(id).notes.length) countOk = false;
    const first = ws.getCell(r, 9).value.richText[0].text.trim();
    const lastCell = ws.getCell(r, 8 + count).value.richText[0].text.trim();
    if (first !== D.fmtDow(times[0]) || lastCell !== D.fmtDow(times[times.length - 1])) orderOk = false;
    checked++;
  }
  ok(checked === q.with_notes, `checked ${checked} leads with notes`);
  ok(orderOk, 'Follow-up 1 = oldest note, last follow-up = newest note');
  ok(countOk, 'Notes (count) column matches the JSON');

  // sort: created day ascending, most notes first inside a day
  let sortOk = true, prev = null;
  for (let r = 2; r <= ws.rowCount; r++) {
    const c = ws.getCell(r, 3).value; if (!(c instanceof Date)) continue;
    const day = Math.floor(c.getTime() / 86400000), n = ws.getCell(r, 6).value;
    if (prev && (day < prev.day || (day === prev.day && n > prev.n))) sortOk = false;
    prev = { day, n };
  }
  ok(sortOk, 'rows sorted by created day, then most notes first');

  // no user IDs and no phone/email in any text
  let idLeak = 0, phoneLeak = 0;
  const walk = (v) => {
    const t = typeof v === 'string' ? v : v && v.richText ? v.richText.map((x) => x.text).join('') : '';
    if (/\b(95\d\d|9821|9473|9999)\b/.test(t)) idLeak++;
    if (/\+971\s?\d|@example\.com|a@b\.com/.test(t)) phoneLeak++;
  };
  wb.eachSheet((s) => s.eachRow((row) => row.eachCell((c) => walk(c.value))));
  ok(idLeak === 0, 'no Pixi user IDs shown in the report');
  ok(phoneLeak === 0, 'no phone numbers or emails in the report');

  // "(by Name)" appears only when the author differs from the agent
  let byOk = true, byCount = 0;
  for (let r = 2; r <= ws.rowCount; r++) {
    const agent = ws.getCell(r, 4).value; if (!agent) continue;
    for (let c = 9; c <= ws.columnCount; c++) {
      const v = ws.getCell(r, c).value; if (!v || !v.richText) continue;
      const m = /\(by ([^)]+)\)$/.exec(v.richText[1].text);
      if (m) { byCount++; if (m[1].toLowerCase() === String(agent).toLowerCase()) byOk = false; }
    }
  }
  ok(byOk && byCount > 0, `"(by Name)" only for other authors (${byCount} found)`);

  const wsW = wb.getWorksheet('Leads without Notes');
  ok(wsW.model.merges && wsW.model.merges.length > 0, 'agent cells merged in Leads without Notes');
  let wcount = 0; for (let r = 2; r <= wsW.rowCount; r++) if (wsW.getCell(r, 2).value) wcount++;
  ok(wcount === q.without_notes, 'every lead without notes is listed');

  // new author names get saved and used
  const res2 = await runNotesReport(csv, xlsx, { outRoot: path.join(tmp, 'out2'), authorsPath: authorsCopy, newAuthors: { 9999: 'Test Person' } });
  ok(res2.qc.unknown_author_ids.length === 0, 'unknown author resolved after adding a name');

  if (spawnSync('soffice', ['--version']).error) console.log('SKIP Summary formula check (LibreOffice not installed).');
  else {
    const out = path.join(tmp, 'calc');
    spawnSync('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', out, res.file], { stdio: 'ignore' });
    const c = new ExcelJS.Workbook(); await c.xlsx.readFile(path.join(out, path.basename(res.file)));
    const s = c.getWorksheet('Summary');
    let t = 5; while (s.getCell(t, 1).value !== 'TOTAL') t++;
    ok(val(s.getCell(t, 2)) === q.leads, `Summary total leads = ${q.leads}`);
    ok(val(s.getCell(t, 3)) === q.with_notes, `Summary with notes = ${q.with_notes}`);
    ok(val(s.getCell(t, 4)) === q.without_notes, `Summary without notes = ${q.without_notes}`);
    ok(val(s.getCell(t, 6)) === q.total_notes, `Summary total notes = ${q.total_notes}`);
  }
  console.log(failed ? `\n${failed} check(s) failed` : '\nAll checks passed');
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });