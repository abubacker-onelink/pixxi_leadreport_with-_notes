'use strict';
// Run with: npm test   (needs a generated sample; this script makes one)
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync, spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const { runReport } = require('../src/pipeline');

let failed = 0;
const ok = (cond, msg) => { console.log((cond ? 'PASS ' : 'FAIL ') + msg); if (!cond) failed++; };

(async () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'leads-test-'));
  const sample = path.join(tmp, 'sample.xlsx');
  execFileSync('node', [path.join(__dirname, '..', 'sample', 'make-sample.js'), sample], { stdio: 'ignore' });

  const res = await runReport(sample, { start: '13-10-2026 00:00', end: '15-10-2026 23:59', export: '16-10-2026 10:00', outRoot: path.join(tmp, 'out') });
  const qc = res.qc;
  ok(qc.test_leads_excluded === 1, 'test account lead excluded');
  ok(qc.leads_outside_period_excluded === 1, 'lead outside the period dropped');
  ok(qc.leads_in_report === 360, `360 leads in report (got ${qc.leads_in_report})`);
  ok(!!res.files.excel_pool, 'lead pool workbook created when pool has leads');

  // dashboard data is valid JSON and matches the QC count
  const html = fs.readFileSync(res.files.dashboard, 'utf8');
  const m = /var D=(\{.*?\});\n/s.exec(html);
  let D = null;
  try { D = JSON.parse(m[1].replace(/<\\\//g, '</')); } catch (e) { /* handled below */ }
  ok(D && D.rows.length === qc.leads_in_report, 'dashboard data parses and row count matches');
  ok(D && D.teams.length === 5 && D.teams[4].name === 'Unmapped', 'teams include Unmapped');
  ok(D && D.rows.every((r) => r.length === 13), 'every dashboard row has 13 fields');

  // privacy: no phone numbers or emails in any output
  const leak = /\+971\s?\d|@example\.com|a@b\.com/;
  ok(!leak.test(html), 'no phone/email in dashboard HTML');
  for (const f of [res.files.excel, res.files.excel_pool]) {
    const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(f);
    let bad = 0;
    wb.eachSheet((ws) => ws.eachRow((row) => row.eachCell((c) => { if (typeof c.value === 'string' && leak.test(c.value)) bad++; })));
    ok(bad === 0, 'no phone/email in ' + path.basename(f));
  }

  // recalculate with LibreOffice (if installed) and check totals + formula errors
  const so = spawnSync('soffice', ['--version']);
  if (so.error) {
    console.log('SKIP Excel recalculation checks (LibreOffice not installed).');
  } else {
    const calcDir = path.join(tmp, 'calc');
    for (const f of [res.files.excel, res.files.excel_pool]) {
      spawnSync('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', calcDir, f], { stdio: 'ignore' });
      const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(path.join(calcDir, path.basename(f)));
      const name = path.basename(f);
      let errors = 0;
      wb.eachSheet((ws) => ws.eachRow((row) => row.eachCell((c) => {
        const v = c.value;
        if (v && typeof v === 'object' && 'error' in v) errors++;
        if (typeof v === 'string' && /^#(REF|NAME|VALUE|DIV|N\/A|NUM)/.test(v)) errors++;
        if (v && typeof v === 'object' && typeof v.result === 'string' && /^#/.test(v.result)) errors++;
      })));
      ok(errors === 0, `no formula errors in ${name}`);
      const val = (c) => { const v = c.value; return v && typeof v === 'object' && 'result' in v ? v.result : v; };
      const sum = wb.getWorksheet('Summary');
      const isPool = /Pool/.test(name);
      const expected = isPool ? qc.by_status.pool : qc.by_status.active;
      ok(val(sum.getCell('C10')) === expected, `${name}: Summary total = ${expected} (got ${val(sum.getCell('C10'))})`);
      const team = wb.getWorksheet('By Team');
      let tt = 5; while (team.getCell(tt, 1).value !== 'TOTAL') tt++;
      ok(val(team.getCell(tt, 3)) === expected, `${name}: By Team holding-now total matches`);
      const agent = wb.getWorksheet('By Agent');
      let at = 5; while (agent.getCell(at, 1).value !== 'TOTAL') at++;
      ok(val(agent.getCell(at, 3)) === expected, `${name}: By Agent holding-now total matches`);
      const day = wb.getWorksheet('Daily Trend');
      let dt = 5; while (day.getCell(dt, 1).value !== 'TOTAL') dt++;
      ok(val(day.getCell(dt, 2)) === expected, `${name}: Daily Trend total matches`);
      const timing = wb.getWorksheet('Contact Timing');
      let mt = 5; while (timing.getCell(mt, 1).value !== 'TOTAL') mt++;
      ok(val(timing.getCell(mt, 3)) + val(timing.getCell(mt, 10)) === expected, `${name}: contact timing buckets add up`);
      // Meta Ads numbers counted independently from this workbook's own Lead Data tab
      const raw = new ExcelJS.Workbook(); await raw.xlsx.readFile(f);
      const ld = raw.getWorksheet('Lead Data');
      let meta = 0, moved = 0;
      for (let r = 2; r <= ld.rowCount; r++) {
        if (ld.getCell(r, 5).value === 'Meta Ads') { meta++; if (ld.getCell(r, 11).value !== ld.getCell(r, 3).value) moved++; }
      }
      ok(val(sum.getCell('C23')) === meta, `${name}: Meta Ads count matches (${meta})`);
      ok(val(sum.getCell('C24')) === moved, `${name}: Meta reassigned count matches (${moved})`);
      if (!isPool) {
        const mw = wb.getWorksheet('Meta Ads Round Robin');
        let t = 15; while (mw.getCell(t, 1).value !== 'TOTAL') t++;
        ok(val(mw.getCell(t, 3)) === val(mw.getCell(t, 6)) - val(mw.getCell(t, 5)) + val(mw.getCell(t, 4)), `${name}: round robin scorecard balances`);
      }
    }
  }
  console.log(failed ? `\n${failed} check(s) failed` : '\nAll checks passed');
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
