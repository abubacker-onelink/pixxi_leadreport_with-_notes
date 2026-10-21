'use strict';
const fs = require('fs');
const path = require('path');
const { readExport } = require('./readExport');
const { prepare, buildHtml } = require('./prepareData');
const { build } = require('./excelReport');
const D = require('./dates');

const ROOT = path.resolve(__dirname, '..');

/**
 * Main entry point. Reuse this from any other workflow.
 * @param {string} xlsxPath  Pixi Buy Leads export (Active Leads + Leads Pool in one file)
 * @param {object} [options]
 *   start, end      "DD-MM-YYYY HH:MM" (default: first Create Time day 00:00 to last day 23:59, or now if the last day is today)
 *   export          "DD-MM-YYYY HH:MM" (default: current Dubai time)
 *   source          "" for all sources, or "Meta Ads" to pre-select Meta Ads in the dashboard
 *   excludeIds      array or comma string of extra lead IDs to exclude
 *   outRoot         output folder (default ./output)
 *   teamsPath, templatePath, log
 * @returns {Promise<{qc, files:{dashboard,excel,excel_pool?}, outDir, period, source}>}
 */
async function runReport(xlsxPath, options = {}) {
  const log = options.log || (() => {});
  const cfg = JSON.parse(fs.readFileSync(options.teamsPath || path.join(ROOT, 'teams.json'), 'utf8'));
  const template = fs.readFileSync(options.templatePath || path.join(ROOT, 'assets', 'dashboard_template.html'), 'utf8');

  log('Reading Excel file...');
  const { rows } = await readExport(xlsxPath);
  if (!rows.length) throw new Error('The Excel file has no lead rows.');

  const times = rows.map((r) => D.toMinutes(r['Create Time'])).filter((m) => m !== null);
  if (!times.length) throw new Error('No valid Create Time values found in the file.');
  const now = D.dubaiNow();
  const lastDay = D.startOfDay(Math.max(...times));
  const parse = (s, name) => {
    const m = D.toMinutes(s);
    if (m === null) throw new Error(`Could not read ${name} "${s}". Use DD-MM-YYYY HH:MM.`);
    return m;
  };
  const start = options.start ? parse(options.start, 'start') : D.startOfDay(Math.min(...times));
  const end = options.end ? parse(options.end, 'end') : (lastDay === D.startOfDay(now) ? now : lastDay + 23 * 60 + 59);
  const exportMin = options.export ? parse(options.export, 'export time') : now;
  if (end < start) throw new Error('The end of the period is before the start.');
  const excludeIds = new Set(
    (Array.isArray(options.excludeIds) ? options.excludeIds : String(options.excludeIds || '').split(','))
      .map((x) => Number(String(x).trim())).filter((x) => Number.isFinite(x) && x > 0)
  );
  const source = options.source && /meta/i.test(options.source) ? 'Meta Ads' : '';
  log(`Period ${D.fmtInput(start)} to ${D.fmtInput(end)} (Dubai time), ${source || 'all sources'}`);

  const { D: data, cleanRows, qc } = prepare(rows, { start, end, export: exportMin, cfg, excludeIds, source, title: 'Buy leads report' });
  if (!cleanRows.length) throw new Error('No leads left after applying the period and exclusions.');
  const lastPartial = end < lastDay + 23 * 60 + 59 && D.startOfDay(end) === D.startOfDay(now);
  if (lastPartial) qc.warnings.push('The last day is a part day (report ends at the current time).');

  const label = `${D.fmtIso(start)}_to_${D.fmtIso(end)}`;
  const outDir = path.join(options.outRoot || path.join(ROOT, 'output'), `leads-report_${label}`);
  fs.mkdirSync(outDir, { recursive: true });

  const files = {};
  files.dashboard = path.join(outDir, 'leads_dashboard.html');
  fs.writeFileSync(files.dashboard, buildHtml(template, data), 'utf8');
  log('Dashboard written');

  const stamp = D.startOfDay(start) === D.startOfDay(end) ? D.fmtDay(start) : `${D.fmtDay(start)} to ${D.fmtDay(end)}`;
  const active = cleanRows.filter((r) => !r.ls), pool = cleanRows.filter((r) => r.ls);
  files.excel = path.join(outDir, `Buy_Leads_Report_${label}.xlsx`);
  await build(active, cfg, `Buy Leads Report: ${stamp}`, start, end, exportMin, files.excel);
  log('Excel report written');
  if (pool.length) {
    files.excel_pool = path.join(outDir, `Buy_Leads_from_Lead_Pool_${label}.xlsx`);
    await build(pool, cfg, `Buy Leads from Lead Pool: ${stamp}`, start, end, exportMin, files.excel_pool);
    log('Lead pool Excel written');
  }
  return { qc, files, outDir, period: [D.fmtInput(start), D.fmtInput(end)], source: source || 'All sources' };
}

module.exports = { runReport };
