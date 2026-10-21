'use strict';
const ExcelJS = require('exceljs');

// Only these columns are ever read. Phone, email and everything else are ignored on purpose.
const WANTED = [
  'ID', 'Lead Name', 'Agent Name', 'Create By', 'Source Of Lead', 'Stage', 'Ad campaign Name', 'Last Note',
  'Create Time', 'First Contact Date', 'Last Follow Up Time', 'Project Type', 'Budget', 'Preferred Location', 'Status',
];
const REQUIRED = ['ID', 'Lead Name', 'Agent Name', 'Create By', 'Source Of Lead', 'Stage', 'Last Note', 'Create Time'];

function cellValue(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('');
    if ('result' in v) return cellValue(v.result);
    if ('text' in v) return String(v.text);
    if ('error' in v) return null;
    return String(v);
  }
  return v;
}

/** Returns { rows: [{ID, 'Lead Name', ..., LS}], hasStatus } */
async function readExport(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('The Excel file has no sheets.');
  const header = {};
  ws.getRow(1).eachCell((cell, col) => {
    const name = String(cellValue(cell.value) ?? '').trim();
    if (name) header[name] = col;
  });
  const lower = {};
  Object.keys(header).forEach((k) => { lower[k.toLowerCase()] = header[k]; });
  const colOf = (name) => header[name] ?? lower[name.toLowerCase()];

  const missing = REQUIRED.filter((c) => !colOf(c));
  if (missing.length) {
    throw new Error(
      'This does not look like a Pixi Buy Leads export. Missing columns: ' + missing.join(', ') +
      '. If Pixi renamed a column, update REQUIRED in src/readExport.js.'
    );
  }
  const cols = {};
  WANTED.forEach((c) => { const i = colOf(c); if (i) cols[c] = i; });
  const hasStatus = !!cols.Status;

  const rows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    if (!row.hasValues) continue;
    const o = {};
    for (const [name, idx] of Object.entries(cols)) o[name] = cellValue(row.getCell(idx).value);
    if (o.ID === null && o['Lead Name'] === null && o['Create Time'] === null) continue;
    o.LS = hasStatus && /pool/i.test(String(o.Status ?? '')) ? 1 : 0;
    rows.push(o);
  }
  return { rows, hasStatus };
}

module.exports = { readExport, REQUIRED, WANTED };
