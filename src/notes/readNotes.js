'use strict';
const fs = require('fs');
const ExcelJS = require('exceljs');
const { toMinutes } = require('../dates');

/** Minimal RFC 4180 CSV parser (quotes, doubled quotes, multi-line fields, BOM, comma or semicolon). */
function parseCsv(text) {
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  const firstLine = text.split(/\r?\n/, 1)[0] || '';
  const delim = (firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length ? ';' : ',';
  const rows = [];
  let row = [], f = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c;
    } else if (c === '"') q = true;
    else if (c === delim) { row.push(f); f = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(f); f = ''; rows.push(row); row = [];
    } else f += c;
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  return rows.filter((r) => r.some((x) => x !== ''));
}

/** Note time -> Dubai wall-clock minutes. Epoch numbers and ISO strings with Z/offset are true instants (UTC) so +4h. */
function noteMinutes(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number' || /^\d{10,13}$/.test(String(v).trim())) {
    const n = Number(v);
    const ms = n > 1e11 ? n : n * 1000;
    return Math.floor(ms / 60000) + 240;
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}T[\d:.]+(Z|[+-]\d\d:?\d\d)$/.test(s)) {
    const t = Date.parse(s);
    return Number.isNaN(t) ? null : Math.floor(t / 60000) + 240;
  }
  return toMinutes(s);
}

/** Parses the "All Notes (JSON)" cell into [{by, message, time}] (unsorted). */
function parseNotesJson(raw) {
  const s = String(raw ?? '').trim();
  if (!s) return [];
  let data;
  try { data = JSON.parse(s); } catch { return null; } // null = unreadable
  if (!Array.isArray(data) && data && typeof data === 'object') {
    data = data.notes || data.data || data.items || Object.values(data).find(Array.isArray) || [];
  }
  if (!Array.isArray(data)) return [];
  return data.map((n) => ({
    by: String(n.createdBy ?? n.created_by ?? n.createdById ?? n.created_by_id ?? n.createdByUserId ?? n.authorId ?? n.userId ?? n.user_id ?? '').trim(),
    message: String(n.message ?? n.note ?? n.text ?? n.content ?? n.noteContent ?? n.description ?? '').replace(/\r/g, '').trim(),
    time: noteMinutes(n.createTime ?? n.created_at ?? n.createdAt ?? n.activityTime ?? n.createdDate ?? n.time),
  }));
}

/** Reads the notes CSV. Returns { leads: Map(id -> {id, fetchStatus, notes[]}), warnings[], statusCounts }. */
function readNotesCsv(filePath) {
  const rows = parseCsv(fs.readFileSync(filePath, 'utf8'));
  if (!rows.length) throw new Error('The notes CSV is empty.');
  const head = rows[0].map((h) => h.trim());
  const idx = (name) => head.findIndex((h) => h.toLowerCase() === name.toLowerCase());
  const iId = idx('Lead ID'), iJson = idx('All Notes (JSON)'), iStatus = idx('Fetch Status');
  if (iId < 0 || iJson < 0) {
    throw new Error('The notes CSV needs the columns "Lead ID" and "All Notes (JSON)". Found: ' + head.join(', '));
  }
  const leads = new Map(), warnings = [], statusCounts = {};
  let unreadable = 0;
  for (const r of rows.slice(1)) {
    const id = Number(String(r[iId]).trim());
    if (!Number.isFinite(id) || id <= 0) continue;
    const st = iStatus >= 0 ? String(r[iStatus] || '').trim() : '';
    if (st) statusCounts[st] = (statusCounts[st] || 0) + 1;
    let notes = parseNotesJson(r[iJson]);
    if (notes === null) { unreadable++; notes = []; }
    const prev = leads.get(id);
    if (!prev || notes.length > prev.notes.length) leads.set(id, { id, fetchStatus: st, notes });
  }
  if (unreadable) warnings.push(`${unreadable} leads had notes JSON that could not be read (treated as no notes).`);
  return { leads, warnings, statusCounts };
}

// ---- leads Excel (Buy export, Portal export, or an already prepared leads.xlsx)
const ALIASES = {
  id: ['ID', 'Lead ID'],
  name: ['Lead Name', 'Name'],
  agent: ['Agent Name', 'Agent'],
  stage: ['Stage', 'Type'],
  source: ['Source Of Lead', 'Source', 'Portal'],
  created: ['Create Time', 'Created', 'Lead Time'],
  campaign: ['Ad campaign Name', 'Campaign'],
};
function cellVal(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('');
    if ('result' in v) return cellVal(v.result);
    if ('text' in v) return String(v.text);
    return null;
  }
  return v;
}

async function readLeads(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('The leads Excel file has no sheets.');
  const header = {};
  ws.getRow(1).eachCell((c, col) => { const n = String(cellVal(c.value) ?? '').trim().toLowerCase(); if (n) header[n] = col; });
  const col = {};
  for (const [k, names] of Object.entries(ALIASES)) {
    const hit = names.find((n) => header[n.toLowerCase()]);
    if (hit) col[k] = header[hit.toLowerCase()];
  }
  const missing = ['id', 'name', 'agent'].filter((k) => !col[k]);
  if (missing.length) {
    throw new Error('The leads Excel needs ID, Lead Name (or Name) and Agent Name (or Agent) columns. Missing: ' + missing.join(', '));
  }
  const kind = header['portal'] ? 'Portal' : header['source of lead'] ? 'Buy' : 'Leads';
  const leads = new Map();
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    if (!row.hasValues) continue;
    const g = (k) => (col[k] ? cellVal(row.getCell(col[k]).value) : null);
    const id = Number(g('id'));
    if (!Number.isFinite(id) || id <= 0 || leads.has(id)) continue;
    leads.set(id, {
      id, name: String(g('name') ?? '').replace(/\uFFFD/g, '').trim(), agent: String(g('agent') ?? '').trim() || 'Unassigned',
      stage: String(g('stage') ?? '').trim(), source: String(g('source') ?? '').trim(), campaign: String(g('campaign') ?? '').trim(), created: toMinutes(g('created')),
    });
  }
  return { leads, kind };
}

module.exports = { parseCsv, parseNotesJson, noteMinutes, readNotesCsv, readLeads };
