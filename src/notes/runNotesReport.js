'use strict';
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const D = require('../dates');
const { redact } = require('../util');
const { readNotesCsv, readLeads } = require('./readNotes');
const { fetchPixiNotes } = require('./fetchPixiNotes');

const ROOT = path.resolve(__dirname, '..', '..');
const NAVY = 'FF1F2A44', TOTAL = 'FFD9E1F2', GREY = 'FF5B6776';
const F = (o = {}) => ({ name: 'Arial', size: 10, ...o });
const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });

function loadAuthors(p) {
  const raw = JSON.parse(fs.readFileSync(p, 'utf8'));
  const m = {};
  for (const [k, v] of Object.entries(raw)) if (!k.startsWith('_')) m[String(k)] = String(v);
  return m;
}

/**
 * @param notesCsv   path to the lead-notes CSV, or null to fetch notes from Pixi
 * @param leadsXlsx  path to the Pixi leads export (Buy or Portal)
 * @param options    { title, outRoot, authorsPath, newAuthors: {id: name}, log }
 * @returns {{qc, file, outDir, kind}}
 */
async function runNotesReport(notesCsv, leadsXlsx, options = {}) {
  const log = options.log || (() => {});
  const authorsPath = options.authorsPath || path.join(ROOT, 'authors.json');
  const authors = loadAuthors(authorsPath);
  if (options.newAuthors && Object.keys(options.newAuthors).length) {
    const raw = JSON.parse(fs.readFileSync(authorsPath, 'utf8'));
    for (const [id, name] of Object.entries(options.newAuthors)) {
      if (/^\d+$/.test(String(id)) && String(name).trim()) { raw[id] = String(name).trim(); authors[id] = String(name).trim(); }
    }
    fs.writeFileSync(authorsPath, JSON.stringify(raw, null, 2) + '\n');
  }

  log('Reading leads file...');
  const { leads: leadInfo, kind } = await readLeads(leadsXlsx);
  const { leads: notesByLead, warnings, statusCounts } = options.fetchFromPixi
    ? await fetchPixiNotes([...leadInfo.keys()], { token: options.pixxiToken, fetchImpl: options.fetchImpl, log })
    : readNotesCsv(notesCsv);
  if (!notesByLead.size) throw new Error(options.fetchFromPixi ? 'No valid lead IDs found in the leads Excel.' : 'No leads found in the notes CSV.');

  // ---- join (universe = IDs in the notes source)
  const unknownIds = new Set();
  const notFound = [];
  const items = [];
  for (const [id, n] of notesByLead) {
    let info = leadInfo.get(id);
    if (!info) { notFound.push(id); info = { id, name: '(not in leads file)', agent: 'Unknown', stage: '', source: '', created: null }; }
    const notes = n.notes
      .map((x) => {
        const who = authors[x.by];
        if (!who) unknownIds.add(x.by || '(blank)');
        return { time: x.time, text: redact(x.message), by: who || 'Unknown' };
      })
      .sort((a, b) => (a.time ?? 9e15) - (b.time ?? 9e15));
    items.push({ ...info, notes });
  }
  const withNotes = items.filter((i) => i.notes.length > 0);
  const without = items.filter((i) => i.notes.length === 0);
  const notInNotes = [...leadInfo.keys()].filter((id) => !notesByLead.has(id)).length;

  const allTimes = withNotes.flatMap((i) => i.notes.map((x) => x.time)).filter((t) => t !== null);
  const latest = allTimes.length ? Math.max(...allTimes) : null;
  const nowMin = D.dubaiNow();
  const qc = {
    kind, leads: items.length, with_notes: withNotes.length, without_notes: without.length,
    total_notes: withNotes.reduce((s, i) => s + i.notes.length, 0),
    lead_ids_not_found_in_leads_file: notFound, unknown_author_ids: [...unknownIds],
    leads_in_leads_file_but_not_in_notes_file: notInNotes, fetch_status_counts: statusCounts,
    notes_through: latest === null ? null : D.fmtFull(latest), warnings,
  };
  if (notFound.length) qc.warnings.push(`${notFound.length} lead IDs are not in the leads file and were added with agent "Unknown".`);
  if (unknownIds.size) qc.warnings.push('Unknown note author IDs: ' + [...unknownIds].join(', ') + '. Add their names to authors.json (shown as "Unknown" meanwhile).');

  // ---- workbook
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Buy Leads Dashboard';
  wb.calcProperties = { fullCalcOnLoad: true };
  const title = options.title || `${kind} leads notes report`;

  // ===== Leads with Notes
  const wsN = wb.addWorksheet('Leads with Notes');
  const maxN = withNotes.reduce((m, i) => Math.max(m, i.notes.length), 0);
  const fixed = ['Lead ID', 'Lead Name', 'Created', 'Agent', 'Stage', 'Notes', 'First note', 'Last note'];
  const heads = fixed.concat(Array.from({ length: maxN }, (_, k) => `Follow-up ${k + 1}`));
  heads.forEach((h, i) => {
    const c = wsN.getCell(1, i + 1);
    c.value = h; c.font = F({ bold: true, color: { argb: 'FFFFFFFF' } }); c.fill = fill(NAVY);
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });
  wsN.getRow(1).height = 26;
  const day = (i) => (i.created === null ? -1 : D.startOfDay(i.created));
  withNotes.sort((a, b) => day(a) - day(b) || b.notes.length - a.notes.length || a.agent.localeCompare(b.agent) || a.id - b.id);
  let prevDay = null, shade = false;
  withNotes.forEach((it, ri) => {
    const r = ri + 2, d = day(it);
    const newDay = prevDay !== null && d !== prevDay;
    if (prevDay === null || newDay) shade = prevDay === null ? true : !shade;
    prevDay = d;
    const base = shade ? 'FFF3F5F9' : 'FFFFFFFF';
    const times = it.notes.map((x) => x.time).filter((t) => t !== null);
    const vals = [it.id, redact(it.name), it.created === null ? null : D.toDate(it.created), it.agent, it.stage, it.notes.length,
      times.length ? D.toDate(times[0]) : null, times.length ? D.toDate(times[times.length - 1]) : null];
    vals.forEach((v, i) => { if (v !== null) wsN.getCell(r, i + 1).value = v; });
    wsN.getCell(r, 3).numFmt = 'dd-mmm-yyyy hh:mm'; wsN.getCell(r, 7).numFmt = 'dd-mmm-yyyy hh:mm'; wsN.getCell(r, 8).numFmt = 'dd-mmm-yyyy hh:mm';
    const sameAgent = (a) => a.trim().toLowerCase() === it.agent.trim().toLowerCase();
    it.notes.forEach((nt, k) => {
      const dateLine = nt.time === null ? '(no date)' : D.fmtDow(nt.time);
      const by = sameAgent(nt.by) ? '' : ` (by ${nt.by})`;
      wsN.getCell(r, 9 + k).value = { richText: [
        { font: F({ bold: true }), text: dateLine + '\n' },
        { font: F(), text: (nt.text || '(empty note)').slice(0, 3000) + by },
      ] };
    });
    for (let c = 1; c <= heads.length; c++) {
      const cell = wsN.getCell(r, c);
      const followIdx = c - 8; // 1-based follow-up number
      const tint = followIdx > 0 && followIdx % 2 === 0 ? (shade ? 'FFE3E9F4' : 'FFEEF2F8') : base;
      cell.fill = fill(tint);
      cell.alignment = { vertical: 'top', wrapText: c >= 9 || c === 2, horizontal: c === 1 || c === 6 ? 'center' : 'left' };
      if (!cell.font || cell.font.name !== 'Arial') cell.font = F();
      if (newDay) cell.border = { top: { style: 'medium', color: { argb: 'FF8A96A3' } } };
    }
  });
  if (!withNotes.length) { wsN.getCell('A2').value = 'No leads with notes.'; wsN.getCell('A2').font = F({ italic: true, color: { argb: GREY } }); }
  wsN.views = [{ state: 'frozen', xSplit: 2, ySplit: 1 }];
  [10, 24, 18, 22, 22, 8, 18, 18].forEach((w, i) => { wsN.getColumn(i + 1).width = w; });
  for (let k = 0; k < maxN; k++) wsN.getColumn(9 + k).width = 38;

  // ===== Leads without Notes
  const wsW = wb.addWorksheet('Leads without Notes');
  ['Agent', 'Lead ID', 'Lead Name', 'Created', 'Stage', 'Source'].forEach((h, i) => {
    const c = wsW.getCell(1, i + 1);
    c.value = h; c.font = F({ bold: true, color: { argb: 'FFFFFFFF' } }); c.fill = fill(NAVY);
    c.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  const groups = new Map();
  without.forEach((i) => { if (!groups.has(i.agent)) groups.set(i.agent, []); groups.get(i.agent).push(i); });
  const ordered = [...groups.entries()].sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0]));
  let wr = 2;
  ordered.forEach(([agent, list]) => {
    list.sort((a, b) => (a.created ?? 9e15) - (b.created ?? 9e15) || a.id - b.id);
    const start = wr;
    list.forEach((it) => {
      wsW.getCell(wr, 2).value = it.id; wsW.getCell(wr, 3).value = redact(it.name);
      if (it.created !== null) { wsW.getCell(wr, 4).value = D.toDate(it.created); wsW.getCell(wr, 4).numFmt = 'dd-mmm-yyyy hh:mm'; }
      wsW.getCell(wr, 5).value = it.stage; wsW.getCell(wr, 6).value = it.source;
      for (let c = 1; c <= 6; c++) wsW.getCell(wr, c).font = F();
      wr++;
    });
    wsW.getCell(start, 1).value = agent;
    wsW.getCell(start, 1).font = F({ bold: true });
    if (wr - 1 > start) wsW.mergeCells(start, 1, wr - 1, 1);
    wsW.getCell(start, 1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    for (let c = 1; c <= 6; c++) wsW.getCell(wr - 1, c).border = { bottom: { style: 'thin', color: { argb: 'FF8A96A3' } } };
  });
  if (!without.length) { wsW.getCell('A2').value = 'Every lead has at least one note.'; wsW.getCell('A2').font = F({ italic: true, color: { argb: GREY } }); }
  wsW.views = [{ state: 'frozen', ySplit: 1 }];
  [26, 10, 28, 18, 22, 18].forEach((w, i) => { wsW.getColumn(i + 1).width = w; });

  // ===== Summary (formulas on Leads with Notes; total leads per agent is data)
  const wsS = wb.addWorksheet('Summary');
  wsS.getCell('A1').value = title; wsS.getCell('A1').font = F({ size: 14, bold: true, color: { argb: NAVY } });
  wsS.getCell('A2').value = `Notes included up to ${qc.notes_through || 'n/a'} (report created ${D.fmtFull(nowMin)}, Dubai time).`;
  wsS.getCell('A2').font = F({ size: 9, italic: true, color: { argb: GREY } });
  ['Agent', 'Total leads', 'With notes', 'Without notes', '% with notes', 'Total notes'].forEach((h, i) => {
    const c = wsS.getCell(4, i + 1);
    c.value = h; c.font = F({ bold: true, color: { argb: 'FFFFFFFF' } }); c.fill = fill(NAVY);
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });
  const perAgent = new Map();
  items.forEach((i) => perAgent.set(i.agent, (perAgent.get(i.agent) || 0) + 1));
  const agentRows = [...perAgent.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  const lastRow = Math.max(2, withNotes.length + 1);
  agentRows.forEach(([agent, total], i) => {
    const r = 5 + i;
    wsS.getCell(r, 1).value = agent;
    wsS.getCell(r, 2).value = total;
    wsS.getCell(r, 3).value = { formula: `COUNTIF('Leads with Notes'!$D$2:$D$${lastRow},A${r})` };
    wsS.getCell(r, 4).value = { formula: `B${r}-C${r}` };
    wsS.getCell(r, 5).value = { formula: `IFERROR(C${r}/B${r},0)` };
    wsS.getCell(r, 6).value = { formula: `SUMIF('Leads with Notes'!$D$2:$D$${lastRow},A${r},'Leads with Notes'!$F$2:$F$${lastRow})` };
    wsS.getCell(r, 5).numFmt = '0.0%';
    for (let c = 1; c <= 6; c++) wsS.getCell(r, c).font = F();
  });
  const t = 5 + agentRows.length;
  wsS.getCell(t, 1).value = 'TOTAL';
  ['B', 'C', 'D', 'F'].forEach((L) => { wsS.getCell(`${L}${t}`).value = { formula: `SUM(${L}5:${L}${t - 1})` }; });
  wsS.getCell(t, 5).value = { formula: `IFERROR(C${t}/B${t},0)` }; wsS.getCell(t, 5).numFmt = '0.0%';
  for (let c = 1; c <= 6; c++) { wsS.getCell(t, c).font = F({ bold: true }); wsS.getCell(t, c).fill = fill(TOTAL); }
  [28, 12, 12, 14, 12, 12].forEach((w, i) => { wsS.getColumn(i + 1).width = w; });
  wsS.views = [{ state: 'frozen', ySplit: 4 }];

  // tab order like the skill: notes, without notes, summary
  const outDir = path.join(options.outRoot || path.join(ROOT, 'output'), `lead-notes_${D.fmtIso(nowMin)}`);
  fs.mkdirSync(outDir, { recursive: true });
  const file = path.join(outDir, `Lead_Notes_Report_${kind}_${D.fmtIso(nowMin)}.xlsx`);
  await wb.xlsx.writeFile(file);
  log('Lead notes report written');
  return { qc, file, outDir, kind };
}

module.exports = { runNotesReport };
