'use strict';
// Fake lead-notes CSV + leads Excel. Usage: node sample/make-notes-sample.js [outDir]
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

let seed = 7;
const rnd = () => { seed = (seed * 1664525 + 1013904223) % 4294967296; return seed / 4294967296; };
const pick = (a) => a[Math.floor(rnd() * a.length)];
const p2 = (n) => String(n).padStart(2, '0');
const fmt = (d) => `${p2(d.getUTCDate())}-${p2(d.getUTCMonth() + 1)}-${d.getUTCFullYear()} ${p2(d.getUTCHours())}:${p2(d.getUTCMinutes())}`;
const iso = (d) => d.toISOString().slice(0, 19).replace('T', ' ');

(async () => {
  const dir = process.argv[2] || __dirname;
  fs.mkdirSync(dir, { recursive: true });
  const authors = { 9509: 'Mazen Abou Karroum', 9528: 'Adel Saad', 9511: 'Gouthaman Asokan', 9541: 'Ainam Dehgani', 9506: 'Omama Hussain' };
  const ids = Object.keys(authors);
  const stages = ['Not Contacted', 'No Answer', 'Interested', 'Follow-up', 'Not Interested'];
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Buy Leads');
  ws.addRow(['ID', 'Lead Name', 'Phone', 'Agent Name', 'Stage', 'Source Of Lead', 'Create Time']);
  const csv = ['Lead ID,Fetch Status,Note Count,Latest Note,Latest Note Time,All Notes (JSON)'];
  const q = (s) => '"' + String(s).replace(/"/g, '""') + '"';
  const base = Date.UTC(2026, 9, 12, 8, 0);
  for (let i = 0; i < 60; i++) {
    const id = 1550000 + i;
    const aid = pick(ids), agent = authors[aid];
    const created = new Date(base + Math.floor(rnd() * 3 * 1440) * 60000);
    ws.addRow([id, `Customer ${i}`, '+971500000000', agent, pick(stages), 'Meta Ads', fmt(created)]);
    const n = rnd() < 0.35 ? 0 : 1 + Math.floor(rnd() * 5);
    const notes = [];
    for (let k = 0; k < n; k++) {
      const by = rnd() < 0.2 ? pick(ids) : aid;
      notes.push({ createdBy: Number(by), message: k === 1 && i % 7 === 0 ? 'Call back on +971 50 123 4567' : `Follow-up ${k} for lead ${i}, said "call later"`,
        createTime: iso(new Date(created.getTime() + (k + 1) * 3 * 3600000 + Math.floor(rnd() * 3600000))) });
    }
    if (i === 3) notes.push({ createdBy: 9999, message: 'Note from unknown user', createTime: iso(new Date(created.getTime() + 9 * 3600000)) });
    notes.sort(() => rnd() - 0.5); // stored in random order on purpose
    const latest = [...notes].sort((a, b) => (a.createTime < b.createTime ? 1 : -1))[0];
    csv.push([id, 'success', notes.length, q(latest ? latest.message : ''), latest ? latest.createTime : '', q(JSON.stringify(notes))].join(','));
  }
  csv.push([1559999, 'success', 1, q('Lead missing from leads file'), '2026-10-13 10:00:00', q(JSON.stringify([{ createdBy: 9509, message: 'Orphan note', createTime: '2026-10-13 10:00:00' }]))].join(','));
  fs.writeFileSync(path.join(dir, 'sample-notes.csv'), csv.join('\r\n'));
  await wb.xlsx.writeFile(path.join(dir, 'sample-leads.xlsx'));
  console.log('Wrote sample-notes.csv and sample-leads.xlsx in ' + dir);
})();
