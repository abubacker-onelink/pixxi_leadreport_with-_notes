'use strict';
// Builds a fake Pixi export so you can test without real data. Usage: node sample/make-sample.js [out.xlsx]
const ExcelJS = require('exceljs');
const path = require('path');
const cfg = require('../teams.json');

let seed = 42;
const rnd = () => { seed = (seed * 1664525 + 1013904223) % 4294967296; return seed / 4294967296; };
const pick = (a) => a[Math.floor(rnd() * a.length)];
const p2 = (n) => String(n).padStart(2, '0');
const fmt = (d) => `${p2(d.getUTCDate())}-${p2(d.getUTCMonth() + 1)}-${d.getUTCFullYear()} ${p2(d.getUTCHours())}:${p2(d.getUTCMinutes())}`;

(async () => {
  const out = process.argv[2] || path.join(__dirname, 'sample-export.xlsx');
  const agents = cfg.teams.flatMap((t) => t.members).filter((m) => m !== 'Marina Tok-Ogly').concat(['Zed Newagent']);
  const stages = cfg.stage_order;
  const sources = ['Meta Ads', 'Meta Ads', 'Meta Ads', 'Property Finder', 'Bayut', 'Referral', 'Cold Call'];
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Buy Leads');
  ws.addRow(['ID', 'Lead Name', 'Phone', 'Email', 'Agent Name', 'Create By', 'Source Of Lead', 'Ad campaign Name', 'Stage', 'Project Type',
    'Budget', 'Preferred Location', 'Create Time', 'First Contact Date', 'Last Follow Up Time', 'Last Note', 'Status']);
  const base = Date.UTC(2026, 9, 13, 0, 0);
  for (let i = 0; i < 360; i++) {
    const created = new Date(base + Math.floor(rnd() * 2.4 * 1440) * 60000);
    const source = pick(sources);
    const holder = pick(agents);
    const creator = source === 'Meta Ads' ? (rnd() < 0.3 ? pick(agents) : holder) : pick(['Lead Loader', holder]);
    const contacted = rnd() < 0.8;
    const fc = contacted ? new Date(created.getTime() + Math.floor(rnd() * rnd() * 600) * 60000) : null;
    const pool = rnd() < 0.08;
    const id = 1540000 + i;
    ws.addRow([id, `Lead ${i} Customer`, `+9715${Math.floor(rnd() * 1e8)}`, `lead${i}@example.com`, holder, creator, source,
      source === 'Meta Ads' ? pick(['Marina Q4', 'Downtown Launch']) : null, pick(stages),
      pick(['Apartment', 'Villa', 'Apartment|Villa', '']), pick(['1M', '2M', '3M']), pick(['Marina', 'Downtown', 'Palm']),
      fmt(created), fc ? fmt(fc) : null, fc && rnd() < 0.5 ? fmt(new Date(fc.getTime() + 3600000)) : null,
      i === 7 ? 'test lead, ignore' : rnd() < 0.1 ? 'Call me on +971 50 123 4567 or a@b.com' : 'Asked for brochure',
      pool ? 'Leads Pool' : 'Active Leads']);
  }
  // test account (must be excluded) and a lead outside the period (must be dropped)
  ws.addRow([1549001, 'Test Account Lead', '+971500000000', 't@example.com', 'Abubacker Sideek', 'Abubacker Sideek', 'Meta Ads', 'X', 'Not Contacted', '', '', '', fmt(new Date(base + 3600000)), null, null, '', 'Active Leads']);
  ws.addRow([1549002, 'Old Lead', '+971500000001', 'o@example.com', 'Adel Saad', 'Adel Saad', 'Bayut', null, 'Interested', '', '', '', '01-09-2026 10:00', null, null, '', 'Active Leads']);
  await wb.xlsx.writeFile(out);
  console.log('Sample written to ' + out);
})();
