'use strict';
// Builds one Excel workbook with live formulas that point at a "Lead Data" tab.
// Excel recalculates everything when the file is opened (fullCalcOnLoad).
const ExcelJS = require('exceljs');
const { toDate, startOfDay, fmtDay, fmtFull } = require('./dates');

const NAVY = 'FF1F3864', TOTAL = 'FFD9E1F2', BLUE = 'FF0000FF', GREY = 'FF5B6776', BAR = 'FF4F5BD5';
const LD = "'Lead Data'!";
const C = { id: 'A', name: 'B', agent: 'C', team: 'D', source: 'E', stage: 'F', project: 'G', created: 'H', fc: 'I', lf: 'J', by: 'K',
  camp: 'L', ls: 'M', budget: 'N', loc: 'O', note: 'P', cdate: 'Q', contacted: 'R', mins: 'S', reassigned: 'T', cteam: 'U', inp: 'V' };
const DT = 'dd-mmm-yyyy hh:mm';

function build(cleanRows, cfg, title, start, end, exportMin, outPath) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Buy Leads Dashboard';
  wb.calcProperties = { fullCalcOnLoad: true };
  const n = cleanRows.length;
  const last = Math.max(2, n + 1);
  const rng = (k) => `${LD}$${C[k]}$2:$${C[k]}$${last}`;
  // COUNTIFS limited to the report period (Summary Start/End inputs). pairs: [colKey, criteria]
  const cnt = (...pairs) => `COUNTIFS(${[...pairs, ['inp', '1']].map(([k, c]) => `${rng(k)},${c}`).join(',')})`;
  const avg = (...pairs) => `IFERROR(AVERAGEIFS(${rng('mins')},${[...pairs, ['inp', '1']].map(([k, c]) => `${rng(k)},${c}`).join(',')}),"-")`;

  const excluded = new Set(cfg.exclude_agents || []);
  const teams = cfg.teams;
  const teamOf = new Map();
  teams.forEach((t) => t.members.forEach((m) => teamOf.set(m, t.name)));
  const hotStages = cfg.hot_stages, lostStages = cfg.lost_stages;

  // ---- agent / stage / source lists worked out from the data (formulas do the counting)
  const holders = new Map(), creators = new Map();
  cleanRows.forEach((r) => { holders.set(r.agent, (holders.get(r.agent) || 0) + 1); creators.set(r.by, (creators.get(r.by) || 0) + 1); });
  const extras = [...new Set([...holders.keys(), ...creators.keys()])].filter((a) => !teamOf.has(a) && !excluded.has(a));
  const hasUnmapped = extras.length > 0;
  const byHold = (a, b) => (holders.get(b) || 0) - (holders.get(a) || 0);
  const agentList = [];
  teams.forEach((t) => [...t.members].filter((m) => !excluded.has(m)).sort(byHold).forEach((m) => agentList.push(m)));
  extras.sort(byHold).forEach((a) => agentList.push(a));
  const activeAgents = agentList.filter((a) => (holders.get(a) || 0) > 0);
  const teamNames = teams.map((t) => t.name).concat(hasUnmapped ? ['Unmapped'] : []);
  const teamLeader = (name) => (teams.find((t) => t.name === name) || { leader: '-' }).leader;
  const stagesPresent = new Set(cleanRows.map((r) => r.stage));
  const order = cfg.stage_order || [];
  const stages = order.filter((s) => stagesPresent.has(s)).concat([...stagesPresent].filter((s) => !order.includes(s)));
  const srcCount = new Map();
  cleanRows.forEach((r) => srcCount.set(r.source, (srcCount.get(r.source) || 0) + 1));
  const sources = [...srcCount.entries()].sort((a, b) => b[1] - a[1]).map((e) => e[0]);
  const hasMeta = cleanRows.some((r) => r.source === 'Meta Ads');

  // ---- sheets, in display order
  const wsSum = wb.addWorksheet('Summary');
  const wsTeam = wb.addWorksheet('By Team');
  const wsAgent = wb.addWorksheet('By Agent');
  const wsStage = wb.addWorksheet('Team-Agent x Stage');
  const wsTime = wb.addWorksheet('Contact Timing');
  const wsSrc = wb.addWorksheet('Agent x Source');
  const wsDay = wb.addWorksheet('Daily Trend');
  const wsMeta = hasMeta ? wb.addWorksheet('Meta Ads Round Robin') : null;
  const wsData = wb.addWorksheet('Lead Data');
  const wsRoster = wb.addWorksheet('Teams');

  // ---- helpers
  const head = (ws, row, labels, from = 1) => {
    labels.forEach((l, i) => {
      const c = ws.getCell(row, from + i);
      c.value = l;
      c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    ws.getRow(row).height = 30;
  };
  const totalRow = (ws, row, ncols) => {
    for (let i = 1; i <= ncols; i++) {
      const c = ws.getCell(row, i);
      c.font = { name: 'Arial', size: 10, bold: true };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TOTAL } };
    }
  };
  const titleBlock = (ws, sub) => {
    ws.getCell('A1').value = title;
    ws.getCell('A1').font = { name: 'Arial', size: 14, bold: true, color: { argb: NAVY } };
    ws.getCell('A2').value = sub || '';
    ws.getCell('A2').font = { name: 'Arial', size: 9, italic: true, color: { argb: GREY } };
  };
  const widths = (ws, arr) => arr.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const fmtCols = (ws, r1, r2, cols, f) => { for (let r = r1; r <= r2; r++) cols.forEach((c) => { ws.getCell(r, c).numFmt = f; }); };
  const teamLookup = (cellRef) => `IFERROR(VLOOKUP(${cellRef},Teams!$A:$B,2,FALSE),"Unmapped")`;

  // =========================================================== Lead Data
  const dataHeaders = ['ID', 'Lead Name', 'Agent', 'Team', 'Source', 'Stage', 'Project Type', 'Created', 'First Contact', 'Last Follow-up',
    'Created By', 'Campaign', 'Lead Status', 'Budget', 'Preferred Location', 'Last Note',
    'Created Date', 'Contacted', 'Mins to First Contact', 'Reassigned (Meta)', 'Creator Team', 'In Period'];
  head(wsData, 1, dataHeaders);
  cleanRows.forEach((r, i) => {
    const x = i + 2;
    const set = (col, v) => { if (v !== null && v !== undefined && v !== '') wsData.getCell(`${col}${x}`).value = v; };
    set('A', r.id); set('B', r.name); set('C', r.agent);
    wsData.getCell(`D${x}`).value = { formula: teamLookup(`C${x}`) };
    set('E', r.source); set('F', r.stage); set('G', r.project);
    set('H', toDate(r.ct)); if (r.fc) set('I', toDate(r.fc)); if (r.lf) set('J', toDate(r.lf));
    set('K', r.by); set('L', r.camp); set('M', r.ls ? 'Leads Pool' : 'Active Leads');
    set('N', r.budget); set('O', r.location); set('P', r.note);
    wsData.getCell(`Q${x}`).value = { formula: `INT(H${x})` };
    wsData.getCell(`R${x}`).value = { formula: `IF(I${x}="",0,1)` };
    wsData.getCell(`S${x}`).value = { formula: `IF(I${x}="","",ROUND((I${x}-H${x})*1440,0))` };
    wsData.getCell(`T${x}`).value = { formula: `IF(AND(E${x}="Meta Ads",K${x}<>C${x}),1,0)` };
    wsData.getCell(`U${x}`).value = { formula: teamLookup(`K${x}`) };
    wsData.getCell(`V${x}`).value = { formula: `IF(AND(H${x}>=Summary!$C$5,H${x}<Summary!$C$6+1/1440),1,0)` };
  });
  wsData.views = [{ state: 'frozen', ySplit: 1, xSplit: 2 }];
  wsData.autoFilter = { from: 'A1', to: `V${last}` };
  widths(wsData, [10, 26, 24, 16, 14, 22, 14, 18, 18, 18, 24, 26, 13, 14, 22, 40, 12, 10, 12, 12, 16, 10]);
  fmtCols(wsData, 2, last, [8, 9, 10], DT);
  fmtCols(wsData, 2, last, [17], 'dd-mmm-yyyy');

  // =========================================================== Teams roster
  head(wsRoster, 1, ['Agent', 'Team', 'Role']);
  let rr = 2;
  teams.forEach((t) => t.members.forEach((m) => {
    wsRoster.getCell(rr, 1).value = m; wsRoster.getCell(rr, 2).value = t.name;
    wsRoster.getCell(rr, 3).value = m === t.leader ? 'Leader' : 'Member'; rr++;
  }));
  wsRoster.getCell(rr + 1, 1).value = 'Edit teams.json (not this tab) when people join or move, then re-run the report.';
  wsRoster.getCell(rr + 1, 1).font = { name: 'Arial', size: 9, italic: true, color: { argb: GREY } };
  widths(wsRoster, [30, 20, 10]);
  wsRoster.views = [{ state: 'frozen', ySplit: 1 }];

  // =========================================================== Summary
  titleBlock(wsSum, `Report period inputs are blue: change Start/End and every tab updates. Exported ${fmtFull(exportMin)} (Dubai time).`);
  head(wsSum, 4, ['', 'Report period', ''], 1);
  [['Start', start], ['End', end], ['Export time', exportMin]].forEach(([l, v], i) => {
    wsSum.getCell(5 + i, 2).value = l;
    const c = wsSum.getCell(5 + i, 3);
    c.value = toDate(v); c.numFmt = DT; c.font = { name: 'Arial', size: 10, color: { argb: BLUE } };
  });
  head(wsSum, 9, ['', 'Key numbers', 'Value'], 1);
  const hotFirst = 10, hotLast = 9 + hotStages.length, lostLast = 9 + lostStages.length;
  const hotRange = `Summary!$E$${hotFirst}:$E$${hotLast}`, lostRange = `Summary!$H$${hotFirst}:$H$${lostLast}`;
  // simpler, explicit KPI formulas (rows 10..25)
  const K = [
    ['Total leads', `COUNTIFS(${rng('inp')},1)`, '#,##0'],
    ['Active leads', cnt(['ls', '"Active Leads"']), '#,##0'],
    ['Buy leads from lead pool', cnt(['ls', '"Leads Pool"']), '#,##0'],
    ['Leads with a first contact', cnt(['contacted', '1']), '#,##0'],
    ['% with a first contact', 'IFERROR(C13/C10,0)', '0.0%'],
    ['Leads not contacted yet', 'C10-C13', '#,##0'],
    ['Hot leads', `SUM(F${hotFirst}:F${hotLast})`, '#,##0'],
    ['Hot lead rate', 'IFERROR(C16/C10,0)', '0.0%'],
    ['Lost leads', `SUM(I${hotFirst}:I${lostLast})`, '#,##0'],
    ['Lost lead rate', 'IFERROR(C18/C10,0)', '0.0%'],
    ['Avg minutes to first contact', avg(['contacted', '1']), '#,##0'],
    ['First contacted within 30 min', cnt(['contacted', '1'], ['mins', '"<=30"']), '#,##0'],
    ['% of contacted within 30 min', 'IFERROR(C21/C13,0)', '0.0%'],
    ['Meta Ads leads', cnt(['source', '"Meta Ads"']), '#,##0'],
    ['Meta Ads reassigned by round robin', cnt(['reassigned', '1']), '#,##0'],
    ['% Meta Ads reassigned', 'IFERROR(C24/C23,0)', '0.0%'],
  ];
  K.forEach(([l, f, nf], i) => {
    const r = 10 + i;
    wsSum.getCell(r, 2).value = l;
    wsSum.getCell(r, 3).value = { formula: f };
    wsSum.getCell(r, 3).numFmt = nf;
  });
  head(wsSum, 9, ['Hot stages', 'Leads'], 5);
  head(wsSum, 9, ['Lost stages', 'Leads'], 8);
  hotStages.forEach((s, i) => {
    wsSum.getCell(hotFirst + i, 5).value = s;
    wsSum.getCell(hotFirst + i, 6).value = { formula: cnt(['stage', `E${hotFirst + i}`]) };
  });
  lostStages.forEach((s, i) => {
    wsSum.getCell(hotFirst + i, 8).value = s;
    wsSum.getCell(hotFirst + i, 9).value = { formula: cnt(['stage', `H${hotFirst + i}`]) };
  });
  wsSum.getCell('B28').value = 'Leads outside the period are already removed from the data. Test accounts are excluded (see teams.json).';
  wsSum.getCell('B28').font = { name: 'Arial', size: 9, italic: true, color: { argb: GREY } };
  widths(wsSum, [3, 38, 20, 3, 22, 10, 3, 20, 10]);

  // =========================================================== By Team
  titleBlock(wsTeam, 'Holding now = Agent Name on the team. Created by team = Create By agent on the team. Team view = either of the two.');
  head(wsTeam, 4, ['Team', 'Leader', 'Holding now', 'Created by team (assigned at creation)', 'Team view (holding or created)', 'With first contact',
    '% with first contact', 'Hot leads', 'Hot rate', 'Lost leads', 'Stage: Not Contacted', 'Avg mins to first contact']);
  teamNames.forEach((t, i) => {
    const r = 5 + i, a = `$A${r}`;
    wsTeam.getCell(r, 1).value = t; wsTeam.getCell(r, 2).value = teamLeader(t);
    const F = {
      3: cnt(['team', a]), 4: cnt(['cteam', a]),
      5: `C${r}+D${r}-${cnt(['team', a], ['cteam', a])}`,
      6: cnt(['team', a], ['contacted', '1']), 7: `IFERROR(F${r}/C${r},0)`,
      8: `SUMPRODUCT(${cnt(['team', a], ['stage', hotRange])})`, 9: `IFERROR(H${r}/C${r},0)`,
      10: `SUMPRODUCT(${cnt(['team', a], ['stage', lostRange])})`,
      11: cnt(['team', a], ['stage', '"Not Contacted"']), 12: avg(['team', a], ['contacted', '1']),
    };
    Object.entries(F).forEach(([c, f]) => { wsTeam.getCell(r, +c).value = { formula: f }; });
  });
  const tt = 5 + teamNames.length;
  wsTeam.getCell(tt, 1).value = 'TOTAL';
  [3, 4, 6, 8, 10, 11].forEach((c) => { const L = wsTeam.getColumn(c).letter; wsTeam.getCell(tt, c).value = { formula: `SUM(${L}5:${L}${tt - 1})` }; });
  wsTeam.getCell(tt, 5).value = '-';
  wsTeam.getCell(tt, 7).value = { formula: `IFERROR(F${tt}/C${tt},0)` };
  wsTeam.getCell(tt, 9).value = { formula: `IFERROR(H${tt}/C${tt},0)` };
  wsTeam.getCell(tt, 12).value = { formula: avg(['contacted', '1']) };
  totalRow(wsTeam, tt, 12);
  fmtCols(wsTeam, 5, tt, [7, 9], '0.0%'); fmtCols(wsTeam, 5, tt, [12], '#,##0');
  widths(wsTeam, [22, 22, 12, 18, 18, 14, 14, 10, 10, 10, 14, 16]);
  wsTeam.views = [{ state: 'frozen', ySplit: 4, xSplit: 1 }];

  // =========================================================== By Agent
  titleBlock(wsAgent, 'Leads = holding now (Agent Name). Created by = assigned at creation (Create By). Agents who only created leads still get a row.');
  head(wsAgent, 4, ['Agent', 'Team', 'Leads (holding now)', 'Created by (assigned at creation)', 'With first contact', '% with first contact',
    'Hot leads', 'Lost leads', 'Stage: Not Contacted', 'Avg mins to first contact']);
  agentList.forEach((a, i) => {
    const r = 5 + i, g = `$A${r}`;
    wsAgent.getCell(r, 1).value = a;
    const F = {
      2: teamLookup(`A${r}`), 3: cnt(['agent', g]), 4: cnt(['by', g]), 5: cnt(['agent', g], ['contacted', '1']),
      6: `IFERROR(E${r}/C${r},0)`,
      7: `SUMPRODUCT(${cnt(['agent', g], ['stage', hotRange])})`, 8: `SUMPRODUCT(${cnt(['agent', g], ['stage', lostRange])})`,
      9: cnt(['agent', g], ['stage', '"Not Contacted"']), 10: avg(['agent', g], ['contacted', '1']),
    };
    Object.entries(F).forEach(([c, f]) => { wsAgent.getCell(r, +c).value = { formula: f }; });
  });
  const at = 5 + agentList.length;
  wsAgent.getCell(at, 1).value = 'TOTAL';
  [3, 4, 5, 7, 8, 9].forEach((c) => { const L = wsAgent.getColumn(c).letter; wsAgent.getCell(at, c).value = { formula: `SUM(${L}5:${L}${at - 1})` }; });
  wsAgent.getCell(at, 6).value = { formula: `IFERROR(E${at}/C${at},0)` };
  wsAgent.getCell(at, 10).value = { formula: avg(['contacted', '1']) };
  totalRow(wsAgent, at, 10);
  fmtCols(wsAgent, 5, at, [6], '0.0%'); fmtCols(wsAgent, 5, at, [10], '#,##0');
  widths(wsAgent, [28, 18, 14, 18, 14, 14, 10, 10, 14, 16]);
  wsAgent.views = [{ state: 'frozen', ySplit: 4, xSplit: 1 }];
  // note: "Created by" total can be lower than "holding now" total when creators are test accounts or outside the period

  // =========================================================== Team-Agent x Stage
  titleBlock(wsStage, 'Leads holding now, by stage.');
  head(wsStage, 4, ['Team', 'Agent', ...stages, 'Total']);
  activeAgents.forEach((a, i) => {
    const r = 5 + i;
    wsStage.getCell(r, 1).value = { formula: teamLookup(`B${r}`) };
    wsStage.getCell(r, 2).value = a;
    stages.forEach((s, j) => {
      const L = wsStage.getColumn(3 + j).letter;
      wsStage.getCell(r, 3 + j).value = { formula: cnt(['agent', `$B${r}`], ['stage', `${L}$4`]) };
    });
    const f = wsStage.getColumn(3).letter, l = wsStage.getColumn(2 + stages.length).letter;
    wsStage.getCell(r, 3 + stages.length).value = { formula: `SUM(${f}${r}:${l}${r})` };
  });
  const st = 5 + activeAgents.length;
  wsStage.getCell(st, 1).value = 'TOTAL';
  for (let c = 3; c <= 3 + stages.length; c++) { const L = wsStage.getColumn(c).letter; wsStage.getCell(st, c).value = { formula: `SUM(${L}5:${L}${st - 1})` }; }
  totalRow(wsStage, st, 3 + stages.length);
  widths(wsStage, [18, 28, ...stages.map(() => 13), 10]);
  wsStage.views = [{ state: 'frozen', ySplit: 4, xSplit: 2 }];

  // =========================================================== Contact Timing
  titleBlock(wsTime, 'Time from Create Time to First Contact Date, by team (holding now).');
  head(wsTime, 4, ['Team', 'Leads', 'With first contact', '5 min or less', '5 to 30 min', '30 to 60 min', '1 to 4 h', '4 to 24 h', 'Over 24 h',
    'No first contact', '% within 30 min']);
  teamNames.forEach((t, i) => {
    const r = 5 + i, a = `$A${r}`, T = ['team', a], Rc = ['contacted', '1'];
    wsTime.getCell(r, 1).value = t;
    const F = {
      2: cnt(T), 3: cnt(T, Rc),
      4: cnt(T, Rc, ['mins', '"<=5"']),
      5: cnt(T, Rc, ['mins', '">5"'], ['mins', '"<=30"']),
      6: cnt(T, Rc, ['mins', '">30"'], ['mins', '"<=60"']),
      7: cnt(T, Rc, ['mins', '">60"'], ['mins', '"<=240"']),
      8: cnt(T, Rc, ['mins', '">240"'], ['mins', '"<=1440"']),
      9: cnt(T, Rc, ['mins', '">1440"']),
      10: cnt(T, ['contacted', '0']), 11: `IFERROR((D${r}+E${r})/C${r},0)`,
    };
    Object.entries(F).forEach(([c, f]) => { wsTime.getCell(r, +c).value = { formula: f }; });
  });
  const tm = 5 + teamNames.length;
  wsTime.getCell(tm, 1).value = 'TOTAL';
  for (let c = 2; c <= 10; c++) { const L = wsTime.getColumn(c).letter; wsTime.getCell(tm, c).value = { formula: `SUM(${L}5:${L}${tm - 1})` }; }
  wsTime.getCell(tm, 11).value = { formula: `IFERROR((D${tm}+E${tm})/C${tm},0)` };
  totalRow(wsTime, tm, 11);
  fmtCols(wsTime, 5, tm, [11], '0.0%');
  widths(wsTime, [22, 10, 14, 12, 12, 12, 10, 10, 10, 14, 14]);
  wsTime.views = [{ state: 'frozen', ySplit: 4, xSplit: 1 }];

  // =========================================================== Agent x Source
  titleBlock(wsSrc, 'Leads holding now, by lead source.');
  head(wsSrc, 4, ['Team', 'Agent', ...sources, 'Total']);
  activeAgents.forEach((a, i) => {
    const r = 5 + i;
    wsSrc.getCell(r, 1).value = { formula: teamLookup(`B${r}`) };
    wsSrc.getCell(r, 2).value = a;
    sources.forEach((s, j) => {
      const L = wsSrc.getColumn(3 + j).letter;
      wsSrc.getCell(r, 3 + j).value = { formula: cnt(['agent', `$B${r}`], ['source', `${L}$4`]) };
    });
    const f = wsSrc.getColumn(3).letter, l = wsSrc.getColumn(2 + sources.length).letter;
    wsSrc.getCell(r, 3 + sources.length).value = { formula: `SUM(${f}${r}:${l}${r})` };
  });
  const sr = 5 + activeAgents.length;
  wsSrc.getCell(sr, 1).value = 'TOTAL';
  for (let c = 3; c <= 3 + sources.length; c++) { const L = wsSrc.getColumn(c).letter; wsSrc.getCell(sr, c).value = { formula: `SUM(${L}5:${L}${sr - 1})` }; }
  totalRow(wsSrc, sr, 3 + sources.length);
  widths(wsSrc, [18, 28, ...sources.map(() => 14), 10]);
  wsSrc.views = [{ state: 'frozen', ySplit: 4, xSplit: 2 }];

  // =========================================================== Daily Trend (data bars)
  titleBlock(wsDay, 'Leads created per day. The bars are Excel data bars and update with the numbers.');
  head(wsDay, 4, ['Date', 'Leads', 'Active leads', 'Lead pool', 'With first contact', 'Hot leads', 'Meta Ads']);
  const day0 = startOfDay(start), ndays = Math.min(400, Math.round((startOfDay(end) - day0) / 1440) + 1);
  for (let i = 0; i < ndays; i++) {
    const r = 5 + i, a = `$A${r}`;
    wsDay.getCell(r, 1).value = toDate(day0 + i * 1440);
    wsDay.getCell(r, 1).numFmt = 'ddd dd-mmm-yyyy';
    const F = {
      2: cnt(['cdate', a]), 3: cnt(['cdate', a], ['ls', '"Active Leads"']), 4: cnt(['cdate', a], ['ls', '"Leads Pool"']),
      5: cnt(['cdate', a], ['contacted', '1']), 6: `SUMPRODUCT(${cnt(['cdate', a], ['stage', hotRange])})`,
      7: cnt(['cdate', a], ['source', '"Meta Ads"']),
    };
    Object.entries(F).forEach(([c, f]) => { wsDay.getCell(r, +c).value = { formula: f }; });
  }
  const dt = 5 + ndays;
  wsDay.getCell(dt, 1).value = 'TOTAL';
  for (let c = 2; c <= 7; c++) { const L = wsDay.getColumn(c).letter; wsDay.getCell(dt, c).value = { formula: `SUM(${L}5:${L}${dt - 1})` }; }
  totalRow(wsDay, dt, 7);
  wsDay.addConditionalFormatting({
    ref: `B5:B${dt - 1}`,
    rules: [{ type: 'dataBar', gradient: false, cfvo: [{ type: 'num', value: 0 }, { type: 'max' }], color: { argb: BAR } }],
  });
  widths(wsDay, [20, 30, 14, 12, 18, 12, 12]);
  wsDay.views = [{ state: 'frozen', ySplit: 4 }];

  // =========================================================== Meta Ads Round Robin
  if (wsMeta) {
    titleBlock(wsMeta, 'Create By = agent the system assigned at creation. Agent Name = agent holding the lead now. A difference means round robin moved it.');
    head(wsMeta, 4, ['Meta Ads summary', 'Value']);
    const M = [
      ['Meta Ads leads', cnt(['source', '"Meta Ads"']), '#,##0'],
      ['Reassigned by round robin', cnt(['reassigned', '1']), '#,##0'],
      ['% reassigned', 'IFERROR(B6/B5,0)', '0.0%'],
      ['Avg mins to first contact (kept by original agent)', avg(['source', '"Meta Ads"'], ['reassigned', '0'], ['contacted', '1']), '#,##0'],
      ['Avg mins to first contact (reassigned)', avg(['source', '"Meta Ads"'], ['reassigned', '1'], ['contacted', '1']), '#,##0'],
      ['First contacted within 30 min', cnt(['source', '"Meta Ads"'], ['contacted', '1'], ['mins', '"<=30"']), '#,##0'],
      ['No first contact on record', cnt(['source', '"Meta Ads"'], ['contacted', '0']), '#,##0'],
    ];
    M.forEach(([l, f, nf], i) => {
      wsMeta.getCell(5 + i, 1).value = l; wsMeta.getCell(5 + i, 2).value = { formula: f }; wsMeta.getCell(5 + i, 2).numFmt = nf;
    });
    const sc0 = 14;
    wsMeta.getCell(sc0 - 1, 1).value = 'Agent scorecard (every roster member is listed, even with 0)';
    wsMeta.getCell(sc0 - 1, 1).font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY } };
    head(wsMeta, sc0, ['Agent', 'Team', 'Assigned at creation', 'Lost to round robin', 'Received by round robin', 'Holding now']);
    const metaAgents = agentList.slice();
    metaAgents.forEach((a, i) => {
      const r = sc0 + 1 + i, g = `$A${r}`, MA = ['source', '"Meta Ads"'];
      wsMeta.getCell(r, 1).value = a;
      const F = {
        2: teamLookup(`A${r}`), 3: cnt(MA, ['by', g]), 4: cnt(MA, ['by', g], ['reassigned', '1']),
        5: cnt(MA, ['agent', g], ['reassigned', '1']), 6: cnt(MA, ['agent', g]),
      };
      Object.entries(F).forEach(([c, f]) => { wsMeta.getCell(r, +c).value = { formula: f }; });
    });
    const mt = sc0 + 1 + metaAgents.length;
    wsMeta.getCell(mt, 1).value = 'TOTAL';
    for (let c = 3; c <= 6; c++) { const L = wsMeta.getColumn(c).letter; wsMeta.getCell(mt, c).value = { formula: `SUM(${L}${sc0 + 1}:${L}${mt - 1})` }; }
    totalRow(wsMeta, mt, 6);

    // static list of reassigned leads (no phone / email)
    const moved = cleanRows.filter((r) => r.source === 'Meta Ads' && r.by !== r.agent).sort((a, b) => b.ct - a.ct).slice(0, 300);
    const l0 = mt + 3;
    wsMeta.getCell(l0 - 1, 1).value = `Reassigned leads (${moved.length}${moved.length === 300 ? ', newest 300' : ''}). This list is a snapshot.`;
    wsMeta.getCell(l0 - 1, 1).font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY } };
    head(wsMeta, l0, ['Lead ID', 'Created', 'Create By (auto-assigned)', 'Agent Name (now)', 'Mins to first contact', 'Stage']);
    moved.forEach((r, i) => {
      const x = l0 + 1 + i;
      wsMeta.getCell(x, 1).value = r.id; wsMeta.getCell(x, 2).value = toDate(r.ct); wsMeta.getCell(x, 2).numFmt = DT;
      wsMeta.getCell(x, 3).value = r.by; wsMeta.getCell(x, 4).value = r.agent;
      if (r.fc) wsMeta.getCell(x, 5).value = Math.round(r.fc - r.ct);
      wsMeta.getCell(x, 6).value = r.stage;
    });
    widths(wsMeta, [48, 22, 22, 22, 22, 22]);
  }

  // ---- final pass: Arial everywhere, tab colours
  wb.worksheets.forEach((ws) => {
    ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
      if (!cell.font || cell.font.name !== 'Arial') cell.font = { ...(cell.font || {}), name: 'Arial', size: 10 };
    }));
  });
  [wsSum, wsTeam, wsAgent, wsStage, wsTime, wsSrc, wsDay].forEach((w) => { w.properties.tabColor = { argb: NAVY }; });
  if (wsMeta) wsMeta.properties.tabColor = { argb: 'FFF28C28' };

  return wb.xlsx.writeFile(outPath);
}

module.exports = { build };
