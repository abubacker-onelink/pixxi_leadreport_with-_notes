'use strict';
// Port of the skill's prepare_data.py: cleans the export, runs the QC checks and builds the
// data object D that the dashboard template expects.
const { toMinutes, startOfDay, fmtDay, fmtFull, fmtShort, fmtInput } = require('./dates');
const { redact } = require('./util');

const clean = (v) => String(v ?? '').replace(/\uFFFD/g, '').trim();
function valueCounts(list) {
  const m = new Map();
  list.forEach((v, i) => { if (!m.has(v)) m.set(v, { n: 0, i }); m.get(v).n++; });
  return [...m.entries()].sort((a, b) => b[1].n - a[1].n || a[1].i - b[1].i).map((e) => e[0]);
}
const indexOf = (arr) => new Map(arr.map((v, i) => [v, i]));

function projectOf(v) {
  const s = clean(v);
  if (!s || s === '|') return 'Not specified';
  const parts = s.split('|').map((x) => x.trim()).filter(Boolean);
  return parts.length === 1 ? parts[0] : parts.length ? 'Multiple types' : 'Not specified';
}

/**
 * @param rawRows rows from readExport (LS = 0 active, 1 pool)
 * @param o { start, end, export (minutes), cfg (teams.json), excludeIds:Set, source, title }
 */
function prepare(rawRows, o) {
  const cfg = o.cfg;
  const qc = { period: [fmtInput(o.start), fmtInput(o.end)], export_time: fmtInput(o.export), warnings: [] };
  qc.rows_read = { active: rawRows.filter((r) => !r.LS).length, pool: rawRows.filter((r) => r.LS).length };

  // ---- clean text fields
  let rows = [];
  let noId = 0;
  for (const r of rawRows) {
    const id = Number(r.ID);
    if (!Number.isFinite(id)) { noId++; continue; }
    rows.push({
      id,
      name: redact(clean(r['Lead Name'])),
      agent: clean(r['Agent Name']) || 'Unassigned',
      by: clean(r['Create By']) || 'Unknown',
      source: clean(r['Source Of Lead']) || 'Unknown',
      stage: clean(r.Stage) || 'No stage',
      camp: clean(r['Ad campaign Name']) || 'Not specified',
      note: redact(clean(r['Last Note'])),
      project: projectOf(r['Project Type']),
      budget: r.Budget === null || r.Budget === undefined ? null : redact(clean(r.Budget)),
      location: r['Preferred Location'] === null || r['Preferred Location'] === undefined ? null : redact(clean(r['Preferred Location'])),
      ct: toMinutes(r['Create Time']),
      fc: toMinutes(r['First Contact Date']),
      lf: toMinutes(r['Last Follow Up Time']),
      ls: r.LS,
    });
  }
  if (noId) qc.warnings.push(`${noId} rows had no valid ID and were skipped.`);

  // ---- duplicates
  const seen = new Set();
  const before = rows.length;
  rows = rows.filter((r) => (seen.has(r.id) ? false : (seen.add(r.id), true)));
  if (before - rows.length) qc.warnings.push(`${before - rows.length} duplicate lead IDs (kept the first occurrence).`);

  // ---- dates
  const bad = rows.filter((r) => r.ct === null).length;
  if (bad) qc.warnings.push(`${bad} leads have no valid Create Time and were excluded.`);
  const cts = rows.filter((r) => r.ct !== null).map((r) => r.ct);
  qc.create_time_range_in_files = cts.length ? [fmtInput(Math.min(...cts)), fmtInput(Math.max(...cts))] : [null, null];
  const inside = rows.filter((r) => r.ct !== null && r.ct >= o.start && r.ct <= o.end);
  qc.leads_in_period = inside.length;
  qc.leads_outside_period_excluded = rows.filter((r) => r.ct !== null).length - inside.length;
  rows = inside;

  // ---- test leads
  const exIds = new Set([...(cfg.exclude_ids || []).map(Number), ...(o.excludeIds || [])]);
  const exAgents = new Set(cfg.exclude_agents || []);
  const dropped = rows.filter((r) => exIds.has(r.id) || exAgents.has(r.agent));
  qc.test_leads_excluded = dropped.length;
  rows = rows.filter((r) => !(exIds.has(r.id) || exAgents.has(r.agent)));
  const pat = new RegExp('\\b(?:' + (cfg.test_markers || ['test']).map((m) => m.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|') + ')\\b', 'i');
  const sus = rows.filter((r) => pat.test(r.name) || pat.test(r.note));
  qc.possible_test_leads_still_included = sus.slice(0, 20).map((r) => ({ id: r.id, name: r.name, agent: r.agent, note: r.note.slice(0, 60) }));
  if (sus.length) qc.warnings.push(`${sus.length} leads look like test leads (name/note contains "test"). Check them and pass their IDs as extra excluded IDs if needed.`);

  // ---- teams
  const teams = cfg.teams;
  const a2t = new Map();
  teams.forEach((t, i) => t.members.forEach((m) => a2t.set(m, i)));
  const um = teams.length;
  const agents = valueCounts(rows.map((r) => r.agent));
  const unmapped = agents.filter((a) => !a2t.has(a));
  qc.unmapped_agents = Object.fromEntries(unmapped.map((a) => [a, rows.filter((r) => r.agent === a).length]));
  if (unmapped.length) qc.warnings.push('Agents not in teams.json (shown as Unmapped): ' + unmapped.join(', '));
  const agentSet = new Set(agents);
  qc.roster_members_with_zero_leads = teams.flatMap((t) => t.members).filter((m) => !agentSet.has(m));
  const agentTeam = agents.map((a) => (a2t.has(a) ? a2t.get(a) : um));
  const teamObjs = teams.map((t) => ({ name: t.name, leader: t.leader, color: t.color, members: t.members }))
    .concat([{ name: 'Unmapped', leader: '-', color: '#8A96A3' }]);

  // ---- dimensions
  const sources = valueCounts(rows.map((r) => r.source));
  const order = cfg.stage_order || [];
  const present = new Set(rows.map((r) => r.stage));
  const stages = order.filter((s) => present.has(s)).concat([...new Set(rows.map((r) => r.stage))].filter((s) => !order.includes(s)));
  const hot = new Set(cfg.hot_stages), lost = new Set(cfg.lost_stages);
  const stageGroup = stages.map((s) => (hot.has(s) ? 2 : lost.has(s) ? 3 : s === 'Not Contacted' ? 0 : 1));
  const pr = valueCounts(rows.map((r) => r.project)).filter((p) => p !== 'Multiple types' && p !== 'Not specified');
  const projSet = new Set(rows.map((r) => r.project));
  const projects = pr.concat(['Multiple types', 'Not specified'].filter((p) => projSet.has(p)));
  const creators = [...new Set(rows.map((r) => r.by))].sort();
  const campaigns = valueCounts(rows.map((r) => r.camp));
  const ai = indexOf(agents), si = indexOf(sources), sti = indexOf(stages), pi = indexOf(projects), ci = indexOf(creators), cmi = indexOf(campaigns);

  const dRows = rows.map((r) => [
    r.id, r.name, ai.get(r.agent), si.get(r.source), sti.get(r.stage), pi.get(r.project),
    r.ct, r.fc ?? 0, r.lf ?? 0, r.note.slice(0, 200), r.ls, ci.get(r.by), cmi.get(r.camp),
  ]);

  const day0 = startOfDay(o.start);
  const ndays = Math.round((startOfDay(o.end) - day0) / 1440) + 1;
  if (ndays > 93) qc.warnings.push(`Period is ${ndays} days; the daily chart will be crowded. Consider a shorter range.`);
  const label = day0 === startOfDay(o.end) ? fmtDay(o.start) : `${fmtDay(o.start)} to ${fmtDay(o.end)}`;
  const hasPool = rawRows.some((r) => r.LS);
  const D = {
    agents, agentTeam, teams: teamObjs, sources, stages, stageGroup, projects,
    statuses: ['Active leads', 'Buy leads from lead pool'], creators, campaigns,
    day0, ndays, now: o.export, rows: dRows,
    title: `${o.title || 'Buy leads report'}: ${label}`, hasPool, presetSource: o.source || '',
    exportLabel: fmtShort(o.export),
    subtitle: `All leads created ${fmtFull(o.start)} to ${fmtFull(o.end)}, exported ${fmtFull(o.export)}. ` +
      (hasPool ? 'This includes the buy leads from lead pool, which you can filter on their own. ' : '') +
      'Click any bar, day, team or row to filter everything else.',
  };

  const meta = rows.filter((r) => r.source === 'Meta Ads');
  qc.leads_in_report = rows.length;
  qc.by_status = { active: rows.filter((r) => !r.ls).length, pool: rows.filter((r) => r.ls).length };
  qc.meta_ads = { leads: meta.length, reassigned: meta.filter((r) => r.by !== r.agent).length };
  qc.non_meta_create_by_differs_from_agent = rows.filter((r) => r.source !== 'Meta Ads' && r.by !== r.agent).length;

  return { D, cleanRows: rows, qc };
}

function buildHtml(template, D) {
  const data = JSON.stringify(D).replace(/\uFFFD/g, '').replace(/<\//g, '<\\/');
  if (!template.includes('/*DATA*/')) throw new Error('Template is missing the /*DATA*/ placeholder.');
  return template.replace('/*DATA*/', () => data);
}

module.exports = { prepare, buildHtml };
