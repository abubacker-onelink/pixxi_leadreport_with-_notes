'use strict';
// All times are Dubai wall-clock times stored as "minutes since 1970-01-01" with NO timezone conversion
// (same convention the dashboard template expects).
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const pad = (n) => String(n).padStart(2, '0');

function fromParts(y, mo, d, h = 0, mi = 0) {
  const t = Date.UTC(y, mo - 1, d, h, mi);
  const dt = new Date(t);
  // reject things like 31-02
  if (dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return null;
  return t / 60000;
}

/** Accepts Date, Excel serial number, "DD-MM-YYYY HH:MM" or ISO "YYYY-MM-DD HH:MM". Returns minutes or null. */
function toMinutes(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) {
    if (isNaN(v)) return null;
    return Math.floor(Math.round(v.getTime() / 1000) / 60);
  }
  if (typeof v === 'number') return Number.isFinite(v) ? Math.floor(Math.round((v - 25569) * 86400) / 60) : null;
  const s = String(v).trim();
  let m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::\d{2})?)?$/.exec(s);
  if (m) return fromParts(+m[3], +m[2], +m[1], +(m[4] || 0), +(m[5] || 0));
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::\d{2})?)?/.exec(s);
  if (m) return fromParts(+m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0));
  return null;
}

const toDate = (min) => new Date(min * 60000); // read it with getUTC* methods
const startOfDay = (min) => Math.floor(min / 1440) * 1440;

function dparts(min) {
  const d = toDate(min);
  return { y: d.getUTCFullYear(), mo: d.getUTCMonth(), d: d.getUTCDate(), h: d.getUTCHours(), mi: d.getUTCMinutes() };
}
/** "13 Oct 2026" (no leading zero on the day) */
function fmtDay(min) { const p = dparts(min); return `${p.d} ${MON[p.mo]} ${p.y}`; }
/** "13 Oct 2026 09:40" (zero padded day) */
function fmtFull(min) { const p = dparts(min); return `${pad(p.d)} ${MON[p.mo]} ${p.y} ${pad(p.h)}:${pad(p.mi)}`; }
/** "13 Oct 09:40" */
function fmtShort(min) { const p = dparts(min); return `${p.d} ${MON[p.mo]} ${pad(p.h)}:${pad(p.mi)}`; }
/** "DD-MM-YYYY HH:MM" */
function fmtInput(min) { const p = dparts(min); return `${pad(p.d)}-${pad(p.mo + 1)}-${p.y} ${pad(p.h)}:${pad(p.mi)}`; }
/** "Tue 14 Oct, 09:40" */
function fmtDow(min) { const p = dparts(min); return `${DOW[toDate(min).getUTCDay()]} ${pad(p.d)} ${MON[p.mo]}, ${pad(p.h)}:${pad(p.mi)}`; }
function fmtIso(min) { const p = dparts(min); return `${p.y}-${pad(p.mo + 1)}-${pad(p.d)}`; }

function dubaiNow() {
  const t = Date.now() + 4 * 3600 * 1000; // UTC+4
  return Math.floor(t / 60000);
}

module.exports = { toMinutes, toDate, startOfDay, fmtDay, fmtFull, fmtShort, fmtInput, fmtDow, fmtIso, dubaiNow, MON, pad };
