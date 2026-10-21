#!/usr/bin/env node
'use strict';
const { runNotesReport } = require('./notes/runNotesReport');

(async () => {
  const a = process.argv.slice(2);
  const get = (f) => { const i = a.indexOf(f); return i >= 0 ? a[i + 1] : undefined; };
  const notes = get('--notes'), leads = get('--leads');
  if (!notes || !leads) {
    console.log('Usage: node src/notes-cli.js --notes notes.csv --leads leads.xlsx [--title "Buy leads notes"] [--out folder]');
    return;
  }
  try {
    const res = await runNotesReport(notes, leads, { title: get('--title'), outRoot: get('--out'), log: console.log });
    console.log('\nQC summary:\n' + JSON.stringify(res.qc, null, 2));
    console.log('\nSaved: ' + res.file);
  } catch (e) { console.error('Error: ' + e.message); process.exit(1); }
})();
