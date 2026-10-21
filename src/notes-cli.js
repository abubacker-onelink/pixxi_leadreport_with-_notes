#!/usr/bin/env node
'use strict';
const { runNotesReport } = require('./notes/runNotesReport');

(async () => {
  const a = process.argv.slice(2);
  const get = (f) => { const i = a.indexOf(f); return i >= 0 ? a[i + 1] : undefined; };
  const notes = get('--notes'), leads = get('--leads');
  const fetchFromPixi = a.includes('--fetch');
  if (!leads || (!notes && !fetchFromPixi)) {
    console.log('Usage: node src/notes-cli.js --leads leads.xlsx (--notes notes.csv | --fetch) [--title "Buy leads notes"] [--out folder]');
    console.log('  --fetch  pulls Note logs from the Pixi API for every lead ID in the leads file (needs PIXXI_TOKEN)');
    return;
  }
  try {
    const res = await runNotesReport(fetchFromPixi ? null : notes, leads, { fetchFromPixi, title: get('--title'), outRoot: get('--out'), log: console.log });
    console.log('\nQC summary:\n' + JSON.stringify(res.qc, null, 2));
    console.log('\nSaved: ' + res.file);
  } catch (e) { console.error('Error: ' + e.message); process.exit(1); }
})();
