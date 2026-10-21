#!/usr/bin/env node
'use strict';
const { runReport } = require('./pipeline');

function usage() {
  console.log(`Usage: node src/cli.js <export.xlsx> [--start "DD-MM-YYYY HH:MM"] [--end "DD-MM-YYYY HH:MM"]
                          [--export "DD-MM-YYYY HH:MM"] [--source "Meta Ads"] [--exclude-ids 1,2] [--out folder]
With no file, run "npm start" instead to get the upload page.`);
}

(async () => {
  const args = process.argv.slice(2);
  if (!args.length || args.includes('-h') || args.includes('--help')) return usage();
  const opt = {};
  let file = null;
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--start') opt.start = args[++i];
    else if (a === '--end') opt.end = args[++i];
    else if (a === '--export') opt.export = args[++i];
    else if (a === '--source') opt.source = args[++i];
    else if (a === '--exclude-ids') opt.excludeIds = args[++i];
    else if (a === '--out') opt.outRoot = args[++i];
    else if (!a.startsWith('--')) file = a;
  }
  if (!file) return usage();
  try {
    const res = await runReport(file, { ...opt, log: (m) => console.log(m) });
    console.log('\nQC summary:\n' + JSON.stringify(res.qc, null, 2));
    console.log('\nOutput folder: ' + res.outDir);
    Object.values(res.files).forEach((f) => console.log('  ' + f));
  } catch (e) {
    console.error('Error: ' + e.message);
    process.exit(1);
  }
})();
