'use strict';
const express = require('express');
const multer = require('multer');
const os = require('os');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const { runReport } = require('./pipeline');
const { runNotesReport } = require('./notes/runNotesReport');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(ROOT, 'output');
const upload = multer({ dest: path.join(os.tmpdir(), 'leads-upload'), limits: { fileSize: 100 * 1024 * 1024 } });
const app = express();

app.use(express.static(path.join(ROOT, 'public')));
app.use('/output', express.static(OUT));

// ---- one upload -> dashboard + Excel now, notes report as a background job (Pixi allows ~1 request / 2 s)
const jobs = new Map();
const rel = (f) => '/output/' + path.relative(OUT, f).split(path.sep).map(encodeURIComponent).join('/');

async function runNotesJob(job, xlsxPath, outDir, leadIds, period) {
  try {
    const r = await runNotesReport(null, xlsxPath, {
      fetchFromPixi: true, leadIds, outRoot: outDir, periodLabel: period,
      title: 'Lead Notes Report',
      onProgress: (done, total) => { job.done = done; job.total = total; },
      log: (m) => console.log('[notes]', m),
    });
    job.status = 'done';
    job.qc = r.qc;
    job.file = { name: path.basename(r.file), url: rel(r.file) };
  } catch (e) {
    console.error(e);
    job.status = 'error'; job.error = e.message;
  } finally { fs.unlink(xlsxPath, () => {}); }
}

app.post('/api/generate', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Please choose an Excel (.xlsx) file.' });
  if (!/\.xlsx$/i.test(req.file.originalname)) {
    fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: 'Please upload an .xlsx file (the Pixi Buy Leads export).' });
  }
  let keepFile = false;
  try {
    const b = req.body || {};
    const result = await runReport(req.file.path, {
      start: b.start || undefined, end: b.end || undefined, source: b.source || '', excludeIds: b.excludeIds || '',
      outRoot: OUT, log: (m) => console.log('[report]', m),
    });
    const out = {
      qc: result.qc, period: result.period, source: result.source, outDir: result.outDir,
      files: Object.fromEntries(Object.entries(result.files).map(([k, v]) => [k, { url: rel(v), name: path.basename(v) }])),
    };
    // the Pixi token only ever lives on the server (PIXXI_TOKEN); it is never sent to or read from the browser
    if (!process.env.PIXXI_TOKEN) {
      out.notes = { status: 'error', error: 'Notes were skipped: the server has no PIXXI_TOKEN set. Ask the admin to set it and restart.' };
    } else {
      const id = crypto.randomBytes(8).toString('hex');
      const job = { status: 'running', done: 0, total: result.leadIds.length };
      jobs.set(id, job);
      keepFile = true;                       // the notes job still needs the uploaded file
      runNotesJob(job, req.file.path, result.outDir, result.leadIds, `${result.period[0]} to ${result.period[1]}`);
      out.notes = { jobId: id, ...job };
    }
    res.json(out);
  } catch (e) {
    console.error(e);
    res.status(422).json({ error: e.message });
  } finally {
    if (!keepFile) fs.unlink(req.file.path, () => {});
  }
});

app.get('/api/jobs/:id', (req, res) => {
  const j = jobs.get(req.params.id);
  if (!j) return res.status(404).json({ error: 'Unknown job.' });
  res.json(j);
});

app.get('/notes', (req, res) => res.sendFile(path.join(ROOT, 'public', 'notes.html')));

app.post('/api/notes', upload.single('leads'), async (req, res) => {
  const leads = req.file;
  const cleanup = () => leads && fs.unlink(leads.path, () => {});
  if (!leads) return res.status(400).json({ error: 'Please choose the Pixi leads Excel file.' });
  if (!/\.xlsx$/i.test(leads.originalname)) {
    cleanup(); return res.status(400).json({ error: 'Please upload a .xlsx leads file.' });
  }
  try {
    let newAuthors = {};
    try { newAuthors = JSON.parse((req.body && req.body.authors) || '{}'); } catch { /* ignore */ }
    const r = await runNotesReport(null, leads.path, {
      fetchFromPixi: true, title: (req.body && req.body.title) || undefined, outRoot: OUT, newAuthors,
      log: (m) => console.log('[notes]', m),
    });
    res.json({ qc: r.qc, kind: r.kind, outDir: r.outDir,
      file: { name: path.basename(r.file), url: '/output/' + path.relative(OUT, r.file).split(path.sep).map(encodeURIComponent).join('/') } });
  } catch (e) {
    console.error(e);
    res.status(422).json({ error: e.message });
  } finally { cleanup(); }
});

function openBrowser(url) {
  const cmd = process.platform === 'win32' ? `start "" "${url}"` : process.platform === 'darwin' ? `open "${url}"` : `xdg-open "${url}"`;
  exec(cmd, () => {});
}

function listen(port, tries = 0) {
  const server = app.listen(port, '127.0.0.1', () => {
    const url = `http://localhost:${port}`;
    console.log(`Buy Leads Dashboard is running at ${url}`);
    console.log('Upload the Pixi Buy Leads Excel file in the page. Press Ctrl+C to stop.');
    if (!process.env.NO_OPEN) openBrowser(url);
  });
  server.on('error', (e) => {
    if (e.code === 'EADDRINUSE' && tries < 10) listen(port + 1, tries + 1);
    else throw e;
  });
}

if (require.main === module) listen(Number(process.env.PORT) || 3000);
module.exports = { app };
