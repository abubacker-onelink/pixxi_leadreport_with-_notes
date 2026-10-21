'use strict';
const express = require('express');
const multer = require('multer');
const os = require('os');
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

app.post('/api/generate', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Please choose an Excel (.xlsx) file.' });
  if (!/\.xlsx$/i.test(req.file.originalname)) {
    fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: 'Please upload an .xlsx file (the Pixi Buy Leads export).' });
  }
  try {
    const b = req.body || {};
    const result = await runReport(req.file.path, {
      start: b.start || undefined, end: b.end || undefined, source: b.source || '', excludeIds: b.excludeIds || '',
      outRoot: OUT, log: (m) => console.log('[report]', m),
    });
    const rel = (f) => '/output/' + path.relative(OUT, f).split(path.sep).map(encodeURIComponent).join('/');
    res.json({
      qc: result.qc, period: result.period, source: result.source, outDir: result.outDir,
      files: Object.fromEntries(Object.entries(result.files).map(([k, v]) => [k, { url: rel(v), name: path.basename(v) }])),
    });
  } catch (e) {
    console.error(e);
    res.status(422).json({ error: e.message });
  } finally {
    fs.unlink(req.file.path, () => {});
  }
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
