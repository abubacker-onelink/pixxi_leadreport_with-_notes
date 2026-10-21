'use strict';

const { parseNotesJson } = require('./readNotes');

const ENDPOINT = 'https://dataapi.pixxicrm.ae/pixxiapi/v1/lead/activity-logs';
const PAGE_SIZE = 100;
const MIN_GAP_MS = 2200;   // Pixi answers "Too frequent requests - wait 2 seconds" otherwise
const MAX_RETRIES = 4;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function fetchPixiNotes(leadIds, options = {}) {
  const token = options.token || process.env.PIXXI_TOKEN;
  if (!token) throw new Error('Pixi API token is required. Set PIXXI_TOKEN before starting the dashboard.');
  const fetchImpl = options.fetchImpl || globalThis.fetch;
  if (typeof fetchImpl !== 'function') throw new Error('This Node.js version does not provide fetch(). Use Node.js 18 or later.');

  const endpoint = options.endpoint || ENDPOINT;
  const log = options.log || (() => {});
  const leads = new Map();
  // tests inject fetchImpl and pass minGapMs: 0; real calls are throttled
  const minGap = options.minGapMs !== undefined ? options.minGapMs : (options.fetchImpl ? 0 : MIN_GAP_MS);
  let lastCall = 0;
  const post = async (body) => {
    for (let attempt = 0; ; attempt++) {
      const wait = lastCall + minGap - Date.now();
      if (wait > 0) await sleep(wait);
      lastCall = Date.now();
      const response = await fetchImpl(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-PIXXI-TOKEN': token },
        body: JSON.stringify(body),
      });
      let payload;
      try { payload = await response.json(); } catch { payload = undefined; }
      const msg = payload && payload.message ? String(payload.message) : '';
      if (/too frequent|rate limit/i.test(msg) && attempt < MAX_RETRIES) {
        log(`Pixi rate limit hit, waiting before retry ${attempt + 1}...`);
        await sleep(2500);
        continue;
      }
      return { response, payload };
    }
  };

  const ids = [...leadIds];
  let done = 0;
  for (const leadId of ids) {
    const notes = [];
    let page = 1;
    let totalPages = 1;
    do {
      log(`Fetching notes for lead ${leadId} (page ${page})...`);
      const { response, payload } = await post({ leadId, type: 'Note', page, size: PAGE_SIZE });
      if (payload === undefined) {
        throw new Error(`Pixi returned an invalid JSON response for lead ${leadId} (HTTP ${response.status}).`);
      }
      if (!response.ok) {
        throw new Error(`Pixi notes request failed for lead ${leadId} (HTTP ${response.status}).`);
      }
      if (payload && payload.statusCode !== undefined) {
        const statusCode = String(payload.statusCode).trim().toLowerCase();
        if (!['0', '200', 'success'].includes(statusCode)) {
          throw new Error(`Pixi notes request failed for lead ${leadId}: ${payload.message || `status ${payload.statusCode}`}.`);
        }
      }
      const data = payload && payload.data;
      if (!data || !Array.isArray(data.list)) {
        throw new Error(`Pixi returned an unexpected activity-log response for lead ${leadId}: expected data.list.`);
      }
      for (const item of data.list) {
        const parsed = parseNotesJson(JSON.stringify([item]));
        if (!parsed || !parsed.length || !parsed[0].message) {
          throw new Error(`Pixi returned an activity log without a recognized note-text field for lead ${leadId}.`);
        }
        notes.push(parsed[0]);
      }
      const pages = Number(data.totalPages);
      if (Number.isFinite(pages) && pages > 0) totalPages = pages;
      else if (data.list.length === PAGE_SIZE) totalPages = page + 1;
      else totalPages = page;
      page++;
    } while (page <= totalPages);
    leads.set(Number(leadId), { id: Number(leadId), fetchStatus: 'OK', notes });
    if (options.onProgress) options.onProgress(++done, ids.length);
  }

  return { leads, warnings: [], statusCounts: { OK: leads.size } };
}

module.exports = { fetchPixiNotes, ENDPOINT, PAGE_SIZE };
