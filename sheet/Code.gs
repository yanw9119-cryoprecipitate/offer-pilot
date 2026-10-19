/**
 * Offer Pilot data sheet
 * ----------------------
 * 1. Live Melbourne Airport flights. This script calls AeroDataBox on a timer and keeps the next 7 days of
 *    flights in the "Flights" tab. Every Offer Pilot (any phone, tablet or computer) reads them from here,
 *    so the pilots use none of your RapidAPI calls. The key is kept in this script's settings, never in the app.
 * 2. Your shift log. When you punch out, the pilot sends the shift here: one tab per month
 *    ("2026-10 October"), one row per shift, and a "Summary" tab with a line for each month.
 *
 * Setup (once, on a computer): see the "Offer Pilot" menu at the top of the sheet.
 *   ✈️ Offer Pilot ▸ 1. Set flight key   (paste your X-RapidAPI-Key)
 *   ✈️ Offer Pilot ▸ 2. Start hourly updates
 *   Deploy ▸ New deployment ▸ Web app · Execute as: Me · Who has access: Anyone ▸ Deploy
 *   ✈️ Offer Pilot ▸ 3. Phone link       (open it on your phone once)
 */
const TZ = 'Australia/Melbourne';
const CODE = 'MEL';
const HOST = 'aerodatabox.p.rapidapi.com';
const AHEAD = 14;                    // 14 blocks of 12 hours = the next 7 days
const REFRESH_H = 3;                 // the block you're in is refreshed about every 3 hours (as the allowance allows)
const PILOT = 'https://yanw9119-cryoprecipitate.github.io/offer-pilot/';
const P = PropertiesService.getScriptProperties();
const FL_HEAD = ['Block', 'Checked', 'Direction', 'Flight', 'Airline', 'From / to', 'Scheduled', 'Expected', 'Terminal', 'Status', 'International'];
const SH_HEAD = ['Date', 'Day', 'Punch in', 'Punch out', 'Hours', 'Car', 'Trips', 'Offers', 'Accepted', 'Declined', 'Ride types',
  'Fares $', 'Car costs $', 'Take-home $', '$ per hour', 'Target $', 'Target hit', 'km', 'Started in', 'Finished in', 'Areas', 'Route map', 'Offers list', 'Shift ID'];

// ---------------------------------------------------------------- menu
function onOpen() {
  SpreadsheetApp.getUi().createMenu('✈️ Offer Pilot')
    .addItem('1. Set flight key', 'setKey')
    .addItem('2. Start hourly updates', 'start')
    .addItem('3. Phone link', 'phoneLink')
    .addSeparator()
    .addItem('Update flights now', 'updateNow')
    .addItem('Status', 'status')
    .addItem('Stop updates', 'stop')
    .addToUi();
}

function setKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('AeroDataBox key', 'Paste your X-RapidAPI-Key from rapidapi.com. It is saved in this script\'s private settings, not in the sheet.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const k = r.getResponseText().trim();
  if (!/^[\w-]{20,}$/.test(k)) { ui.alert('That doesn\'t look like a RapidAPI key.'); return; }
  P.setProperty('AERO_KEY', k);
  const c = ui.prompt('Calls a month', 'How many calls may this sheet use a month? The free Basic plan is 400 units, 2 per call = 200. Leave empty for 190.', ui.ButtonSet.OK_CANCEL);
  if (c.getSelectedButton() === ui.Button.OK && +c.getResponseText() > 0) P.setProperty('CAP', String(Math.round(+c.getResponseText())));
  ui.alert('✓ Key saved. Next: ✈️ Offer Pilot ▸ 2. Start hourly updates.');
}

function start() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  token_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'tick').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tick').timeBased().everyHours(1).create();
  flightsSheet_(); summarySheet_();
  const msg = tick();
  SpreadsheetApp.getUi().alert('✓ Hourly updates are on.\n\n' + msg + '\n\nNext: Deploy ▸ New deployment ▸ Web app (Execute as: Me · Who has access: Anyone), then ✈️ Offer Pilot ▸ 3. Phone link.');
}

function stop() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'tick').forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert('Hourly flight updates stopped. Pilots keep reading the flights already in the sheet.');
}

function updateNow() { SpreadsheetApp.getUi().alert(tick(true)); }

function status() {
  const st = blockState_(), have = blocksAhead_().filter(s => st[s]).length;
  SpreadsheetApp.getUi().alert(`Flight key: ${P.getProperty('AERO_KEY') ? 'saved' : 'not set'}\nWeek ahead: ${have} of ${AHEAD} half-days loaded\n` +
    `Calls this month: ${calls_()} of ${cap_()}\nLast update: ${P.getProperty('LAST_RUN') ? fmt_(+P.getProperty('LAST_RUN'), 'EEE d MMM h:mm a') : 'never'}` +
    (P.getProperty('LAST_ERR') ? `\nLast problem: ${P.getProperty('LAST_ERR')}` : ''));
}

function phoneLink() {
  const ui = SpreadsheetApp.getUi();
  // Always the address you copy from Deploy ▸ Manage deployments (Google's own guess can point at an old or test deployment)
  let url = P.getProperty('WEB_URL') || '';
  const r = ui.prompt('Web app address', 'In Apps Script: Deploy ▸ Manage deployments ▸ copy the Web app URL (ends in /exec) and paste it here.' +
    (url ? '\nLeave empty to keep: ' + url : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  if (r.getResponseText().trim()) url = r.getResponseText().trim();
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(url)) { ui.alert('That isn\'t a web app address (https://script.google.com/macros/s/…/exec).'); return; }
  P.setProperty('WEB_URL', url);
  const mine = `${PILOT}#sheet=${encodeURIComponent(url)}&sheetkey=${token_()}`, flightsOnly = `${PILOT}#sheet=${encodeURIComponent(url)}`;
  const h = HtmlService.createHtmlOutput(
    `<div style="font:14px/1.45 -apple-system,Segoe UI,Roboto,sans-serif">` +
    `<p><b>Your phone link</b> (flights + shift log). Open it once on each of your devices. Keep it private, like a password: it lets a pilot write to this sheet.</p>` +
    `<p><a href="${mine}" target="_blank">Open in Offer Pilot</a></p><textarea style="width:100%;height:70px" onclick="this.select()">${mine}</textarea>` +
    `<p><b>Flights only</b> (safe to give to other drivers: reads flights, can't write anything).</p>` +
    `<textarea style="width:100%;height:52px" onclick="this.select()">${flightsOnly}</textarea>` +
    `<p style="color:#666">Tip: save the private link in your "Offer Pilot – private key setup link" Google Doc.</p></div>`).setWidth(460).setHeight(400);
  ui.showModalDialog(h, '✈️ Offer Pilot link');
}

// ---------------------------------------------------------------- flights
// Runs every hour. Each call fetches one 12-hour block (midnight–noon or noon–midnight, Melbourne time).
//  1. Missing blocks in the next 7 days, nearest first (up to 4 a run).
//  2. The next block once more when it's under 12 hours away (schedule changes, delays).
//  3. The block you're in, about every 3 hours from 5am to 1am, while the monthly allowance allows.
function tick(manual) {
  const key = P.getProperty('AERO_KEY');
  if (!key) return 'No flight key yet: ✈️ Offer Pilot ▸ 1. Set flight key.';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return 'Another update is running.';
  try {
    const st = blockState_(), starts = blocksAhead_(), now = Date.now(), todo = [];
    starts.filter(s => !st[s]).slice(0, 4).forEach(s => todo.push(s));
    const nx = starts[1];
    if (st[nx] && !todo.includes(nx) && now - st[nx] > 12 * 3600000) todo.push(nx);
    const cur = starts[0], h = +fmt_(now, 'H');
    if (st[cur] && !todo.includes(cur) && (manual || ((h >= 5 || h < 1) && now - st[cur] >= REFRESH_H * 3600000 - 5 * 60000 && budgetOk_()))) todo.push(cur);
    const got = {}; let err = '';
    for (const s of todo) {
      if (calls_() >= cap_() * 0.97) { err = 'Monthly allowance reached; flights already in the sheet stay.'; break; }
      try { got[s] = fetchBlock_(key, s); st[s] = Date.now(); } catch (e) { err = String(e.message || e); break; }
      Utilities.sleep(1200);
    }
    // forget blocks that ended more than 6 hours ago
    Object.keys(st).forEach(s => { if (blockEndMs_(s) < now - 6 * 3600000) delete st[s]; });
    writeFlights_(got, st);
    P.setProperty('BLOCKS', JSON.stringify(st));
    P.setProperty('LAST_RUN', String(Date.now()));
    if (err) P.setProperty('LAST_ERR', err); else P.deleteProperty('LAST_ERR');
    const have = starts.filter(s => st[s]).length;
    return `${Object.keys(got).length} block${Object.keys(got).length === 1 ? '' : 's'} updated · week ahead ${have} of ${AHEAD} half-days · ${calls_()} of ${cap_()} calls this month${err ? '\n⚠️ ' + err : ''}`;
  } finally { lock.releaseLock(); }
}

function fetchBlock_(key, start) {
  const end = start.slice(0, 11) + (start.slice(11, 13) === '00' ? '11:59' : '23:59');
  const url = `https://${HOST}/flights/airports/iata/${CODE}/${start}/${end}?withLeg=false&direction=Both` +
    '&withCancelled=true&withCodeshared=false&withCargo=false&withPrivate=false&withLocation=false';
  const r = UrlFetchApp.fetch(url, { headers: { 'X-RapidAPI-Key': key, 'X-RapidAPI-Host': HOST }, muteHttpExceptions: true });
  addCall_();
  const c = r.getResponseCode();
  if (c === 204) return { arr: [], dep: [] };
  if (c === 401 || c === 403) throw new Error(/not subscribed/i.test(r.getContentText()) ? 'Not subscribed to AeroDataBox on RapidAPI (subscribe to the free Basic plan).' : 'The key was refused (' + c + ').');
  if (c === 429) throw new Error('RapidAPI limit reached (429). Trying again next hour.');
  if (c !== 200) throw new Error('AeroDataBox answered ' + c + '.');
  const j = JSON.parse(r.getContentText() || '{}');
  return { arr: parse_(j.arrivals), dep: parse_(j.departures) };
}

function flTime_(x) {
  const s = x && (x.utc || x.local || (typeof x === 'string' ? x : ''));
  const t = s ? Date.parse(String(s).replace(' ', 'T')) : NaN;
  return isNaN(t) ? null : t;
}
function parse_(list) {
  return (list || []).map(f => {
    const m = f.movement || f.arrival || f.departure || {};
    const t = flTime_(m.scheduledTime || m.scheduledTimeUtc || m.scheduledTimeLocal);
    const rt = flTime_(m.revisedTime || m.actualTimeUtc || m.estimatedTimeUtc || m.revisedTimeUtc);
    const ap = m.airport || {}, term = m.terminal ? String(m.terminal).replace(/^T/i, '') : '';
    return { n: f.number || '', al: (f.airline || {}).name || '', ap: ap.shortName || ap.municipalityName || ap.name || ap.iata || '',
      t, rt: rt || null, term, st: f.status || '', intl: CODE === 'MEL' ? term === '2' : !!(ap.countryCode && ap.countryCode !== 'AU') };
  }).filter(f => f.t);
}

function flightsSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Flights');
  if (!sh) { sh = ss.insertSheet('Flights', 0); }
  if (sh.getLastRow() === 0 || sh.getRange(1, 1).getValue() !== FL_HEAD[0]) {
    sh.clear(); sh.getRange(1, 1, 1, FL_HEAD.length).setValues([FL_HEAD]).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  sh.getRange('A:A').setNumberFormat('@');          // block names stay text (not turned into dates)
  return sh;
}

// Replaces the rows of the blocks just fetched, drops old blocks, keeps everything sorted by expected time
function writeFlights_(got, st) {
  const sh = flightsSheet_(), n = sh.getLastRow() - 1;
  let rows = n > 0 ? sh.getRange(2, 1, n, FL_HEAD.length).getValues() : [];
  rows = rows.filter(r => st[String(r[0])] && !got[String(r[0])]);
  Object.keys(got).forEach(s => {
    const at = new Date(st[s]);
    [['Arrival', got[s].arr], ['Departure', got[s].dep]].forEach(([dir, list]) => list.forEach(f => rows.push(
      [s, at, dir, f.n, f.al, f.ap, new Date(f.t), f.rt ? new Date(f.rt) : '', f.term, f.st, f.intl ? 'Yes' : ''])));
  });
  rows.sort((a, b) => (a[7] || a[6]) - (b[7] || b[6]));
  if (n > 0) sh.getRange(2, 1, n, FL_HEAD.length).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, FL_HEAD.length).setValues(rows);
    sh.getRange(2, 2, rows.length, 1).setNumberFormat('ddd d mmm h:mm am/pm');
    sh.getRange(2, 7, rows.length, 2).setNumberFormat('ddd d mmm h:mm am/pm');
  }
  const extra = sh.getMaxRows() - Math.max(rows.length + 1, 2);
  if (extra > 200) sh.deleteRows(rows.length + 202, extra - 200);
}

// Melbourne-time blocks: '2026-10-18T00:00', '2026-10-18T12:00', …
function blocksAhead_() {
  const now = new Date(), day0 = fmt_(now, 'yyyy-MM-dd'), h = +fmt_(now, 'H'), out = [];
  let day = day0, half = h < 12 ? 0 : 12;
  for (let i = 0; i < AHEAD; i++) {
    out.push(`${day}T${half ? '12' : '00'}:00`);
    if (half) { half = 0; day = addDays_(day, 1); } else half = 12;
  }
  return out;
}
function addDays_(day, n) { const [y, m, d] = day.split('-').map(Number); return Utilities.formatDate(new Date(Date.UTC(y, m - 1, d + n, 12)), 'UTC', 'yyyy-MM-dd'); }
function blockEndMs_(s) {
  const [y, m, d] = s.slice(0, 10).split('-').map(Number), noon = s.slice(11, 13) === '12';
  // end of the block in Melbourne time (+10 or +11 hours from UTC); 11 is safe for "has it ended"
  return Date.UTC(y, m - 1, d, noon ? 24 : 12) - 10 * 3600000;
}
function blockState_() { try { return JSON.parse(P.getProperty('BLOCKS')) || {}; } catch (e) { return {}; } }
function fmt_(t, f) { return Utilities.formatDate(new Date(t), TZ, f); }
function cap_() { return +P.getProperty('CAP') || 190; }
function calls_() { try { const c = JSON.parse(P.getProperty('CALLS')); if (c && c.m === fmt_(Date.now(), 'yyyy-MM')) return c.n; } catch (e) {} return 0; }
function addCall_() { P.setProperty('CALLS', JSON.stringify({ m: fmt_(Date.now(), 'yyyy-MM'), n: calls_() + 1 })); }
// Spread the calls over the month: what may be used by now, plus a little head room
function budgetOk_() {
  const d = new Date(), y = +fmt_(d, 'yyyy'), m = +fmt_(d, 'M'), days = new Date(Date.UTC(y, m, 0)).getUTCDate();
  const share = cap_() * 0.9 * (+fmt_(d, 'd') - 1 + (+fmt_(d, 'H') + 1) / 24) / days;
  return calls_() < Math.min(share + 6, cap_() * 0.9);
}

// ---------------------------------------------------------------- web app
// GET  …/exec            → the flights as JSON, for every pilot (no key needed, read-only)
// POST …/exec {key, kind:'shift', shift} → writes a shift to its month tab (needs your private key)
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.ping) return json_({ ok: true, v: 1 });
  const sh = SpreadsheetApp.getActive().getSheetByName('Flights'), st = blockState_(), blocks = {};
  Object.keys(st).forEach(s => { blocks[s] = { at: st[s], arr: [], dep: [] }; });
  const n = sh ? sh.getLastRow() - 1 : 0;
  if (n > 0) sh.getRange(2, 1, n, FL_HEAD.length).getValues().forEach(r => {
    const b = blocks[String(r[0])]; if (!b || !(r[6] instanceof Date)) return;
    (r[2] === 'Arrival' ? b.arr : b.dep).push([r[3], r[4], r[5], r[6].getTime(), r[7] instanceof Date ? r[7].getTime() : 0, String(r[8] || ''), r[9], r[10] === 'Yes' ? 1 : 0]);
  });
  return json_({ ok: true, v: 1, code: CODE, at: +P.getProperty('LAST_RUN') || 0, calls: calls_(), cap: cap_(), err: P.getProperty('LAST_ERR') || '',
    fields: ['n', 'al', 'ap', 't', 'rt', 'term', 'st', 'intl'], blocks });
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { return json_({ ok: false, err: 'bad request' }); }
  if (!b.key || b.key !== P.getProperty('TOKEN')) return json_({ ok: false, err: 'key' });
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (b.kind === 'ping') return json_({ ok: true });
    if (b.kind === 'shift' && b.shift) { const tab = writeShift_(b.shift); return json_({ ok: true, tab }); }
    if (b.kind === 'shifts' && Array.isArray(b.shifts)) { const tabs = b.shifts.map(writeShift_); return json_({ ok: true, n: tabs.length }); }
    return json_({ ok: false, err: 'unknown' });
  } finally { lock.releaseLock(); }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function token_() {
  let t = P.getProperty('TOKEN');
  if (!t) { t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); P.setProperty('TOKEN', t); }
  return t;
}

// ---------------------------------------------------------------- shift log
function monthTab_(t) { return fmt_(t, 'yyyy-MM MMMM'); }
function monthSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    // month tabs go after Summary, newest first
    const months = ss.getSheets().map(s => s.getName()).filter(n => /^\d{4}-\d{2} /.test(n));
    const after = months.filter(n => n > name).length;
    const sumIdx = ss.getSheetByName('Summary') ? ss.getSheetByName('Summary').getIndex() : 1;
    sh = ss.insertSheet(name, sumIdx + after);
    sh.getRange(1, 1, 1, SH_HEAD.length).setValues([SH_HEAD]).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    sh.setFrozenRows(1); sh.setFrozenColumns(2);
    sh.setColumnWidth(21, 260); sh.setColumnWidth(23, 360); sh.hideColumns(SH_HEAD.length);
  }
  return sh;
}

function writeShift_(s) {
  const name = monthTab_(s.start), sh = monthSheet_(name), id = String(s.id || s.start);
  const st = new Date(s.start), en = new Date(s.end || s.start);
  const row = [st, fmt_(st, 'EEE'), st, en, round_(s.hrs, 2), s.car || '', s.trips || 0, s.offers || 0, s.accepted || 0, s.declined || 0, s.types || '',
    round_(s.earn, 2), round_(s.cost, 2), round_(s.net, 2), round_(s.netHr, 2), round_(s.goal, 0), s.net >= s.goal ? '🎯 Yes' : Math.round(100 * (s.net || 0) / Math.max(1, s.goal || 1)) + '%',
    round_(s.km, 1), s.from || '', s.to || '', s.areas || '', s.map ? `=HYPERLINK("${String(s.map).replace(/"/g, '%22')}","🗺 Open route")` : '', s.list || '', id];
  const n = sh.getLastRow() - 1;
  const ids = n > 0 ? sh.getRange(2, SH_HEAD.length, n, 1).getValues().map(r => String(r[0])) : [];
  const k = ids.indexOf(id), r = k >= 0 ? k + 2 : sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setValues([row]);
  sh.getRange(r, 1).setNumberFormat('d mmm yyyy');
  sh.getRange(r, 3, 1, 2).setNumberFormat('h:mm am/pm');
  sh.getRange(r, 12, 1, 5).setNumberFormat('$#,##0.00');
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, SH_HEAD.length).sort({ column: 3, ascending: true });
  updateSummary_(name);
  return name;
}
function round_(v, d) { const f = Math.pow(10, d); return Math.round((+v || 0) * f) / f; }

function summarySheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Summary');
  if (!sh) {
    sh = ss.insertSheet('Summary', ss.getSheetByName('Flights') ? 1 : 0);
    sh.getRange(1, 1, 1, 9).setValues([['Month', 'Shifts', 'Hours', 'Trips', 'Fares $', 'Car costs $', 'Take-home $', '$ per hour', 'km']])
      .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}
function updateSummary_(name) {
  const m = SpreadsheetApp.getActive().getSheetByName(name), n = m.getLastRow() - 1;
  const v = n > 0 ? m.getRange(2, 1, n, SH_HEAD.length).getValues() : [];
  const sum = i => v.reduce((a, r) => a + (+r[i] || 0), 0);
  const hrs = sum(4), net = sum(13);
  const line = [name, v.length, round_(hrs, 1), sum(6), round_(sum(11), 2), round_(sum(12), 2), round_(net, 2), hrs ? round_(net / hrs, 2) : 0, round_(sum(17), 1)];
  const sh = summarySheet_(), k = sh.getLastRow() - 1;
  const names = k > 0 ? sh.getRange(2, 1, k, 1).getValues().map(r => String(r[0])) : [];
  const i = names.indexOf(name), r = i >= 0 ? i + 2 : sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, line.length).setValues([line]);
  sh.getRange(r, 5, 1, 4).setNumberFormat('$#,##0.00');
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, 9).sort({ column: 1, ascending: false });
}
