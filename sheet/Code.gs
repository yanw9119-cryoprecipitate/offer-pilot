/**
 * Offer Pilot data sheet
 * ----------------------
 * 1. Live Melbourne Airport flights. This script calls AeroDataBox on a timer and keeps the next 7 days of
 *    flights in the "Flights" tab. Every Offer Pilot (any phone, tablet or computer) reads them from here,
 *    so the pilots use none of your RapidAPI calls. The key is kept in this script's settings, never in the app.
 * 2. Your shift log. When you punch out, the pilot sends the shift here: one tab per month
 *    ("2026-10 October"), one row per shift, and a "Summary" tab with a line for each month.
 * 3. Sync between your devices (hidden "Sync" tab), shown in their own files in the Offer Pilot folder:
 *    "Offer Pilot Shifts", "Offer Pilot Devices" and "Offer Pilot Trends" (the Uber demand charts you've uploaded).
 * 4. Fuel stations and EV chargers across Greater Melbourne ("Offer Pilot Fuel Stations"), from OpenStreetMap once a week.
 *    Pilots load them from here, so the ⛽ map layer works even when OpenStreetMap's servers are busy.
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
const SHIFTS_FILE = '1q2YZqYaeSMbimIktSgH0bD5BihQ6_zRAAUOZnFpfgcA';
// "Offer Pilot Shifts" (Upcoming + History tabs); a SHIFTS_ID script property overrides it
const DEVICES_FILE = '1Khp9tA29J1cdGAKU6FkOmB4FdH16zhEWK4w-GD1Ltso';  // "Offer Pilot Devices" (one tab); a DEVICES_ID script property overrides it
const TRENDS_FILE = '1AdZbMjLRBmDloWFbRpr4GTVXvLgGVb-2m55ft2Pn2_o';   // "Offer Pilot Trends" (a tab per zone); a TRENDS_ID script property overrides it
const FUEL_FILE = '1qKKRZkaB6mMroShlWu8SnWOcEAF2zSgx8VsqDkuMiuY';     // "Offer Pilot Fuel Stations"; a FUEL_ID script property overrides it
const FUEL_BOX = { s: -38.45, w: 144.40, n: -37.40, e: 145.65 };      // Greater Melbourne: Werribee, Melton, Sunbury, Lilydale, Pakenham, Frankston
const FUEL_EVERY_D = 7;                                                 // refreshed from OpenStreetMap once a week (free, no key)
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
    .addItem('Update fuel stations now', 'updateFuel')
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
  try { writeShiftSheets_(syncRead_(syncSheet_())); } catch (e) {}
  try { writeDevices_(); } catch (e) {}   // keeps "Now" and "Upcoming" current every hour
  try { const m = syncRead_(syncSheet_()); if (m.uberTrends && m.uberTrends.t > (+P.getProperty('TRENDS_AT') || 0)) writeTrends_(m); } catch (e) {}
  try { if (Date.now() - (+P.getProperty('FUEL_AT') || 0) > FUEL_EVERY_D * 864e5 && Date.now() - (+P.getProperty('FUEL_TRY') || 0) > 6 * 3600000) refreshFuel_(); } catch (e) { P.setProperty('FUEL_ERR', e.message); }
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
  if (p.fuel) return json_(fuelOut_());
  if (p.d) seen_(p.d, { n: p.n, ver: p.ver, acc: 'Flights' });
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
    if (b.kind === 'sync') return json_(doSync_(b));
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

// ---------------------------------------------------------------- sync between your devices
// The hidden "Sync" tab keeps one row per pilot item (offers, roster, availability, shifts, cars, settings …):
// key · when it last changed · which device · the value, split into 40,000-character pieces (each starts with ~ so
// Sheets never turns it into a number, date or formula). The newest change of each item wins.
function syncSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Sync');
  if (!sh) { sh = ss.insertSheet('Sync'); sh.getRange(1, 1, 1, 4).setValues([['Item', 'Changed', 'Device', 'Value …']]).setFontWeight('bold'); sh.setFrozenRows(1); sh.hideSheet(); }
  return sh;
}
function syncRead_(sh) {
  const n = sh.getLastRow() - 1, w = Math.max(4, sh.getLastColumn()), m = {};
  if (n > 0) sh.getRange(2, 1, n, w).getValues().forEach((r, i) => {
    const k = String(r[0]); if (!k) return;
    m[k] = { row: i + 2, t: +r[1] || 0, v: r.slice(3).filter(x => x !== '').map(x => String(x).slice(1)).join('') };
  });
  return m;
}
function doSync_(b) {
  if (b.devId) seen_(b.devId, { n: b.dev, ver: b.ver, acc: 'Sync', sync: true });
  const sh = syncSheet_(), m = syncRead_(sh), put = b.put || {}, have = b.have || {};
  // Trend charts: combined zone by zone and hour by hour (not newest-wins), so uploads made on different devices all count.
  // trendDel (zones you deleted or reset) goes first so they aren't brought back from the other device.
  const order = Object.keys(put).sort((a, b) => (a === 'trendDel' ? -1 : b === 'trendDel' ? 1 : 0));
  order.forEach(k => {
    const p = put[k]; if (!p || !(+p.t > 0)) return;
    if (k === 'uberTrends' && m[k] && m[k].v && p.v) {
      const sent = String(p.v), del = jparse_((m.trendDel || {}).v, {}), merged = JSON.stringify(mergeTrends_(jparse_(m[k].v, {}), jparse_(sent, {}), del));
      if (merged === m[k].v) return;                             // nothing new for the sheet
      // newer than both copies; when it differs from what this device sent, it goes back to it in this same reply
      p.t = merged === sent ? Math.max(+p.t, m[k].t + 1) : Math.max(Date.now(), m[k].t + 1, +p.t + 1);
      p.v = merged;
    } else if (m[k] && +p.t <= m[k].t) return;                   // the sheet already has something newer
    const v = String(p.v == null ? '' : p.v), parts = [];
    for (let i = 0; i < v.length; i += 40000) parts.push('~' + v.slice(i, i + 40000));
    if (!parts.length) parts.push('~');
    const row = m[k] ? m[k].row : sh.getLastRow() + 1, vals = [k, +p.t, String(b.dev || '')].concat(parts);
    while (m[k] && vals.length < sh.getLastColumn()) vals.push('');   // blank out pieces left from a longer old value
    sh.getRange(row, 1, 1, vals.length).setValues([vals]);
    m[k] = { row, t: +p.t, v };
  });
  if (['roster', 'shiftHistory', 'shift', 'cars'].some(k => put[k])) { try { writeShiftSheets_(m); } catch (e) { P.setProperty('LAST_ERR', 'Shift sheet: ' + e.message); } }
  if (['uberTrends', 'trendChanges', 'trendDel'].some(k => put[k])) { try { writeTrends_(m); } catch (e) { P.setProperty('LAST_ERR', 'Trends sheet: ' + e.message); } }
  const get = {};
  Object.keys(m).forEach(k => { if (m[k].t > (+have[k] || 0)) get[k] = { t: m[k].t, v: m[k].v }; });
  try { writeDevices_(); } catch (e) {}
  return { ok: true, get, at: Date.now(), devices: devList_() };
}

// ---------------------------------------------------------------- "Offer Pilot Shifts" spreadsheet
// Upcoming: the shift you're punched in to, then every rostered shift still ahead.
// History: each punched shift, planned (from your roster, or the pilot's suggested finish) against actual.
// Rebuilt from the synced roster and shift history whenever a device syncs, and every hour.
const UP_HEAD = ['Status', 'Date', 'Day', 'Start', 'Finish', 'Hours', 'Area', 'Planned by', 'Car', 'When'];
const HI_HEAD = ['Date', 'Day', 'Planned start', 'Planned finish', 'Planned hours', 'Punch in', 'Punch out', 'Actual hours', 'Difference (h)',
  'Start: min late (+) / early (−)', 'Finish: min over (+) / early (−)', 'Plan from', 'Trips', 'Take-home $', '$ per hour', 'Car', 'Started in', 'Finished in'];
function jparse_(s, d) { try { const v = JSON.parse(s); return v == null ? d : v; } catch (e) { return d; } }
function dayNo_(t) { const [y, m, d] = fmt_(t, 'yyyy-MM-dd').split('-').map(Number); return Date.UTC(y, m - 1, d) / 864e5; }
function shiftsBook_() {
  const id = P.getProperty('SHIFTS_ID') || SHIFTS_FILE;
  return id ? SpreadsheetApp.openById(id) : null;
}
function tab_(book, name, head) {
  let sh = book.getSheetByName(name);
  if (!sh) { sh = book.insertSheet(name); }
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}
function fill_(sh, head, rows) {
  const n = sh.getMaxRows() - 1;
  if (n > 0) sh.getRange(2, 1, n, Math.max(head.length, sh.getMaxColumns())).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
}
function writeShiftSheets_(m) {
  const book = shiftsBook_(); if (!book) return;
  if (book.getSpreadsheetTimeZone() !== TZ) book.setSpreadsheetTimeZone(TZ);
  const roster = jparse_((m.roster || {}).v, []), hist = jparse_((m.shiftHistory || {}).v, []), cur = jparse_((m.shift || {}).v, null), cars = jparse_((m.cars || {}).v, []);
  const carName = id => { const c = (cars || []).find(x => x.id === id) || (id ? null : (cars || [])[0]); return c ? (c.rego || c.name || c.kind || '') : ''; };
  const SRC = { manual: 'You', pilot: '✨ Pilot', 'pilot-adj': '✨ Pilot (adjusted)' };
  const now = Date.now(), when = t => { const d = dayNo_(t) - dayNo_(now); return d <= 0 ? 'Today' : d === 1 ? 'Tomorrow' : 'In ' + d + ' days'; };

  // Upcoming
  const up = [];
  if (cur && cur.start) {
    const pe = cur.planEnd || cur.start + 4 * 3600000;
    up.push(['🟢 Punched in', new Date(cur.start), fmt_(cur.start, 'EEE'), new Date(cur.start), new Date(pe), round_((pe - cur.start) / 3600000, 2), cur.zone || '',
      cur.planSrc === 'roster' ? 'Roster' : cur.planSrc === 'pilot' ? '✨ Pilot' : 'Own finish', carName(cur.car), 'Now']);
  }
  (Array.isArray(roster) ? roster : []).filter(x => x && x.e > now).sort((a, b) => a.s - b.s).forEach(x => {
    up.push([x.s <= now ? 'Rostered now' : 'Upcoming', new Date(x.s), fmt_(x.s, 'EEE'), new Date(x.s), new Date(x.e), round_((x.e - x.s) / 3600000, 2), x.z || '', SRC[x.src] || 'You', '', when(x.s)]);
  });
  const su = tab_(book, 'Upcoming', UP_HEAD);
  fill_(su, UP_HEAD, up);
  if (up.length) { su.getRange(2, 2, up.length, 1).setNumberFormat('d mmm yyyy'); su.getRange(2, 4, up.length, 2).setNumberFormat('h:mm am/pm'); }
  su.getRange(1, UP_HEAD.length + 2).setValue('Updated ' + fmt_(now, 'EEE d MMM h:mm a'));

  // History: planned against actual
  const rows = (Array.isArray(hist) ? hist : []).filter(h => h && h.start).sort((a, b) => b.start - a.start).map(h => {
    const end = h.end || h.start + (h.hrs || 0) * 3600000;
    let best = null, ov = 0;
    (Array.isArray(roster) ? roster : []).forEach(x => { const o = Math.min(end, x.e) - Math.max(h.start, x.s); if (o > ov) { ov = o; best = x; } });
    let ps = null, pe = null, from = '';
    if (best) { ps = best.s; pe = best.e; from = 'Roster'; }
    else if (h.planE) { ps = h.start; pe = h.planE; from = h.planSrc === 'pilot' ? '✨ Pilot suggestion' : h.planSrc === 'roster' ? 'Roster' : 'Own finish'; }
    const ph = ps ? (pe - ps) / 3600000 : '', ah = (end - h.start) / 3600000;
    return [new Date(h.start), fmt_(h.start, 'EEE'), ps ? new Date(ps) : '', pe ? new Date(pe) : '', ph === '' ? '' : round_(ph, 2),
      new Date(h.start), new Date(end), round_(ah, 2), ph === '' ? '' : round_(ah - ph, 2),
      ps ? Math.round((h.start - ps) / 60000) : '', pe ? Math.round((end - pe) / 60000) : '', from || 'Not planned',
      h.trips || 0, round_(h.net, 2), ah > 0 ? round_((h.net || 0) / ah, 2) : 0, carName(h.car), h.from || '', h.to || ''];
  });
  const sh = tab_(book, 'History', HI_HEAD);
  fill_(sh, HI_HEAD, rows);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('d mmm yyyy');
    sh.getRange(2, 3, rows.length, 2).setNumberFormat('h:mm am/pm');
    sh.getRange(2, 6, rows.length, 2).setNumberFormat('h:mm am/pm');
    sh.getRange(2, 14, rows.length, 2).setNumberFormat('$#,##0.00');
  }
  // the placeholder first tab from when the file was made
  book.getSheets().forEach(t => { if (!['Upcoming', 'History'].includes(t.getName()) && book.getSheets().length > 2 && t.getLastRow() <= 3 && /^(Sheet1|Untitled|Offer Pilot Shifts)$/.test(t.getName())) book.deleteSheet(t); });
  if (book.getSheets()[0].getName() !== 'Upcoming') { book.setActiveSheet(su); book.moveActiveSheet(1); }
}

// ---------------------------------------------------------------- devices
// Every pilot that syncs (or reads flights) says who it is: a random id kept on that device, its name, the pilot version.
// Kept in the script's settings and shown in the "Devices" tab. "Active" = seen in the last 15 minutes.
function devs_() { try { return JSON.parse(P.getProperty('DEVICES')) || {}; } catch (e) { return {}; } }
function seen_(id, x) {
  id = String(id).replace(/[^\w-]/g, '').slice(0, 40); if (!id) return;
  const d = devs_(), now = Date.now(), o = d[id] || { first: now, syncs: 0 };
  o.n = String(x.n || o.n || 'Device').slice(0, 40); o.ver = String(x.ver || o.ver || '').slice(0, 12); o.last = now;
  if (x.acc === 'Sync' || o.acc !== 'Sync') o.acc = x.acc;
  if (x.sync) { o.lastSync = now; o.syncs = (o.syncs || 0) + 1; }
  d[id] = o;
  const ids = Object.keys(d).sort((a, b) => d[b].last - d[a].last).slice(0, 40), keep = {};   // remember the 40 most recent
  ids.forEach(k => { keep[k] = d[k]; });
  P.setProperty('DEVICES', JSON.stringify(keep));
}
function devList_() {
  const d = devs_(), now = Date.now();
  return Object.keys(d).map(id => Object.assign({ id, active: now - d[id].last < 15 * 60000 }, d[id])).sort((a, b) => b.last - a.last);
}
// Written to its own spreadsheet, "Offer Pilot Devices" (the old Devices tab in this sheet is removed)
function writeDevices_() {
  const old = SpreadsheetApp.getActive().getSheetByName('Devices'); if (old) SpreadsheetApp.getActive().deleteSheet(old);
  const ss = SpreadsheetApp.openById(P.getProperty('DEVICES_ID') || DEVICES_FILE), head = ['Device', 'Access', 'Pilot version', 'Status', 'Last seen', 'Last sync', 'Syncs', 'First seen', 'Device ID'];
  if (ss.getSpreadsheetTimeZone() !== TZ) ss.setSpreadsheetTimeZone(TZ);
  let sh = ss.getSheetByName('Devices');
  if (!sh) { sh = ss.insertSheet('Devices', 0); }
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff'); sh.setFrozenRows(1);
  ss.getSheets().forEach(t => { if (t.getName() !== 'Devices' && ss.getSheets().length > 1 && t.getLastRow() <= 3 && /^(Sheet1|Untitled|Offer Pilot Devices)$/.test(t.getName())) ss.deleteSheet(t); });
  const now = Date.now(), rows = devList_().map(x => [x.n, x.acc === 'Sync' ? 'Sync (private key)' : 'Flights only', x.ver || '',
    x.active ? '🟢 Active now' : now - x.last < 86400000 ? 'Today' : now - x.last < 7 * 86400000 ? 'This week' : 'Not lately',
    new Date(x.last), x.lastSync ? new Date(x.lastSync) : '', x.syncs || 0, new Date(x.first), x.id]);
  const n = sh.getMaxRows() - 1; if (n > 0) sh.getRange(2, 1, n, head.length).clearContent();
  if (rows.length) { sh.getRange(2, 1, rows.length, head.length).setValues(rows); sh.getRange(2, 5, rows.length, 2).setNumberFormat('ddd d mmm h:mm am/pm'); sh.getRange(2, 8, rows.length, 1).setNumberFormat('d mmm yyyy'); }
  sh.getRange(1, head.length + 2).setValue(`${rows.filter(r => String(r[3]).startsWith('🟢')).length} active now · ${rows.filter(r => r[1].startsWith('Sync')).length} linked for sync · ${rows.length} in total`);
}

// ---------------------------------------------------------------- "Offer Pilot Trends" spreadsheet
// The Uber demand charts you've uploaded (any device), combined. Each zone keeps, for every day and hour, the demand
// level (0 none – 1 very high, averaged over screenshots) and how many screenshots that average is built from.
// Combining two devices' copies: for each hour, the one built from more screenshots; equal → the more recently updated zone.
function mergeTrends_(a, b, del) {
  const out = {}, zones = {};
  Object.keys(a || {}).concat(Object.keys(b || {})).forEach(z => { zones[z] = 1; });
  Object.keys(zones).forEach(z => {
    const x = (a || {})[z], y = (b || {})[z];
    const live = o => o && o.h && !(del && del[z] && del[z] >= (o.updated || 0));
    const p = live(x) ? x : null, q = live(y) ? y : null;
    if (!p && !q) return;
    if (!p || !q) { out[z] = p || q; return; }
    const newer = (q.updated || 0) >= (p.updated || 0) ? q : p, older = newer === q ? p : q, h = {};
    Object.keys(older.h).concat(Object.keys(newer.h)).forEach(k => {
      const n = newer.h[k], o = older.h[k];
      h[k] = !o ? n : !n ? o : (o.n > n.n ? o : n);
    });
    out[z] = { h, shots: Math.max(p.shots || 0, q.shots || 0), updated: Math.max(p.updated || 0, q.updated || 0) };
  });
  return out;
}
const TR_DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
const TR_NAMES = ['None', 'Low', 'Medium', 'High', 'Very high'];
function trHour_(h) { return h === 0 ? '12am' : h < 12 ? h + 'am' : h === 12 ? '12pm' : (h - 12) + 'pm'; }
function trColour_(v) {                       // white → dark blue, like the pilot's demand bars
  if (v === '' || v == null) return '#f8fafc';
  const t = Math.max(0, Math.min(1, v / 4)), c = (lo, hi) => Math.round(lo + (hi - lo) * t);
  const hex = n => ('0' + n.toString(16)).slice(-2);
  return '#' + hex(c(255, 30)) + hex(c(255, 64)) + hex(c(255, 175));
}
// The cell for business day d (Uber's chart runs 4am to 3am) at hour h: 12am–3am is stored under the next calendar day
function trCell_(z, d, h) { const e = z.h[(h < 4 ? (d + 1) % 7 : d) + '-' + h]; return e ? e : null; }
function writeTrends_(m) {
  const id = P.getProperty('TRENDS_ID') || TRENDS_FILE; if (!id) return;
  const book = SpreadsheetApp.openById(id);
  if (book.getSpreadsheetTimeZone() !== TZ) book.setSpreadsheetTimeZone(TZ);
  const all = jparse_((m.uberTrends || {}).v, {}), ch = jparse_((m.trendChanges || {}).v, null);
  const zones = Object.keys(all).filter(z => all[z] && all[z].h).sort((a, b) => a === 'Melbourne City' ? -1 : b === 'Melbourne City' ? 1 : a.localeCompare(b));
  const HOURS = []; for (let i = 0; i < 24; i++) HOURS.push((4 + i) % 24);
  const now = Date.now();

  // Zones: one line each
  const ZH = ['Zone', 'Screenshots', 'Updated', 'Busiest times (usual week)', 'Quietest daytime'];
  const zrows = zones.map(z => {
    const cells = [];
    TR_DAYS.forEach((dn, d) => HOURS.forEach(h => { const e = trCell_(all[z], d, h); if (e) cells.push({ d, h, v: e.v }); }));
    const top = cells.slice().sort((x, y) => y.v - x.v).slice(0, 4).map(c => TR_DAYS[c.d] + ' ' + trHour_(c.h)).join(', ');
    const low = cells.filter(c => c.h >= 8 && c.h <= 20).sort((x, y) => x.v - y.v).slice(0, 3).map(c => TR_DAYS[c.d] + ' ' + trHour_(c.h)).join(', ');
    return [z, all[z].shots || 0, all[z].updated ? new Date(all[z].updated) : '', top, low];
  });
  const sz = tab_(book, 'Zones', ZH);
  fill_(sz, ZH, zrows);
  if (zrows.length) sz.getRange(2, 3, zrows.length, 1).setNumberFormat('ddd d mmm yyyy h:mm am/pm');
  sz.getRange(1, ZH.length + 2).setValue('Updated ' + fmt_(now, 'EEE d MMM h:mm a') + ' · demand 0 = none … 4 = very high · each zone has its own tab');
  sz.setColumnWidth(1, 200); sz.setColumnWidth(4, 300); sz.setColumnWidth(5, 240);

  // One tab per zone: hours down, days across, coloured like a heat map
  zones.forEach(z => {
    const name = z.slice(0, 90);
    let sh = book.getSheetByName(name); if (!sh) sh = book.insertSheet(name);
    const vals = [['Hour'].concat(TR_DAYS)], bgs = [['#0f172a', '#0f172a', '#0f172a', '#0f172a', '#0f172a', '#0f172a', '#0f172a', '#0f172a']], ns = [];
    HOURS.forEach(h => {
      const row = [trHour_(h)], bg = ['#ffffff'], nr = [];
      TR_DAYS.forEach((dn, d) => { const e = trCell_(all[z], d, h), v = e ? Math.round(e.v * 40) / 10 : ''; row.push(v); bg.push(trColour_(v)); nr.push(e ? e.n : ''); });
      vals.push(row); bgs.push(bg); ns.push(nr);
    });
    sh.clear();
    sh.getRange(1, 1, vals.length, 8).setValues(vals).setBackgrounds(bgs).setHorizontalAlignment('center');
    sh.getRange(1, 1, 1, 8).setFontWeight('bold').setFontColor('#ffffff');
    sh.getRange(2, 1, 24, 1).setFontWeight('bold').setHorizontalAlignment('left');
    sh.getRange(2, 2, 24, 7).setNumberFormat('0.0').setFontColor('#0f172a');
    // dark cells get white text
    const fc = vals.slice(1).map(r => r.slice(1).map(v => v !== '' && v >= 2.4 ? '#ffffff' : '#0f172a'));
    sh.getRange(2, 2, 24, 7).setFontColors(fc);
    sh.setFrozenRows(1); sh.setFrozenColumns(1);
    sh.getRange(1, 10, 5, 1).setValues([[z], ['Demand: 0 none · 1 low · 2 medium · 3 high · 4 very high'], ['Uber\'s day runs 4am to 3am: 12am–3am belong to the night before'],
      [`${all[z].shots || 0} screenshots · updated ${all[z].updated ? fmt_(all[z].updated, 'EEE d MMM h:mm a') : '—'}`], ['Screenshots behind each hour (max 8) are in "All data"']]);
    sh.getRange(1, 10).setFontWeight('bold');
  });

  // Every hour as a list (easy to filter or chart)
  const AH = ['Zone', 'Day', 'Hour', 'Demand (0–4)', 'Level', 'Screenshots'];
  const arows = [];
  zones.forEach(z => TR_DAYS.forEach((dn, d) => HOURS.forEach(h => {
    const e = trCell_(all[z], d, h); if (!e) return;
    const v = Math.round(e.v * 40) / 10;
    arows.push([z, dn, trHour_(h), v, TR_NAMES[Math.round(v)] || '', e.n]);
  })));
  const sa = tab_(book, 'All data', AH);
  fill_(sa, AH, arows);

  // What changed in your latest uploads
  const CH = ['Zone', 'Day', 'Part of day', 'Change', 'Levels'];
  const crows = ch && Array.isArray(ch.items) ? ch.items.map(c => [c.zone, TR_DAYS[c.day] || '', c.period, c.diff > 0 ? '⬆️ Busier' : '⬇️ Quieter', Math.round(c.diff * 40) / 10]) : [];
  const sc = tab_(book, 'Recent changes', CH);
  fill_(sc, CH, crows);
  sc.getRange(1, CH.length + 2).setValue(ch && ch.at ? 'From uploads on ' + fmt_(ch.at, 'EEE d MMM h:mm a') : 'No changes yet');

  // tidy: order, zones you've deleted, the empty first tab
  const keep = ['Zones'].concat(zones.map(z => z.slice(0, 90)), ['All data', 'Recent changes']);
  book.getSheets().forEach(t => { if (!keep.includes(t.getName()) && book.getSheets().length > 1) book.deleteSheet(t); });
  keep.forEach((n, i) => { const t = book.getSheetByName(n); if (t && t.getIndex() !== i + 1) { book.setActiveSheet(t); book.moveActiveSheet(i + 1); } });
  P.setProperty('TRENDS_AT', String((m.uberTrends || {}).t || now));
}
// Run once from the editor to fill the Trends file straight away
function updateTrends() { writeTrends_(syncRead_(syncSheet_())); }

// ---------------------------------------------------------------- "Offer Pilot Fuel Stations" spreadsheet
// Every fuel station and EV charger in Greater Melbourne, from OpenStreetMap (Overpass, free, no key), once a week.
// Hours missing on a station are borrowed from the servo's shop within 60 m (Coles Express, 7-Eleven …).
// Pilots read them with GET …/exec?fuel=1 (no key needed, nothing counted against any allowance).
const FUEL_HEAD = ['Kind', 'Name', 'Brand', 'Address', 'Fuels', 'Hours', 'Hours from', 'Shop', 'Usually 24/7', 'Lat', 'Lng', 'OSM id'];
const OVERPASS = ['https://overpass-api.de/api/interpreter', 'https://overpass.kumi.systems/api/interpreter', 'https://overpass.private.coffee/api/interpreter'];
function overpass_(q) {
  let last = '';
  for (const url of OVERPASS) {
    try {
      const r = UrlFetchApp.fetch(url, { method: 'post', payload: { data: q }, muteHttpExceptions: true });
      if (r.getResponseCode() === 200) return JSON.parse(r.getContentText()).elements || [];
      last = url.split('/')[2] + ' ' + r.getResponseCode();
    } catch (e) { last = url.split('/')[2] + ' ' + e.message; }
  }
  throw new Error('OpenStreetMap busy (' + last + ')');
}
function refreshFuel_() {
  P.setProperty('FUEL_TRY', String(Date.now()));
  const b = FUEL_BOX, bb = [b.s, b.w, b.n, b.e].join(',');
  const els = overpass_(`[out:json][timeout:120];(node["amenity"~"^(fuel|charging_station)$"](${bb});way["amenity"~"^(fuel|charging_station)$"](${bb}););out center tags;`);
  const shopEls = overpass_(`[out:json][timeout:120];(node["shop"="convenience"](${bb});way["shop"="convenience"](${bb}););out center tags;`);
  // shops in ~100 m squares, so each station only checks the squares around it
  const grid = {}, key = (la, ln) => Math.round(la * 1000) + ':' + Math.round(ln * 1000);
  shopEls.forEach(e => { const la = e.lat != null ? e.lat : (e.center || {}).lat, ln = e.lon != null ? e.lon : (e.center || {}).lon; if (la == null) return;
    const k = key(la, ln); (grid[k] = grid[k] || []).push({ la, ln, t: e.tags || {} }); });
  const dist = (a, b, c, d) => { const x = (d - b) * Math.cos((a + c) * Math.PI / 360), y = c - a; return Math.sqrt(x * x + y * y) * 111320; };
  const rows = [];
  els.forEach(e => {
    const t = e.tags || {}, la = e.lat != null ? e.lat : (e.center || {}).lat, ln = e.lon != null ? e.lon : (e.center || {}).lon;
    if (la == null) return;
    const ev = t.amenity === 'charging_station', fuels = [];
    if (!ev) {
      if (t['fuel:octane_91'] === 'yes' || t['fuel:e10'] === 'yes') fuels.push(t['fuel:e10'] === 'yes' ? 'E10' : '91');
      if (t['fuel:octane_95'] === 'yes') fuels.push('95');
      if (t['fuel:octane_98'] === 'yes') fuels.push('98');
      if (t['fuel:diesel'] === 'yes') fuels.push('Diesel');
      if (t['fuel:lpg'] === 'yes') fuels.push('LPG');
    }
    const addr = [t['addr:housenumber'], t['addr:street'], t['addr:suburb'] || t['addr:city']].filter(Boolean).join(' ');
    let hours = t.opening_hours || '', from = '', shop = '';
    if (!ev) {
      let near = null, nd = 60; const la0 = Math.round(la * 1000), ln0 = Math.round(ln * 1000);
      for (let i = -1; i <= 1; i++) for (let j = -1; j <= 1; j++) (grid[(la0 + i) + ':' + (ln0 + j)] || []).forEach(x => { const d = dist(la, ln, x.la, x.ln); if (d < nd) { nd = d; near = x; } });
      if (near) { shop = near.t.brand || near.t.name || ''; if (!hours && near.t.opening_hours) { hours = near.t.opening_hours; from = shop || 'the shop'; } }
    }
    const brand = t.brand || t.operator || '', name = t.name || t.brand || (ev ? 'EV charger' : 'Service station');
    const u24 = !ev && !hours && /7-?eleven|reddy express|coles express/i.test(brand + ' ' + name + ' ' + shop);
    rows.push([ev ? 'EV charger' : 'Fuel', name, brand, addr, fuels.join(' · '), hours, from, shop, u24 ? 'Yes' : '', Math.round(la * 1e6) / 1e6, Math.round(ln * 1e6) / 1e6, e.type + e.id]);
  });
  if (rows.filter(r => r[0] === 'Fuel').length < 100) throw new Error('Only ' + rows.length + ' stations came back; keeping the old list');
  rows.sort((a, b) => (a[0] > b[0] ? 1 : a[0] < b[0] ? -1 : 0) || String(a[2] || a[1]).localeCompare(String(b[2] || b[1])));
  const book = SpreadsheetApp.openById(P.getProperty('FUEL_ID') || FUEL_FILE);
  if (book.getSpreadsheetTimeZone() !== TZ) book.setSpreadsheetTimeZone(TZ);
  const sh = tab_(book, 'Stations', FUEL_HEAD);
  sh.getRange(1, 1, sh.getMaxRows(), 9).setNumberFormat('@');            // keep hours like 06:00-22:00 as text
  fill_(sh, FUEL_HEAD, rows);
  sh.setFrozenColumns(2); sh.setColumnWidth(2, 220); sh.setColumnWidth(4, 220); sh.setColumnWidth(6, 260);
  const now = Date.now();
  sh.getRange(1, FUEL_HEAD.length + 2).setValue(`Updated ${fmt_(now, 'EEE d MMM yyyy h:mm a')} from OpenStreetMap · ${rows.filter(r => r[0] === 'Fuel').length} fuel stations · ${rows.filter(r => r[0] !== 'Fuel').length} EV chargers · refreshed weekly`);
  book.getSheets().forEach(t => { if (t.getName() !== 'Stations' && book.getSheets().length > 1) book.deleteSheet(t); });
  P.setProperty('FUEL_AT', String(now)); P.deleteProperty('FUEL_ERR');
  return rows.length;
}
function fuelOut_() {
  const book = SpreadsheetApp.openById(P.getProperty('FUEL_ID') || FUEL_FILE), sh = book.getSheetByName('Stations');
  const n = sh ? sh.getLastRow() - 1 : 0;
  const rows = n > 0 ? sh.getRange(2, 1, n, FUEL_HEAD.length).getValues().filter(r => r[0] && r[9] !== '') : [];
  return { ok: true, at: +P.getProperty('FUEL_AT') || 0, box: FUEL_BOX, err: P.getProperty('FUEL_ERR') || '',
    fields: ['kind', 'name', 'brand', 'addr', 'fuels', 'hours', 'from', 'shop', 'u24', 'lat', 'lng', 'id'],
    rows: rows.map(r => [r[0] === 'Fuel' ? 'f' : 'e', String(r[1]), String(r[2]), String(r[3]), String(r[4]), String(r[5]), String(r[6]), String(r[7]), r[8] === 'Yes' ? 1 : 0, +r[9], +r[10], String(r[11])]) };
}
function updateFuel() {
  let msg;
  try { msg = '✓ ' + refreshFuel_() + ' fuel stations and EV chargers saved in "Offer Pilot Fuel Stations".'; }
  catch (e) { P.setProperty('FUEL_ERR', e.message); msg = 'Fuel stations not updated: ' + e.message + '. The hourly run tries again in 6 hours.'; }
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}
