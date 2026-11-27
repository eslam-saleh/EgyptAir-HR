/**
 * misc.gs – backend for the متنوعات spreadsheet (a NEW Apps Script project; do not paste it into the
 * db.gs / data.gs / meetings.gs projects, they all define doGet/doPost).
 *
 * Spreadsheet sheets (names are the constants below):
 *   auth       hash | rights          access codes (SHA-256 of the code) and their rights letters
 *   site-data  key | value            the values document-center.html fills into documents
 *   holidays   date | key | name_ar | name_en      official holidays (yyyy-mm-dd, kept as text)
 *   visitors   الكود | وقت الدخول       one row per successful sign-in, written by this script
 *
 * Setup
 *  1. script.google.com ▸ New project ▸ paste this file.
 *  2. Run setupMisc() once from the editor (authorizes the script and checks every sheet).
 *  3. Deploy ▸ New deployment ▸ Web app ▸ Execute as: Me ▸ Who has access: Anyone ▸ copy the URL
 *     into misc-config.js.
 *  4. Optional: Project Settings ▸ Script properties ▸ MISC_TOKEN = <random string>, and the same
 *     value in misc-config.js. Leave it unset to accept any request.
 *  5. Optional, once: run migrateVisitorsFromMeetings() to copy the old الزائرون rows over.
 *
 * Endpoints
 *   GET  ?action=sitedata        → { ok, data: { key: value, ... } }
 *   GET  ?action=holidays        → { ok, holidays: [ { date, key, ar, en }, ... ] }
 *   POST { action:'login', code } → { ok:true, rights } and the visit is logged, or { ok:false }
 *   The auth sheet is never sent to any page.
 */

var MISC_SPREADSHEET_ID = '1Jo1SmGnGyHVbTT9Ywr-_zZ264OGNzsyLThXXD6YVPds';

var AUTH_SHEET = 'auth';
var SITE_SHEET = 'site-data';
var HOLIDAYS_SHEET = 'holidays';
var VISITORS_SHEET = 'visitors';
var VISITORS_HEADER = ['الكود', 'وقت الدخول'];

/* ───────────────────────── entry points ───────────────────────── */

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (!isAuthorized(p.token)) return reply({ ok: false, error: 'غير مصرح بهذا الطلب.' });
    if (p.action === 'sitedata') return reply({ ok: true, data: readSiteData() });
    if (p.action === 'holidays') return reply({ ok: true, holidays: readHolidays() });
    return reply({ ok: false, error: 'إجراء غير معروف.' });
  } catch (err) {
    return reply({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!isAuthorized(body.token)) return reply({ ok: false, error: 'غير مصرح بهذا الطلب.' });
    if (body.action === 'login') return handleLogin(body);
    return reply({ ok: false, error: 'إجراء غير معروف.' });
  } catch (err) {
    return reply({ ok: false, error: String(err && err.message || err) });
  }
}

/* ───────────────────────── helpers ───────────────────────── */

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function isAuthorized(token) {
  var required = PropertiesService.getScriptProperties().getProperty('MISC_TOKEN');
  return !required || String(token || '') === required;
}

function book() { return SpreadsheetApp.openById(MISC_SPREADSHEET_ID); }

// Reads a sheet into [{header: value}]. Headers are matched trimmed + lower-cased; rows with no
// value in the first column are skipped. Dates typed into the sheet come back as yyyy-mm-dd.
function readTable(name) {
  var sh = book().getSheetByName(name);
  if (!sh) throw new Error('الشيت غير موجود: ' + name);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var tz = book().getSpreadsheetTimeZone();
  var heads = values[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = {}, any = false;
    for (var c = 0; c < heads.length; c++) {
      if (!heads[c]) continue;
      var v = values[r][c];
      if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      row[heads[c]] = (v === null || v === undefined) ? '' : String(v).trim();
      if (c === 0 && row[heads[c]]) any = true;
    }
    if (any) out.push(row);
  }
  return out;
}

function digits(s) {
  return String(s || '')
    .replace(/[\u0660-\u0669]/g, function (d) { return String(d.charCodeAt(0) - 0x0660); })
    .replace(/[\u06F0-\u06F9]/g, function (d) { return String(d.charCodeAt(0) - 0x06F0); })
    .replace(/\s+/g, '');
}

function sha256Hex(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/* ───────────────────────── sign-in ───────────────────────── */

function handleLogin(b) {
  var code = digits(b.code);
  if (!/^\d{1,12}$/.test(code)) return reply({ ok: false });
  var hash = sha256Hex(code);
  var rights = null;
  readTable(AUTH_SHEET).forEach(function (r) {
    if (rights === null && String(r.hash || '').toLowerCase() === hash) rights = String(r.rights || '').toLowerCase();
  });
  if (rights === null) return reply({ ok: false });
  try { logVisit(code); } catch (err) { /* a log failure must never block a sign-in */ }
  return reply({ ok: true, rights: rights });
}

// Appends one row per sign-in: the access code and the exact time. Creates the sheet if missing.
function logVisit(code) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('الشيت مشغول حالياً.');
  try {
    var ss = book();
    var sheet = ss.getSheetByName(VISITORS_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(VISITORS_SHEET);
      sheet.getRange(1, 1, 1, 2).setValues([VISITORS_HEADER]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    var row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setNumberFormat('@').setValue(code);
    sheet.getRange(row, 2).setNumberFormat('dd/mm/yyyy hh:mm:ss').setValue(new Date());
  } finally {
    lock.releaseLock();
  }
}

/* ───────────────────────── site data + holidays ───────────────────────── */

function readSiteData() {
  var data = {};
  readTable(SITE_SHEET).forEach(function (r) { if (r.key) data[r.key] = r.value; });
  return data;
}

function readHolidays() {
  var out = [];
  readTable(HOLIDAYS_SHEET).forEach(function (r) {
    if (/^\d{4}-\d{2}-\d{2}$/.test(r.date || '')) {
      out.push({ date: r.date, key: r.key || '', ar: r.name_ar || '', en: r.name_en || '' });
    }
  });
  out.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  return out;
}

/* ───────────────────────── one-time helpers (run from the editor) ───────────────────────── */

// Authorizes the script and checks every sheet and header it needs.
function setupMisc() {
  var ss = book();
  var need = {};
  need[AUTH_SHEET] = ['hash', 'rights'];
  need[SITE_SHEET] = ['key', 'value'];
  need[HOLIDAYS_SHEET] = ['date', 'key', 'name_ar', 'name_en'];
  Object.keys(need).forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (!sh) throw new Error('الشيت غير موجود: ' + n);
    var heads = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]
      .map(function (h) { return String(h).trim().toLowerCase(); });
    need[n].forEach(function (h) { if (heads.indexOf(h) === -1) throw new Error(n + ': العمود غير موجود: ' + h); });
  });
  if (!ss.getSheetByName(VISITORS_SHEET)) logVisit('0');   // creates the sheet; delete that one test row
  var msg = 'OK — auth: ' + readTable(AUTH_SHEET).length + ' | site-data: ' + Object.keys(readSiteData()).length +
    ' | holidays: ' + readHolidays().length;
  Logger.log(msg);
  return msg;
}

