/**
 * الاجتماعات — Web App backend for meetings.html
 * ─────────────────────────────────────────────────────────────────────
 * SETUP (once)
 *  1. Upload الاجتماعات.xlsx to Drive, open it, then File ▸ Save as Google Sheets.
 *     (Apps Script cannot open a plain .xlsx — it must be a Google Sheet.)
 *  2. In the Google Sheet: Extensions ▸ Apps Script. Delete the sample code and paste this whole file.
 *     Use a NEW project — do not paste it into the db.gs / data.gs project (both define doGet/doPost).
 *  3. Put the Google Sheet's ID in MEETINGS_SPREADSHEET_ID below
 *     (the long string in its URL: /spreadsheets/d/<ID>/edit).
 *  4. Run setupMeetings() once from the editor and accept the permission prompt.
 *     The Execution log should say the structure is OK.
 *  5. Deploy ▸ New deployment ▸ Web app ▸ Execute as: Me ▸ Who has access: Anyone ▸ Deploy.
 *     Copy the /exec URL into MEETINGS_URL at the top of meetings.html.
 *  6. After any later change to this file: Deploy ▸ Manage deployments ▸ pencil ▸ Version: New version.
 *
 * ACCESS CONTROL
 *  "Anyone" means anyone who has the /exec URL can call it. As in db.gs, an optional shared token
 *  is supported: Project Settings ▸ Script properties ▸ add MEETINGS_TOKEN = <random string>, and put
 *  the same value in MEETINGS_TOKEN at the top of meetings.html. Leave the property unset and the
 *  endpoint stays open. Setting it is recommended. (The page's sign-in code is a client-side gate only.)
 *
 * SHEET STRUCTURE (header text in row 1 is what the script looks for — column ORDER does not matter,
 * but do not rename the sheets or the headers)
 *   الاجتماعات        النوع | رقم الاجتماع | عنوان الاجتماع | تاريخ الانعقاد | تاريخ الاعتماد  (+ computed counts)
 *   شئون العاملين     رقم الاجتماع | الكود | الاسم | الموضوع | مرفوض | تم البت
 *   التظلمات          (same columns as شئون العاملين)
 *  مرفوض = "نعم"  → red text      ·      تم البت = "نعم"  → green background
 */

var MEETINGS_SPREADSHEET_ID = '1CD2PHhpEUltA3fg600OmM6N09rj0YlTY-NpLi803h84';

var REGISTRY_SHEET = 'الاجتماعات';
var KIND_SHEETS = ['شئون العاملين', 'التظلمات'];
var YES = 'نعم';

var HEAD = {
  kind: 'النوع', number: 'رقم الاجتماع', title: 'عنوان الاجتماع', held: 'تاريخ الانعقاد', approved: 'تاريخ الاعتماد',
  code: 'الكود', name: 'الاسم', subject: 'الموضوع', denied: 'مرفوض', decided: 'تم البت'
};
var REGISTRY_KEYS = ['kind', 'number', 'title', 'held', 'approved'];
var REQUEST_KEYS = ['number', 'code', 'name', 'subject', 'denied', 'decided'];
// "رقم الاجتماع" is the first column of every request sheet; its header is HEAD.number.

var CACHE_KEY = 'meetings-v1';
var CACHE_TTL_SECONDS = 30;

/* ───────────────────────── access control ───────────────────────── */

function isAuthorized(token) {
  var required = PropertiesService.getScriptProperties().getProperty('MEETINGS_TOKEN');
  if (!required) return true;
  return token === required;
}

/* ───────────────────────── web app entry points ───────────────────────── */

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (!isAuthorized(p.token)) return respond(false, 'غير مصرح بهذا الطلب.');
    if (p.action === 'meetings') return respond(true, 'تم تحميل البيانات.', getMeetingsCached());
    return ContentService.createTextOutput('الاجتماعات sync endpoint is running.');
  } catch (err) {
    return respond(false, 'خطأ في الخادم: ' + err.message);
  }
}

function doPost(e) {
  var lock = null;
  var touched = false;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!isAuthorized(body.token)) return respond(false, 'غير مصرح بهذا الطلب.');
    var handlers = {
      setStatus: handleSetStatus,
      addRequest: handleAddRequest,
      addMeeting: handleAddMeeting,
      updateMeeting: handleUpdateMeeting
    };
    var handler = handlers[body.action];
    if (!handler) return respond(false, 'إجراء غير معروف: ' + body.action);

    // Handlers read-modify-write; serialize them so two people never overwrite each other.
    lock = LockService.getScriptLock();
    if (!lock.tryLock(25000)) return respond(false, 'الشيت مشغول حالياً، حاول مرة أخرى.');
    touched = true;

    var ss = SpreadsheetApp.openById(MEETINGS_SPREADSHEET_ID);
    var message = handler(ss, body);
    clearCache();
    return respond(true, message, readMeetings(ss));
  } catch (err) {
    return respond(false, err.message);
  } finally {
    if (touched) clearCache();
    if (lock) lock.releaseLock();
  }
}

function respond(ok, message, data) {
  return ContentService.createTextOutput(JSON.stringify({ ok: ok, message: message, data: data || null }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────── helpers ───────────────────────── */

function norm(s) {
  return String(s == null ? '' : s)
    .replace(/[\u064B-\u065F\u0640]/g, '')
    .replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه')
    .replace(/\s+/g, '').toLowerCase();
}

function digits(s) {
  return String(s == null ? '' : s)
    .replace(/[\u0660-\u0669]/g, function (d) { return String(d.charCodeAt(0) - 0x0660); })
    .replace(/[\u06F0-\u06F9]/g, function (d) { return String(d.charCodeAt(0) - 0x06F0); });
}

function clean(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }

function isYes(v) {
  if (v === true) return true;
  var n = norm(v);
  return n === norm(YES) || n === 'yes' || n === 'true' || n === '1';
}

function pad2(n) { return (n < 10 ? '0' : '') + n; }

// Cell value → 'yyyy-MM-dd' (or '' when empty / not a date).
function toIso(v, tz) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = digits(v).trim();
  var m = /^(\d{1,2})\s*[\/\-.]\s*(\d{1,2})\s*[\/\-.]\s*(\d{4})$/.exec(s);
  if (m) return m[3] + '-' + pad2(+m[2]) + '-' + pad2(+m[1]);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  return '';
}

// 'yyyy-MM-dd' → Date for a sheet cell ('' clears the cell).
function fromIso(s) {
  s = clean(s);
  if (!s) return '';
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (!m) throw new Error('تاريخ غير صحيح: ' + s);
  var d = new Date(+m[1], +m[2] - 1, +m[3]);
  if (d.getFullYear() !== +m[1] || d.getMonth() !== +m[2] - 1 || d.getDate() !== +m[3]) throw new Error('تاريخ غير صحيح: ' + s);
  return d;
}

function colIndex(headers, label) {
  var want = norm(label);
  for (var i = 0; i < headers.length; i++) if (norm(headers[i]) === want) return i;
  return -1;
}

// Loads a sheet once: { sheet, headers, rows, col: {key: index}, lastRow }
function loadTable(ss, name, keys) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('لم يتم العثور على الشيت: ' + name);
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0];
  var col = {};
  keys.forEach(function (k) {
    col[k] = colIndex(headers, HEAD[k]);
    if (col[k] === -1) throw new Error('العمود "' + HEAD[k] + '" غير موجود في الشيت "' + name + '".');
  });
  return { sheet: sheet, headers: headers, rows: values.slice(1), col: col, lastCol: lastCol };
}

// First empty row (1-based) judged by the first column — fills gaps, otherwise appends.
function firstEmptyRow(table, keyCol) {
  for (var i = 0; i < table.rows.length; i++) {
    if (clean(table.rows[i][keyCol]) === '') return i + 2;
  }
  return table.rows.length + 2;
}

function kindName(kind) {
  var k = clean(kind);
  for (var i = 0; i < KIND_SHEETS.length; i++) if (norm(KIND_SHEETS[i]) === norm(k)) return KIND_SHEETS[i];
  throw new Error('نوع غير معروف: ' + kind);
}

function meetingNumber(v) {
  var n = parseInt(digits(v).trim(), 10);
  if (!isFinite(n) || n < 1) throw new Error('رقم الاجتماع غير صحيح: ' + v);
  return n;
}

/* ───────────────────────── reading ───────────────────────── */

function getMeetingsCached() {
  var cache = null;
  try {
    cache = CacheService.getScriptCache();
    var hit = cache.get(CACHE_KEY);
    if (hit) return JSON.parse(hit);
  } catch (err) { cache = null; }

  var data = readMeetings(SpreadsheetApp.openById(MEETINGS_SPREADSHEET_ID));
  if (cache) {
    try {
      var s = JSON.stringify(data);
      if (Utilities.newBlob(s).getBytes().length <= 95000) cache.put(CACHE_KEY, s, CACHE_TTL_SECONDS);
    } catch (err) { /* a cache miss is only a performance matter */ }
  }
  return data;
}

function clearCache() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch (err) { /* ignore */ }
}

// → { kinds: [ { name, meetings: [ { number, title, held, approved, requests: [...] } ] } ] }
function readMeetings(ss) {
  var tz = ss.getSpreadsheetTimeZone();
  var reg = loadTable(ss, REGISTRY_SHEET, REGISTRY_KEYS);
  var kinds = KIND_SHEETS.map(function (name) {
    var meetings = {};
    reg.rows.forEach(function (r) {
      var number = parseInt(digits(r[reg.col.number]), 10);
      if (!isFinite(number) || norm(r[reg.col.kind]) !== norm(name)) return;
      meetings[number] = {
        number: number,
        title: clean(r[reg.col.title]) || ('الاجتماع ' + number),
        held: toIso(r[reg.col.held], tz),
        approved: toIso(r[reg.col.approved], tz),
        requests: []
      };
    });

    var t = loadTable(ss, name, REQUEST_KEYS);
    t.rows.forEach(function (r, i) {
      var number = parseInt(digits(r[t.col.number]), 10);
      var code = clean(r[t.col.code]), nm = clean(r[t.col.name]), subject = clean(r[t.col.subject]);
      if (!isFinite(number) || (!code && !nm && !subject)) return;
      // a request whose meeting is missing from the registry is still shown, never hidden
      if (!meetings[number]) meetings[number] = { number: number, title: 'الاجتماع ' + number, held: '', approved: '', requests: [] };
      meetings[number].requests.push({
        row: i + 2, code: code, name: nm, subject: subject,
        denied: isYes(r[t.col.denied]), decided: isYes(r[t.col.decided])
      });
    });

    var list = Object.keys(meetings).map(function (k) { return meetings[k]; })
      .sort(function (a, b) { return b.number - a.number; });   // newest first
    return { name: name, meetings: list };
  });
  return { kinds: kinds, generatedAt: new Date().toISOString() };
}

/* ───────────────────────── writing ───────────────────────── */

// Finds the sheet row of one request. The row number sent by the page is trusted only if the
// meeting number and code still match (someone may have sorted/inserted rows since the page loaded).
function findRequestRow(t, meeting, code, hintRow) {
  var wantCode = clean(digits(code));
  function matches(i) {
    return parseInt(digits(t.rows[i][t.col.number]), 10) === meeting && clean(digits(t.rows[i][t.col.code])) === wantCode;
  }
  var hint = parseInt(hintRow, 10);
  if (hint >= 2 && hint - 2 < t.rows.length && matches(hint - 2)) return hint;
  var found = [];
  for (var i = 0; i < t.rows.length; i++) if (matches(i)) found.push(i + 2);
  if (found.length === 1) return found[0];
  if (!found.length) throw new Error('لم يتم العثور على الطلب (كود ' + wantCode + ') في الاجتماع ' + meeting + '. حدّث الصفحة وحاول مرة أخرى.');
  throw new Error('يوجد أكثر من طلب بنفس الكود في هذا الاجتماع. عدّل الحالة من الشيت مباشرة.');
}

function handleSetStatus(ss, b) {
  var kind = kindName(b.kind), meeting = meetingNumber(b.meeting);
  if (b.denied === undefined && b.decided === undefined) throw new Error('لا توجد حالة لتعديلها.');
  var t = loadTable(ss, kind, REQUEST_KEYS);
  var row = findRequestRow(t, meeting, b.code, b.row);
  if (b.denied !== undefined) t.sheet.getRange(row, t.col.denied + 1).setValue(b.denied ? YES : '');
  if (b.decided !== undefined) t.sheet.getRange(row, t.col.decided + 1).setValue(b.decided ? YES : '');
  return 'تم حفظ الحالة.';
}

function handleAddRequest(ss, b) {
  var kind = kindName(b.kind), meeting = meetingNumber(b.meeting);
  var code = clean(digits(b.code)), name = clean(b.name), subject = clean(b.subject);
  if (!code || !name || !subject) throw new Error('الكود والاسم والموضوع مطلوبة.');

  var reg = loadTable(ss, REGISTRY_SHEET, REGISTRY_KEYS);
  var exists = reg.rows.some(function (r) {
    return norm(r[reg.col.kind]) === norm(kind) && parseInt(digits(r[reg.col.number]), 10) === meeting;
  });
  if (!exists) throw new Error('الاجتماع ' + meeting + ' غير مسجل في شيت "' + REGISTRY_SHEET + '" لنوع ' + kind + '.');

  var t = loadTable(ss, kind, REQUEST_KEYS);
  var dup = t.rows.some(function (r) {
    return parseInt(digits(r[t.col.number]), 10) === meeting && clean(digits(r[t.col.code])) === code
      && norm(r[t.col.subject]) === norm(subject);
  });
  if (dup) throw new Error('هذا الطلب مسجل بالفعل في نفس الاجتماع.');

  var row = new Array(t.lastCol);
  for (var i = 0; i < row.length; i++) row[i] = '';
  row[t.col.number] = meeting;
  row[t.col.code] = /^\d+$/.test(code) ? parseInt(code, 10) : code;
  row[t.col.name] = name;
  row[t.col.subject] = subject;
  row[t.col.denied] = b.denied ? YES : '';
  row[t.col.decided] = b.decided ? YES : '';
  var target = firstEmptyRow(t, t.col.number);
  t.sheet.getRange(target, 1, 1, t.lastCol).setValues([row]);
  return 'تمت إضافة الطلب.';
}

function handleAddMeeting(ss, b) {
  var kind = kindName(b.kind), number = meetingNumber(b.number);
  var reg = loadTable(ss, REGISTRY_SHEET, REGISTRY_KEYS);
  var exists = reg.rows.some(function (r) {
    return norm(r[reg.col.kind]) === norm(kind) && parseInt(digits(r[reg.col.number]), 10) === number;
  });
  if (exists) throw new Error('الاجتماع ' + number + ' مسجل بالفعل لنوع ' + kind + '.');
  var title = clean(b.title) || ('الاجتماع ' + number);
  var target = firstEmptyRow(reg, reg.col.kind);
  // only the five input columns are written — the count columns are formulas already filled down
  var sh = reg.sheet;
  sh.getRange(target, reg.col.kind + 1).setValue(kind);
  sh.getRange(target, reg.col.number + 1).setValue(number);
  sh.getRange(target, reg.col.title + 1).setValue(title);
  var held = fromIso(b.held), approved = fromIso(b.approved);
  sh.getRange(target, reg.col.held + 1).setValue(held);
  sh.getRange(target, reg.col.approved + 1).setValue(approved);
  if (held) sh.getRange(target, reg.col.held + 1).setNumberFormat('dd/mm/yyyy');
  if (approved) sh.getRange(target, reg.col.approved + 1).setNumberFormat('dd/mm/yyyy');
  return 'تمت إضافة الاجتماع.';
}

function handleUpdateMeeting(ss, b) {
  var kind = kindName(b.kind), number = meetingNumber(b.number);
  var reg = loadTable(ss, REGISTRY_SHEET, REGISTRY_KEYS);
  var row = -1;
  for (var i = 0; i < reg.rows.length; i++) {
    if (norm(reg.rows[i][reg.col.kind]) === norm(kind) && parseInt(digits(reg.rows[i][reg.col.number]), 10) === number) { row = i + 2; break; }
  }
  if (row === -1) throw new Error('الاجتماع ' + number + ' غير موجود لنوع ' + kind + '.');
  var changes = [];
  if (b.title !== undefined) { reg.sheet.getRange(row, reg.col.title + 1).setValue(clean(b.title)); changes.push('العنوان'); }
  if (b.held !== undefined) {
    var c = reg.sheet.getRange(row, reg.col.held + 1); c.setValue(fromIso(b.held)); c.setNumberFormat('dd/mm/yyyy'); changes.push('تاريخ الانعقاد');
  }
  if (b.approved !== undefined) {
    var d = reg.sheet.getRange(row, reg.col.approved + 1); d.setValue(fromIso(b.approved)); d.setNumberFormat('dd/mm/yyyy'); changes.push('تاريخ الاعتماد');
  }
  if (!changes.length) throw new Error('لا توجد بيانات لتعديلها.');
  return 'تم حفظ بيانات الاجتماع.';
}

/* ───────────────────────── one-time check ───────────────────────── */

// Run once from the editor: authorizes the script and verifies every sheet and header it needs.
function setupMeetings() {
  var ss = SpreadsheetApp.openById(MEETINGS_SPREADSHEET_ID);
  loadTable(ss, REGISTRY_SHEET, REGISTRY_KEYS);
  KIND_SHEETS.forEach(function (n) { loadTable(ss, n, REQUEST_KEYS); });
  var d = readMeetings(ss), summary = [];
  d.kinds.forEach(function (k) {
    var n = 0; k.meetings.forEach(function (m) { n += m.requests.length; });
    summary.push(k.name + ': ' + k.meetings.length + ' اجتماع، ' + n + ' طلب');
  });
  Logger.log('OK — ' + summary.join(' | '));
  return summary;
}
