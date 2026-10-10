/**
 * بيانات.xlsx — Web App backend for tagging.html / document-center.html
 * ─────────────────────────────────────────────────────────
 * Deploy this bound to the بيانات.xlsx spreadsheet (Extensions ▸ Apps
 * Script from inside the sheet itself), then Deploy ▸ New deployment ▸
 * Web app, "Execute as: Me", "Who has access: Anyone". Copy the resulting
 * /exec URL into SHEET_SYNC_URL in document-center.html.
 *
 * Column A on every sheet is always the running order number. Every other
 * column is located by matching its header text (row 1) against the Arabic
 * keys sent from the browser (e.g. "كود", "اسم", "نوع الجزاء") — so the
 * exact column order in each sheet doesn't need to match this script.
 *
 * Every handler reads the sheet's used range with ONE getValues() call,
 * does all matching/edits on the in-memory JavaScript array, and writes
 * back with ONE setValues() call (plus, for deletions, a single
 * deleteRow() to drop the now-stale trailing row). Nothing ever loops
 * calling the Sheets API per-row. The one exception is handleUpdate(),
 * which targets a single cell directly — see the note there.
 *
 * doGet(?action=dashboard) is read-only and powers tagging.html's
 * "Deadlines to track" cards (اجازات / جزاءات / ايقاف) — see
 * DASHBOARD_SHEETS, findSheetByName() and getFilledRecords() below.
 * If this script is already deployed, changes here only go live once you
 * redeploy: Deploy ▸ Manage deployments ▸ (pencil icon) ▸ Version: New
 * version ▸ Deploy. The /exec URL itself doesn't change.
 *
 * ── Access control ─────────────────────────────────────────────────
 * "Who has access: Anyone" means anyone with the /exec URL can currently
 * call doPost (append/update/edit/delete اجازات، جزاءات و ايقاف records)
 * AND doGet?action=dashboard (read every employee's name/code/leave/
 * penalty/suspension data) with no login of any kind — the URL itself is
 * the only thing standing between this HR data and the public internet.
 * SYNC_TOKEN below is an optional, low-effort mitigation: set a Script
 * Property named SYNC_TOKEN (Project Settings ▸ Script properties) to any
 * random string, and set the matching SHEET_SYNC_TOKEN in
 * document-center.html / tagging.html to the same value — every request
 * then has to present it. Leave the property unset and everything keeps
 * working exactly as before (nothing breaks by default), but until it's
 * set this endpoint remains fully open to anyone who has (or guesses) the
 * URL. Setting it is strongly recommended.
 */

// ID of بيانات.xlsx (the long string in its URL: /spreadsheets/d/<ID>/edit).
// Used instead of getActiveSpreadsheet(), so the script no longer has to be
// bound to the sheet.
var SPREADSHEET_ID = '1YCAUm-0kSeXS5mdJDwuq3cSYNyC6d52DBTMHlRtQjHI';

// Numbering sheet written by document-center.html when an executive (تنفيذي)
// .docx is generated. Columns: الرقم | التاريخ | الكود | الاسم | الموضوع.
// Only the 'nextNumber' action may touch it (see doPost).
var NUMBERS_SHEET = 'الارقام';

// Sheet tabs read by the tagging.html "Deadlines to track" dashboard.
var DASHBOARD_SHEETS = ['اجازات', 'جزاءات', 'ايقاف'];
// The dashboard is read far more often than these sheets are written. A short
// cache prevents each page visit from making three Spreadsheet service reads,
// which is the main source of intermittent slow/cold web-app responses. It is
// cleared after every valid write request below, so updates made through the
// site appear on the next dashboard load. Direct edits in the spreadsheet can
// take up to this TTL to appear in the dashboard.
var DASHBOARD_CACHE_KEY = 'tagging-dashboard-v1';
var DASHBOARD_CACHE_TTL_SECONDS = 30;

// Returns true if the caller supplied the correct SYNC_TOKEN, OR if no
// SYNC_TOKEN Script Property has been configured yet (opt-in — see the
// "Access control" note above).
function isAuthorized(token) {
  var required = PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN');
  if (!required) return true;
  return token === required;
}

function doPost(e) {
  var lock = null;
  var touched = false;
  try {
    var body = JSON.parse(e.postData.contents);
    if (!isAuthorized(body.token)) return respond(false, 'غير مصرح بهذا الطلب.');
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    // Dashboard tabs accept the usual actions; the numbering sheet accepts ONLY 'nextNumber'.
    var allowedNames = body.action === 'nextNumber' ? [NUMBERS_SHEET] : DASHBOARD_SHEETS;
    var allowed = allowedNames.some(function (n) { return normalizeAr(n) === normalizeAr(body.sheet); });
    if (!allowed) return respond(false, 'شيت غير مسموح به: ' + body.sheet);
    var sheet = findSheetByName(ss, body.sheet);
    if (!sheet) return respond(false, 'لم يتم العثور على الشيت: ' + body.sheet);

    // Handlers read the whole range and write it back, so two simultaneous
    // requests could overwrite each other. Serialize writes.
    lock = LockService.getScriptLock();
    if (!lock.tryLock(25000)) return respond(false, 'الشيت مشغول حالياً، حاول مرة أخرى.');
    touched = true;

    var result;
    switch (body.action) {
      case 'append':        result = handleAppend(sheet, body); break;
      case 'nextNumber':    result = handleNextNumber(sheet, body); break;
      case 'update':         result = handleUpdate(sheet, body); break;
      case 'updatePenalty':  result = handleUpdatePenalty(sheet, body); break;
      case 'deletePenalty':  result = handleDeletePenalty(sheet, body); break;
      default: return respond(false, 'إجراء غير معروف: ' + body.action);
    }
    return result;
  } catch (err) {
    return respond(false, 'خطأ في الخادم: ' + err.message);
  } finally {
    // Cleared in finally so even a write that failed half-way never leaves a
    // stale dashboard behind. Invalidating for a no-op write is harmless.
    if (touched) clearDashboardCache();
    if (lock) lock.releaseLock();
  }
}

// GET handler. Visiting the deployed URL with no params is a simple health
// check. ?action=dashboard returns the filled records from DASHBOARD_SHEETS,
// used by tagging.html's "Deadlines to track" cards.
function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action;
    var token = e && e.parameter && e.parameter.token;
    if (!isAuthorized(token)) return respond(false, 'غير مصرح بهذا الطلب.');

    if (action === 'dashboard') {
      // The spreadsheet is opened lazily inside getDashboardData(), only on a
      // cache miss, so a cache hit never pays for openById().
      return respond(true, 'تم تحميل البيانات.', getDashboardData());
    }

    return ContentService.createTextOutput('بيانات.xlsx sync endpoint is running.');
  } catch (err) {
    return respond(false, 'خطأ في الخادم: ' + err.message);
  }
}

function respond(ok, message, data) {
  return ContentService.createTextOutput(
    JSON.stringify({ ok: ok, message: message, data: data || null })
  ).setMimeType(ContentService.MimeType.JSON);
}

// Uses the shared script cache rather than a browser cache: every visitor can
// benefit, and a reload does not have to read all three sheets again. Cache
// failures never block the dashboard; the function simply falls back to a
// fresh spreadsheet read.
function getDashboardData() {
  var cache;
  try {
    cache = CacheService.getScriptCache();
    var cached = cache.get(DASHBOARD_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch (err) {
    cache = null;
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var out = {};
  DASHBOARD_SHEETS.forEach(function (name) {
    var sheet = findSheetByName(ss, name);
    out[name] = sheet ? getFilledRecords(sheet) : [];
  });
  // Shown by tagging.html in the "اخر رقم تنفيذي" box (null when unavailable).
  out['__lastExecutiveNumber'] = getLastExecutiveNumber(ss);

  if (cache) {
    try {
      // CacheService stores at most 100 KB per entry. Skip caching unusually
      // large future datasets rather than risking a cache error.
      var serialized = JSON.stringify(out);
      if (Utilities.newBlob(serialized).getBytes().length <= 95000) {
        cache.put(DASHBOARD_CACHE_KEY, serialized, DASHBOARD_CACHE_TTL_SECONDS);
      }
    } catch (err) {
      // A cache miss is a performance concern, not a dashboard failure.
    }
  }
  return out;
}

// Last filled value in the الرقم column of the الارقام sheet, as a number (or null).
function getLastExecutiveNumber(ss) {
  try {
    var sheet = findSheetByName(ss, NUMBERS_SHEET);
    if (!sheet) return null;
    var loaded = loadSheet(sheet);
    var numCol = findHeaderCol(loaded.headers, 'الرقم');
    if (numCol === -1) return null;
    for (var i = loaded.data.length - 1; i >= 0; i--) {
      var raw = loaded.data[i][numCol];
      if (raw === '' || raw === null) continue;
      var digits = String(raw)
        .replace(/[٠-٩]/g, function (d) { return String(d.charCodeAt(0) - 0x0660); })
        .match(/\d+/);
      return digits ? parseInt(digits[0], 10) : null;
    }
  } catch (err) {
    // The extra box must never break the dashboard.
  }
  return null;
}

function clearDashboardCache() {
  try {
    CacheService.getScriptCache().remove(DASHBOARD_CACHE_KEY);
  } catch (err) {
    // Never fail a successful sheet write just because cache cleanup failed.
  }
}

// ── Arabic-aware header matching (mirrors the client's normalizeArabicVariants) ──
function normalizeAr(s) {
  s = String(s == null ? '' : s);
  s = s.replace(/[\u064B-\u0652\u0670\u0640]/g, ''); // tashkeel/tatweel
  s = s.replace(/[إأآا]/g, 'ا').replace(/[يى]/g, 'ي').replace(/[هة]/g, 'ه');
  s = s.toLowerCase().replace(/\s+/g, '');
  return s.trim();
}

function findHeaderCol(headers, keyword) {
  var nk = normalizeAr(keyword);
  for (var i = 0; i < headers.length; i++) {
    if (normalizeAr(headers[i]).indexOf(nk) !== -1) return i; // 0-based
  }
  return -1;
}

// ── Date columns: written as literal "yyyy/mm/dd" TEXT, never as a Sheets
// Date value ──────────────────────────────────────────────────────────
// The browser (document-center.html's toSheetDateFormat/sheetSyncToday)
// always sends dates as an unambiguous "yyyy/mm/dd" string. Left alone,
// Sheets' setValues() auto-detects that string as a date and stores it as
// a real Date value, which then DISPLAYS in whatever the spreadsheet's
// own locale is (e.g. dd/mm/yyyy) when the sheet is opened/exported
// directly — not wrong, just not the literal "yyyy/mm/dd" text an editor
// scanning the raw sheet expects, and easy to misread as a day/month swap.
// isDateHeader() flags the columns that hold dates in this app; every
// write to one of them goes through a cell forced to Plain text ("@")
// format first, so Sheets never re-interprets it — the cell always shows
// exactly the text that was sent, in every locale.
function isDateHeader(header) {
  var n = normalizeAr(header);
  return n.indexOf('تاريخ') !== -1 || n.indexOf('بداية') !== -1 || n.indexOf('نهاية') !== -1;
}

// Normalizes any "yyyy/mm/dd" or "yyyy-mm-dd" value into zero-padded
// "yyyy/mm/dd" text. Values that don't match (blank, or some other kind
// of cell content entirely) pass through unchanged.
function normalizeDateCellText(value) {
  var s = String(value == null ? '' : value).trim();
  var m = s.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (!m) return value;
  var pad = function (n) { return n.length < 2 ? '0' + n : n; };
  return m[1] + '/' + pad(m[2]) + '/' + pad(m[3]);
}

// Reads the whole used range in one call and splits it into headers/data.
function loadSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 1) return { headers: [], data: [], lastCol: lastCol, lastRow: lastRow };
  var all = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  return { headers: all[0], data: all.slice(1), lastCol: lastCol, lastRow: lastRow };
}

// Looks up a sheet by exact tab name first, then falls back to an
// Arabic-normalized comparison across all tabs (handles alif/hamza spelling
// variants like "ايقاف" vs "إيقاف" between DASHBOARD_SHEETS and the actual
// tab name).
function findSheetByName(ss, name) {
  var direct = ss.getSheetByName(name);
  if (direct) return direct;
  var target = normalizeAr(name);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (normalizeAr(sheets[i].getName()) === target) return sheets[i];
  }
  return null;
}

// Read-only: every "filled" record (column A / order number populated) from
// a sheet, as an array of { headerText: value, ... } objects keyed by the
// sheet's own header row — so column order doesn't matter to the caller.
// Dates that come back as real Date objects (any cell written before this
// script started forcing Plain text — see isDateHeader() above, or a cell
// typed directly into the sheet and auto-detected by Sheets) are formatted
// as "yyyy/mm/dd" text using the ACTUAL underlying date value, so this is
// correct regardless of the spreadsheet's display locale. Cells already
// stored as "yyyy/mm/dd" text (every date written from now on) pass
// through as-is — both forms land on the client in the same unambiguous
// format.
function getFilledRecords(sheet) {
  var loaded = loadSheet(sheet);
  if (!loaded.data.length) return [];

  var tz = Session.getScriptTimeZone();
  var records = [];
  loaded.data.forEach(function (row) {
    if (row[0] === '' || row[0] === null) return; // blank order = not a real record
    var record = {};
    for (var c = 0; c < loaded.headers.length; c++) {
      var key = String(loaded.headers[c] || '').trim();
      if (!key) continue;
      var val = row[c];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, tz, 'yyyy/MM/dd');
      }
      record[key] = val;
    }
    records.push(record);
  });
  return records;
}

// ── APPEND (one or more rows) ─────────────────────────────────────────
// If body.overwriteByCode is true, a row whose كود already exists in the
// sheet REPLACES the most recent existing row with that code: the old row
// is removed (all its cells cleared, nothing carried over except what the
// document sent), the new record is placed at the END of the sheet, and
// column A is renumbered 1..n. Codes not found are appended as new rows. Duplicate codes inside the same batch
// collapse into one (the later one wins).
function handleAppend(sheet, body) {
  var rows = body.rows || [];
  if (!rows.length) return respond(false, 'لا توجد بيانات لإضافتها.');

  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var lastRow = sheet.getLastRow();
  var headers = lastRow > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    : [];

  var startOrder = 1;
  if (lastRow > 1) {
    var prevOrder = sheet.getRange(lastRow, 1).getValue();
    startOrder = (Number(prevOrder) || 0) + 1;
  }

  var buildRow = function (rowData, order) {
    var out = new Array(lastCol).fill('');
    out[0] = order; // column A = order, never overwritten below
    Object.keys(rowData).forEach(function (key) {
      var col = findHeaderCol(headers, key);
      if (col < 0) return;
      var value = rowData[key];
      out[col] = isDateHeader(headers[col]) ? normalizeDateCellText(value) : value;
    });
    return out;
  };

  // Existing codes → sheet row number (1-based), newest match wins.
  var codeCol = findHeaderCol(headers, 'كود');
  var existing = {};
  if (body.overwriteByCode && codeCol !== -1 && lastRow > 1) {
    var codes = sheet.getRange(2, codeCol + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < codes.length; i++) {
      var k = String(codes[i][0]).trim();
      if (k) existing[k] = i + 2;
    }
  }

  // Rows to place at the END of the sheet, in the order received. This holds
  // both brand-new records and the replacements for overwritten ones.
  var tail = [];
  var tailByCode = {};   // code → index in tail (batch de-dupe, later wins)
  var removedRows = {};  // old sheet row number → true (these rows get dropped)
  var nextOrder = startOrder;

  rows.forEach(function (rowData) {
    var key = body.overwriteByCode && codeCol !== -1
      ? String(rowData['كود'] == null ? '' : rowData['كود']).trim()
      : '';
    if (key && existing[key]) removedRows[existing[key]] = true;
    if (key && tailByCode.hasOwnProperty(key)) {
      tail[tailByCode[key]] = buildRow(rowData, 0);
    } else {
      if (key) tailByCode[key] = tail.length;
      tail.push(buildRow(rowData, 0));
    }
  });

  var overwrittenCount = Object.keys(removedRows).length;
  var newCount = tail.length - overwrittenCount;

  var setDateFormats = function (row, count) {
    for (var c = 0; c < headers.length; c++) {
      if (isDateHeader(headers[c])) sheet.getRange(row, c + 1, count, 1).setNumberFormat('@');
    }
  };

  var startRow;
  if (!overwrittenCount) {
    // Nothing replaced: single bulk write into previously-empty cells.
    tail.forEach(function (r) { r[0] = nextOrder++; });
    startRow = lastRow + 1;
    if (tail.length) {
      setDateFormats(startRow, tail.length);
      sheet.getRange(startRow, 1, tail.length, lastCol).setValues(tail);
    }
  } else {
    // Overwrite: the old row is dropped entirely (so every cell in it —
    // including its old order number — is cleared unless the document sent
    // a value for it), the replacement goes to the end of the sheet, and
    // column A is renumbered 1..n so the ordering stays continuous.
    var all = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var kept = [];
    for (var j = 0; j < all.length; j++) {
      if (!removedRows[j + 2]) kept.push(all[j]);
    }
    var finalRows = kept.concat(tail);

    // Date columns are rewritten as plain "yyyy/mm/dd" text (see
    // isDateHeader()), so any real Date value read back is converted first.
    var tz = Session.getScriptTimeZone();
    finalRows.forEach(function (r, idx) {
      r[0] = idx + 1;
      for (var c = 1; c < headers.length; c++) {
        if (r[c] instanceof Date && isDateHeader(headers[c])) {
          r[c] = Utilities.formatDate(r[c], tz, 'yyyy/MM/dd');
        }
      }
    });

    // finalRows is never shorter than the old data (every removed row has a
    // replacement in tail), so no stale trailing row is left behind.
    setDateFormats(2, finalRows.length);
    sheet.getRange(2, 1, finalRows.length, lastCol).setValues(finalRows);
    startRow = kept.length + 2;
  }

  var parts = [];
  if (newCount) parts.push('تمت إضافة ' + (newCount === 1 ? 'سجل واحد' : newCount + ' سجلات'));
  if (overwrittenCount) parts.push('تم استبدال ' + (overwrittenCount === 1 ? 'سجل واحد' : overwrittenCount + ' سجلات') + ' موجود بنفس الكود ونقله إلى نهاية الشيت');
  return respond(true, parts.join(' و') + ' في شيت ' + sheet.getName() + '.',
    { firstRow: startRow, count: tail.length, added: newCount, overwritten: overwrittenCount });
}

// ── NEXT NUMBER (الارقام sheet) ───────────────────────────────────────
// Takes the last filled الرقم in the sheet, adds one, and appends a new row:
// الرقم = last + 1 (1 if the sheet only has headers), plus the التاريخ /
// الكود / الاسم / الموضوع sent in body.row. Column A (م) also gets last + 1. The date is written as plain
// "yyyy/mm/dd" text like every other date in this script. Runs under the
// script lock taken in doPost, so two simultaneous documents never get the
// same number.
function handleNextNumber(sheet, body) {
  var row = body.row || {};
  var loaded = loadSheet(sheet);
  if (!loaded.headers.length) return respond(false, 'شيت ' + sheet.getName() + ' فارغ (لا توجد عناوين).');

  var numCol = findHeaderCol(loaded.headers, 'الرقم');
  if (numCol === -1) return respond(false, 'تعذر إيجاد عمود الرقم في شيت ' + sheet.getName() + '.');

  // Last row (from the bottom) that actually holds a number.
  var last = 0;
  for (var i = loaded.data.length - 1; i >= 0; i--) {
    var raw = loaded.data[i][numCol];
    if (raw === '' || raw === null) continue;
    var digits = String(raw)
      .replace(/[٠-٩]/g, function (d) { return String(d.charCodeAt(0) - 0x0660); })
      .replace(/[۰-۹]/g, function (d) { return String(d.charCodeAt(0) - 0x06F0); })
      .match(/\d+/);
    if (!digits) return respond(false, 'آخر قيمة في عمود الرقم ليست رقماً: ' + raw);
    last = parseInt(digits[0], 10);
    break;
  }
  var next = last + 1;

  // Column A (م) is the running order, separate from الرقم: last filled م + 1.
  var orderCol = numCol === 0 ? -1 : 0;
  var nextOrder = 1;
  if (orderCol === 0) {
    for (var j = loaded.data.length - 1; j >= 0; j--) {
      var o = loaded.data[j][0];
      if (o === '' || o === null) continue;
      nextOrder = (parseInt(String(o).replace(/[٠-٩]/g, function (d) { return String(d.charCodeAt(0) - 0x0660); }), 10) || 0) + 1;
      break;
    }
  }

  var out = new Array(loaded.lastCol).fill('');
  if (orderCol === 0) out[0] = nextOrder;
  out[numCol] = next;
  Object.keys(row).forEach(function (key) {
    var col = findHeaderCol(loaded.headers, key);
    if (col < 0 || col === numCol || col === orderCol) return;
    out[col] = isDateHeader(loaded.headers[col]) ? normalizeDateCellText(row[key]) : row[key];
  });

  var targetRow = loaded.lastRow + 1;
  for (var c = 0; c < loaded.headers.length; c++) {
    if (isDateHeader(loaded.headers[c])) sheet.getRange(targetRow, c + 1).setNumberFormat('@');
  }
  sheet.getRange(targetRow, 1, 1, loaded.lastCol).setValues([out]);
  return respond(true, 'تمت إضافة الرقم ' + next + ' في شيت ' + sheet.getName() + '.',
    { number: next, row: targetRow });
}

// ── UPDATE (نهاية الإيقاف: fills the newest still-open row for a code) ──
// Writes ONE cell directly (getRange(row, col).setValue(...)) rather than
// rewriting the whole data range: this column can already hold a mix of
// real Date values (older rows) and plain "yyyy/mm/dd" text (rows written
// after the isDateHeader() change above), and a full-range setValues()
// would force every one of those pre-existing cells through the same
// number format as the one we're actually changing — safe for the one
// cell we mean to touch, but a needless (and, for the Date-valued ones,
// potentially display-breaking) side effect on every other row.
function handleUpdate(sheet, body) {
  var loaded = loadSheet(sheet);
  if (!loaded.data.length) return respond(false, 'لا توجد بيانات في شيت ' + sheet.getName() + '.');

  var codeCol = findHeaderCol(loaded.headers, 'كود');
  var targetCol = findHeaderCol(loaded.headers, body.targetHeader);
  if (codeCol === -1 || targetCol === -1) {
    return respond(false, 'تعذر إيجاد أعمدة الكود أو ' + body.targetHeader + ' في شيت ' + sheet.getName() + '.');
  }

  var data = loaded.data;
  for (var r = data.length - 1; r >= 0; r--) { // newest → oldest
    var sameCode = String(data[r][codeCol]).trim() === String(body.code).trim();
    var isOpen = !data[r][targetCol];
    if (sameCode && isOpen) {
      var header = loaded.headers[targetCol];
      var value = isDateHeader(header) ? normalizeDateCellText(body.value) : body.value;
      var cell = sheet.getRange(r + 2, targetCol + 1);
      if (isDateHeader(header)) cell.setNumberFormat('@'); // plain text — see isDateHeader()
      cell.setValue(value); // single targeted write
      return respond(true, 'تم تحديث ' + body.targetHeader + ' بنجاح.', { row: r + 2 });
    }
  }
  return respond(false, 'لم يتم العثور على سجل إيقاف مفتوح لهذا الكود (' + body.code + ') في شيت ' + sheet.getName() + '.');
}

// ── UPDATE PENALTY (خفض جزاء: swaps نوع الجزاء on the matching record) ──
function handleUpdatePenalty(sheet, body) {
  var loaded = loadSheet(sheet);
  if (!loaded.data.length) return respond(false, 'لا توجد بيانات في شيت ' + sheet.getName() + '.');

  var codeCol = findHeaderCol(loaded.headers, 'كود');
  var typeCol = findHeaderCol(loaded.headers, 'نوع الجزاء');
  if (codeCol === -1 || typeCol === -1) {
    return respond(false, 'تعذر إيجاد أعمدة الكود أو نوع الجزاء في شيت ' + sheet.getName() + '.');
  }

  var data = loaded.data;
  var idx = findMostRecentMatch(data, codeCol, body.code, typeCol, body.oldType);
  if (idx === -1) {
    return respond(false, 'لم يتم العثور على جزاء لهذا الكود (' + body.code + ') في شيت ' + sheet.getName() + '.');
  }

  data[idx][typeCol] = body.newType;
  sheet.getRange(2, 1, data.length, loaded.lastCol).setValues(data); // single bulk write
  return respond(true, 'تم تحديث نوع الجزاء بنجاح.', { row: idx + 2 });
}

// ── DELETE PENALTY (محو/إلغاء جزاء: removes the matching record) ──────
function handleDeletePenalty(sheet, body) {
  var loaded = loadSheet(sheet);
  if (!loaded.data.length) return respond(false, 'لا توجد بيانات في شيت ' + sheet.getName() + '.');

  var codeCol = findHeaderCol(loaded.headers, 'كود');
  var typeCol = findHeaderCol(loaded.headers, 'نوع الجزاء'); // may be -1, that's fine
  if (codeCol === -1) return respond(false, 'تعذر إيجاد عمود الكود في شيت ' + sheet.getName() + '.');

  var data = loaded.data;
  var idx = findMostRecentMatch(data, codeCol, body.code, typeCol, body.type);
  if (idx === -1) {
    return respond(false, 'لم يتم العثور على جزاء لهذا الكود (' + body.code + ') في شيت ' + sheet.getName() + '.');
  }

  data.splice(idx, 1); // remove in memory
  for (var i = 0; i < data.length; i++) data[i][0] = i + 1; // renumber order

  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, loaded.lastCol).setValues(data); // single bulk write
  }
  sheet.deleteRow(loaded.lastRow); // drop the now-stale trailing row

  return respond(true, 'تم حذف السجل بنجاح من شيت ' + sheet.getName() + '.', { deletedRow: idx + 2 });
}

// Finds the most recent (bottom-most) row matching `code`. If `typeCol`
// and `typeValue` are given, a row that ALSO matches that penalty type is
// preferred; otherwise (or if no type match exists) the most recent
// code-only match is used.
function findMostRecentMatch(data, codeCol, code, typeCol, typeValue) {
  var codeOnlyIdx = -1;
  for (var r = data.length - 1; r >= 0; r--) {
    if (String(data[r][codeCol]).trim() !== String(code).trim()) continue;
    if (codeOnlyIdx === -1) codeOnlyIdx = r;
    if (typeCol !== -1 && typeValue && String(data[r][typeCol]).trim() === String(typeValue).trim()) {
      return r; // exact type match, and we're scanning newest-first
    }
  }
  return codeOnlyIdx;
}
