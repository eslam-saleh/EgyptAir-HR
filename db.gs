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

// Sheet tabs read by the tagging.html "Deadlines to track" dashboard.
var DASHBOARD_SHEETS = ['اجازات', 'جزاءات', 'ايقاف'];

// Returns true if the caller supplied the correct SYNC_TOKEN, OR if no
// SYNC_TOKEN Script Property has been configured yet (opt-in — see the
// "Access control" note above).
function isAuthorized(token) {
  var required = PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN');
  if (!required) return true;
  return token === required;
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (!isAuthorized(body.token)) return respond(false, 'غير مصرح بهذا الطلب.');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(body.sheet);
    if (!sheet) return respond(false, 'لم يتم العثور على الشيت: ' + body.sheet);

    switch (body.action) {
      case 'append':        return handleAppend(sheet, body);
      case 'update':         return handleUpdate(sheet, body);
      case 'updatePenalty':  return handleUpdatePenalty(sheet, body);
      case 'deletePenalty':  return handleDeletePenalty(sheet, body);
      default: return respond(false, 'إجراء غير معروف: ' + body.action);
    }
  } catch (err) {
    return respond(false, 'خطأ في الخادم: ' + err.message);
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
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var out = {};
      DASHBOARD_SHEETS.forEach(function (name) {
        var sheet = findSheetByName(ss, name);
        out[name] = sheet ? getFilledRecords(sheet) : [];
      });
      return respond(true, 'تم تحميل البيانات.', out);
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

// ── APPEND (one or more rows in a single bulk write) ──────────────────
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

  var newRows = rows.map(function (rowData, i) {
    var out = new Array(lastCol).fill('');
    out[0] = startOrder + i; // column A = order, never overwritten below
    Object.keys(rowData).forEach(function (key) {
      var col = findHeaderCol(headers, key);
      if (col < 0) return;
      var value = rowData[key];
      out[col] = isDateHeader(headers[col]) ? normalizeDateCellText(value) : value;
    });
    return out;
  });

  var startRow = lastRow + 1;
  // These are brand-new, previously-empty cells, so it's safe to force
  // Plain text on their date columns before writing — see isDateHeader().
  for (var c = 0; c < headers.length; c++) {
    if (isDateHeader(headers[c])) {
      sheet.getRange(startRow, c + 1, newRows.length, 1).setNumberFormat('@');
    }
  }
  sheet.getRange(startRow, 1, newRows.length, lastCol).setValues(newRows); // single bulk write

  var word = newRows.length === 1 ? 'سجل واحد' : newRows.length + ' سجلات';
  return respond(true, 'تمت إضافة ' + word + ' بنجاح إلى شيت ' + sheet.getName() + '.',
    { firstRow: startRow, count: newRows.length });
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
