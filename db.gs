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
 * calling the Sheets API per-row.
 *
 * doGet(?action=dashboard) is read-only and powers tagging.html's
 * "Deadlines to track" cards (اجازات / جزاءات / ايقاف) — see
 * DASHBOARD_SHEETS, findSheetByName() and getFilledRecords() below.
 * If this script is already deployed, changes here only go live once you
 * redeploy: Deploy ▸ Manage deployments ▸ (pencil icon) ▸ Version: New
 * version ▸ Deploy. The /exec URL itself doesn't change.
 */

// Sheet tabs read by the tagging.html "Deadlines to track" dashboard.
var DASHBOARD_SHEETS = ['اجازات', 'جزاءات', 'ايقاف'];

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
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
// Dates that come back as real Date objects are formatted as yyyy-MM-dd
// strings so JSON.stringify doesn't need to touch them; dates already
// stored as plain text (e.g. after migrateSwapDayYear) pass through as-is.
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
        val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
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
      if (col > 0) out[col] = rowData[key];
    });
    return out;
  });

  var startRow = lastRow + 1;
  sheet.getRange(startRow, 1, newRows.length, lastCol).setValues(newRows); // single bulk write

  var word = newRows.length === 1 ? 'سجل واحد' : newRows.length + ' سجلات';
  return respond(true, 'تمت إضافة ' + word + ' بنجاح إلى شيت ' + sheet.getName() + '.',
    { firstRow: startRow, count: newRows.length });
}

// ── UPDATE (نهاية الإيقاف: fills the newest still-open row for a code) ──
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
      data[r][targetCol] = body.value;
      sheet.getRange(2, 1, data.length, loaded.lastCol).setValues(data); // single bulk write
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

// ═══════════════════════════════════════════════════════════════════════
// ONE-TIME MIGRATION — swap day/year in existing date cells.
//
// Run this ONCE manually: open this project from Extensions ▸ Apps Script,
// pick "migrateSwapDayYear" in the function dropdown next to Run, and
// click Run. It converts existing dd/mm/yyyy values to yyyy/mm/dd across
// the known date columns (التاريخ, نهاية الاجازة, بداية الإيقاف,
// نهاية الإيقاف) on every sheet, using one bulk read + one bulk write
// per sheet. Unknown/unrelated columns are never touched.
//
// It's safe if run by accident a second time: once a cell is already in
// yyyy/mm/dd form its first segment has 4 digits, which no longer matches
// the "old format" pattern below, so it's simply skipped.
// ═══════════════════════════════════════════════════════════════════════
function migrateSwapDayYear() {
  var sheetNames = ['اجازات', 'جزاءات', 'ايقاف'];
  var dateHeaderKeywords = ['التاريخ', 'نهاية الاجازة', 'بداية الإيقاف', 'نهاية الإيقاف'];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = [];

  sheetNames.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var loaded = loadSheet(sheet);
    if (!loaded.data.length) return;

    var dateCols = [];
    dateHeaderKeywords.forEach(function (kw) {
      var col = findHeaderCol(loaded.headers, kw);
      if (col !== -1 && dateCols.indexOf(col) === -1) dateCols.push(col);
    });
    if (!dateCols.length) return;

    var changed = 0;
    var data = loaded.data;
    for (var r = 0; r < data.length; r++) {
      for (var c = 0; c < dateCols.length; c++) {
        var col = dateCols[c];
        var m = String(data[r][col]).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (!m) continue; // not old-format, or already migrated — leave as-is
        var dd = ('0' + m[1]).slice(-2);
        var mm = ('0' + m[2]).slice(-2);
        data[r][col] = m[3] + '/' + mm + '/' + dd;
        changed++;
      }
    }

    sheet.getRange(2, 1, data.length, loaded.lastCol).setValues(data); // single bulk write
    report.push(name + ': ' + changed + ' خلية');
  });

  var summary = report.length ? report.join(' | ') : 'لا توجد خلايا لتحديثها.';
  Logger.log(summary);
  try {
    SpreadsheetApp.getUi().alert('تم تبديل الأيام بالسنوات:\n' + report.join('\n'));
  } catch (uiErr) {
    // getUi() is unavailable in some execution contexts — the Logger.log
    // output above still has the full report.
  }
}

// ═══════════════════════════════════════════════════════════════════════
// ONE-TIME MIGRATION — merge لائحة + قانوني into a single جزاءات sheet.
//
// Run this ONCE, the same way as migrateSwapDayYear above (pick
// "mergeLailQanoniIntoJazaat" in the Run dropdown). It creates a جزاءات
// sheet if one doesn't exist yet (copying قانوني's/لائحة's header row and
// appending "الامر التنفيذي" as the 7th column), then copies every row
// from لائحة and قانوني into it — tagging each with its source sheet name
// — renumbering column A sequentially as it goes. Nothing in لائحة or
// قانوني is modified or deleted; review the merged data in جزاءات, then
// delete those two old sheets yourself once you're happy with it.
// ═══════════════════════════════════════════════════════════════════════
function mergeLailQanoniIntoJazaat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lail = ss.getSheetByName('لائحة');
  var qanoni = ss.getSheetByName('قانوني');
  var target = ss.getSheetByName('جزاءات');
  if (!target) target = ss.insertSheet('جزاءات');

  if (target.getLastRow() < 1) {
    var baseHeaders = qanoni ? loadSheet(qanoni).headers : (lail ? loadSheet(lail).headers : []);
    var headers = baseHeaders.concat(['الامر التنفيذي']);
    target.getRange(1, 1, 1, headers.length).setValues([headers]); // single write
  }

  var targetLoaded = loadSheet(target);
  var targetHeaders = targetLoaded.headers;
  var lastCol = targetHeaders.length;
  var nextOrder = targetLoaded.data.length
    ? (Number(targetLoaded.data[targetLoaded.data.length - 1][0]) || 0) + 1
    : 1;
  var srcTypeCol = findHeaderCol(targetHeaders, 'الامر التنفيذي');

  var rowsToAppend = [];
  [{ sheet: lail, label: 'لائحة' }, { sheet: qanoni, label: 'قانوني' }].forEach(function (src) {
    if (!src.sheet) return;
    var loaded = loadSheet(src.sheet); // single bulk read per source sheet
    loaded.data.forEach(function (srcRow) {
      var out = new Array(lastCol).fill('');
      out[0] = nextOrder++;
      loaded.headers.forEach(function (h, i) {
        if (i === 0) return; // skip the source's own order column
        var col = findHeaderCol(targetHeaders, h);
        if (col > 0) out[col] = srcRow[i];
      });
      if (srcTypeCol > 0) out[srcTypeCol] = src.label;
      rowsToAppend.push(out);
    });
  });

  if (rowsToAppend.length) {
    var startRow = targetLoaded.data.length + 2;
    target.getRange(startRow, 1, rowsToAppend.length, lastCol).setValues(rowsToAppend); // single bulk write
  }

  Logger.log('Merged ' + rowsToAppend.length + ' rows into جزاءات.');
  try {
    SpreadsheetApp.getUi().alert(
      'تم دمج ' + rowsToAppend.length + ' سجل فى شيت جزاءات.\n' +
      'راجع البيانات ثم احذف شيتى لائحة و قانوني يدويًا.'
    );
  } catch (uiErr) {
    // getUi() unavailable in some contexts — see Logger.log above.
  }
}
