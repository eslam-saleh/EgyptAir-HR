const CLINIC_SHEET_ID = '1LCk4J_McBFhDuRf4s-5HIirck3uwfZzCFeI3084dGig';
// All timestamps are delivered in Egypt time so the page can group
// submissions by Egyptian day (used by the per-day Excel download).
const EGYPT_TIME_ZONE = 'Africa/Cairo';
const PURGE_AFTER_DAYS = 40;

// ── Short-lived cache for the dashboard read ─────────────────────────────
// Form submissions write to the sheet without going through this script, so
// installClinicTriggers() (run it ONCE) adds an "On form submit" trigger that
// clears this cache the moment a submission lands. Writes made here
// (decisions / purge) clear it at once too.
const CLINIC_CACHE_KEY = 'clinic-data-v1';
const CLINIC_CACHE_TTL_SECONDS = 20;
const CLINIC_CACHE_CHUNK = 30000; // characters; Arabic is 2-3 bytes/char, limit is 100 KB per key
const CLINIC_CACHE_MAX_CHUNKS = 40;

function getSheetDataCached_() {
  const cache = CacheService.getScriptCache();
  try {
    const n = Number(cache.get(CLINIC_CACHE_KEY + '_n') || 0);
    if (n) {
      const keys = [];
      for (let i = 0; i < n; i++) keys.push(CLINIC_CACHE_KEY + '_' + i);
      const got = cache.getAll(keys);
      const parts = [];
      let complete = true;
      for (let i = 0; i < n; i++) {
        const part = got[CLINIC_CACHE_KEY + '_' + i];
        if (part == null) { complete = false; break; }
        parts.push(part);
      }
      if (complete) return JSON.parse(parts.join(''));
    }
  } catch (err) { /* fall through to a live read */ }

  const data = getSheetData();
  if (data && data.ok) {
    try {
      const str = JSON.stringify(data);
      const n = Math.ceil(str.length / CLINIC_CACHE_CHUNK);
      if (n && n <= CLINIC_CACHE_MAX_CHUNKS) {
        const payload = {};
        payload[CLINIC_CACHE_KEY + '_n'] = String(n);
        for (let i = 0; i < n; i++) {
          payload[CLINIC_CACHE_KEY + '_' + i] = str.substr(i * CLINIC_CACHE_CHUNK, CLINIC_CACHE_CHUNK);
        }
        cache.putAll(payload, CLINIC_CACHE_TTL_SECONDS);
      }
    } catch (err) { /* cache is optional */ }
  }
  return data;
}

/** Run ONCE from the Apps Script editor: new form submissions then show instantly. */
function installClinicTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'clearClinicCache') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('clearClinicCache')
    .forSpreadsheet(CLINIC_SHEET_ID)
    .onFormSubmit()
    .create();
}

function clearClinicCache() {
  try {
    const cache = CacheService.getScriptCache();
    const n = Number(cache.get(CLINIC_CACHE_KEY + '_n') || 0);
    const keys = [CLINIC_CACHE_KEY + '_n'];
    for (let i = 0; i < n; i++) keys.push(CLINIC_CACHE_KEY + '_' + i);
    cache.removeAll(keys);
  } catch (err) { /* ignore */ }
}

function doGet(e) {
  const params = e.parameter || {};
  const callback = String(params.callback || '');

  const payload = params.action === 'getData'
    ? getSheetDataCached_()
    : updateDecision(params);

  if (/^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(payload)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return jsonResponse(payload);
}

function getResponsesSheet_() {
  return SpreadsheetApp.openById(CLINIC_SHEET_ID).getSheets()[0];
}

// Keep the same "M/d/yyyy HH:mm:ss" shape the client already parses.
function formatCell_(cell) {
  if (Object.prototype.toString.call(cell) === '[object Date]') {
    return Utilities.formatDate(cell, EGYPT_TIME_ZONE, 'M/d/yyyy HH:mm:ss');
  }
  return cell === null || cell === undefined ? '' : String(cell);
}

// Reads the sheet directly via SpreadsheetApp, which returns the exact stored
// cell values with no per-column type inference (so Arabic-Indic digits in a
// mostly-numeric column are not nulled out like the gviz endpoint did).
function getSheetData() {
  try {
    const sheet = getResponsesSheet_();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 1 || lastCol < 1) {
      return { ok: true, headers: [], rows: [] };
    }

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const numDataRows = lastRow - 1;
    let rows = [];

    if (numDataRows > 0) {
      const values = sheet.getRange(2, 1, numDataRows, lastCol).getValues();
      rows = values.map(row => row.map(formatCell_));
    }

    return { ok: true, headers, rows };
  } catch (error) {
    return { ok: false, error: error.message };
  }
}

function doPost(e) {
  try {
    return jsonResponse(updateDecision(JSON.parse(e.postData.contents || '{}')));
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message });
  }
}

function findTimestampColumn_(headers) {
  const idx = headers.findIndex(h => {
    const label = String(h).trim().toLowerCase();
    return label.indexOf('timestamp') !== -1 || label.indexOf('الطابع') !== -1 || label.indexOf('الوقت') !== -1;
  });
  return idx === -1 ? 0 : idx; // Google Forms always puts the timestamp first
}

function findDateColumn_(headers) {
  return headers.findIndex(header => {
    const label = String(header).trim().toLowerCase();
    return label.indexOf('date') !== -1 || label.indexOf('التاريخ') !== -1;
  });
}

function updateDecision(params) {
  let lock;
  try {
    const rowNumber = Number(params.rowNumber);
    const decision = String(params.isReserved || '').toLowerCase();
    // Optional: admin-picked date for rows that originally had "اقرب وقت"
    const dateValue = params.date ? String(params.date).trim() : '';
    const expectedTs = String(params.ts || '').trim();

    if (!Number.isInteger(rowNumber) || !['y', 'n'].includes(decision)) {
      return { ok: false, error: 'Invalid row or decision.' };
    }
    if (!expectedTs) {
      return { ok: false, error: 'Missing row identifier. Please refresh the page and try again.' };
    }

    lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, error: 'The sheet is busy. Please try again.' };

    const sheet = getResponsesSheet_();
    // Row 1 is the header: never allow it (or anything past the data) to be written.
    if (rowNumber < 2 || rowNumber > sheet.getLastRow()) {
      return { ok: false, error: 'Row out of range. Please refresh the page.' };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const reservedColumn = headers.findIndex(header => String(header).trim() === 'isReserved') + 1;
    if (!reservedColumn) {
      return { ok: false, error: 'isReserved column was not found.' };
    }

    // Rows shift when old records are purged. Make sure this row is still the
    // same submission the admin was looking at before touching it.
    const tsColumn = findTimestampColumn_(headers) + 1;
    const actualTs = formatCell_(sheet.getRange(rowNumber, tsColumn).getValue()).trim();
    if (actualTs !== expectedTs) {
      clearClinicCache();
      return { ok: false, error: 'The sheet changed since it was loaded. Please refresh the page.' };
    }

    sheet.getRange(rowNumber, reservedColumn).setValue(decision);

    // When confirming/denying an "اقرب وقت" row, also replace the placeholder
    // with the real date so later inquiries and reloads show the chosen date.
    if (dateValue) {
      const dateColumn = findDateColumn_(headers) + 1;
      if (dateColumn > 0) {
        sheet.getRange(rowNumber, dateColumn).setValue(dateValue);
      }
    }

    SpreadsheetApp.flush();
    clearClinicCache();
    return { ok: true };
  } catch (error) {
    return { ok: false, error: error.message };
  } finally {
    if (lock) lock.releaseLock();
  }
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================================
// AUTOMATIC PURGE — run from a daily time-driven trigger.
// Uses the appointment-date column (found by header, like updateDecision).
// Rows with no real date (اقرب وقت, or clinics that skip the date question)
// fall back to the submission timestamp, so they no longer pile up forever.
// ==========================================================
function deleteOldRecords() {
  let lock;
  try {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) { Logger.log('Cleanup skipped: sheet busy.'); return; }

    const sheet = getResponsesSheet_();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return; // Empty sheet or header only

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const dateIdx = findDateColumn_(headers);       // -1 if not found
    const tsIdx = findTimestampColumn_(headers);
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const cutoffTime = Date.now() - PURGE_AFTER_DAYS * 24 * 60 * 60 * 1000;
    let deleted = 0;

    // Loop backwards so row deletions don't skip rows or break indices
    for (let i = values.length - 1; i >= 0; i--) {
      let when = null;

      if (dateIdx >= 0) {
        const cellValue = String(values[i][dateIdx]).trim();
        if (cellValue && cellValue !== 'اقرب وقت') {
          // "2026-08-16 - الأحد" -> "2026-08-16"
          const parsed = new Date(cellValue.split(' - ')[0]);
          if (!isNaN(parsed.getTime())) when = parsed.getTime();
        }
      }
      if (when === null) {
        const tsCell = values[i][tsIdx];
        const parsedTs = Object.prototype.toString.call(tsCell) === '[object Date]' ? tsCell : new Date(tsCell);
        if (!isNaN(parsedTs.getTime())) when = parsedTs.getTime();
      }

      if (when !== null && when < cutoffTime) {
        sheet.deleteRow(i + 2); // +2 for 0-index offset and header row
        deleted++;
      }
    }
    if (deleted) clearClinicCache();
    Logger.log('Cleanup completed. Deleted ' + deleted + ' row(s).');
  } catch (error) {
    Logger.log('Error running cleanup: ' + error.message);
  } finally {
    if (lock) lock.releaseLock();
  }
}
