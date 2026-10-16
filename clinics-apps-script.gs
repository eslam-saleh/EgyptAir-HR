const CLINIC_SHEET_ID = '1LCk4J_McBFhDuRf4s-5HIirck3uwfZzCFeI3084dGig';
// All timestamps are delivered in Egypt time so the page can group
// submissions by Egyptian day (used by the per-day Excel download).
const EGYPT_TIME_ZONE = 'Africa/Cairo';
const PURGE_AFTER_DAYS = 40;

// ── Stable submission ID ─────────────────────────────────────────────────
// Timestamps only have 1-second resolution, so two submissions in the same
// second look identical, and row numbers shift when old rows are purged.
// Every response therefore gets its own random ID in this extra column
// (created automatically, to the right of the form's columns). Decisions and
// picked dates are matched on this ID, never on the timestamp alone.
const SUBMISSION_ID_HEADER = 'submissionId';

function findSubmissionIdColumn_(headers) {
  return headers.findIndex(h => String(h).trim().toLowerCase() === SUBMISSION_ID_HEADER.toLowerCase());
}

// Makes sure the sheet has the submissionId column and that every response
// row has a unique, non-empty ID (blank or duplicated IDs get a fresh one).
// Existing IDs are never changed. The caller must hold the script lock.
// Returns true if anything was written.
function ensureSubmissionIds_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return false;
  let changed = false;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let idCol = findSubmissionIdColumn_(headers) + 1;
  if (!idCol) {
    idCol = lastCol + 1;
    if (idCol > sheet.getMaxColumns()) sheet.insertColumnAfter(sheet.getMaxColumns());
    sheet.getRange(1, idCol).setValue(SUBMISSION_ID_HEADER);
    changed = true;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return changed;

  const range = sheet.getRange(2, idCol, lastRow - 1, 1);
  const ids = range.getValues();
  const seen = {};
  let dirty = false;
  for (let i = 0; i < ids.length; i++) {
    let id = String(ids[i][0]).trim();
    if (!id || seen[id]) {
      id = Utilities.getUuid();
      ids[i][0] = id;
      dirty = true;
    }
    seen[id] = true;
  }
  if (dirty) {
    range.setNumberFormat('@'); // keep it as plain text
    range.setValues(ids);
    changed = true;
  }
  return changed;
}

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

/** Run ONCE from the Apps Script editor (run it again after updating this script):
 *  gives every existing response an ID, and installs the "On form submit"
 *  trigger that gives each NEW response an ID and refreshes the dashboard. */
function installClinicTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'clearClinicCache' || fn === 'onClinicFormSubmit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onClinicFormSubmit')
    .forSpreadsheet(CLINIC_SHEET_ID)
    .onFormSubmit()
    .create();
  onClinicFormSubmit(); // backfill IDs for rows that already exist
}

// Runs on every form submission. It scans the whole sheet (not just the new
// row), so even several submissions landing at once all end up with an ID.
function onClinicFormSubmit() {
  let lock;
  try {
    lock = LockService.getScriptLock();
    if (lock.tryLock(30000)) {
      ensureSubmissionIds_(getResponsesSheet_());
      SpreadsheetApp.flush();
    } else {
      Logger.log('ID assignment skipped: sheet busy (the next dashboard load will retry).');
    }
  } catch (error) {
    Logger.log('Error assigning submission IDs: ' + error.message);
  } finally {
    if (lock) lock.releaseLock();
    clearClinicCache();
  }
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

// in doGet, replace the `payload` line with:
const payload = params.action === 'getData' ? getSheetDataCached_()
  : params.action === 'lookup' ? lookupByCode_(params.id)
  : updateDecision(params);


  if (/^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(payload)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return jsonResponse(payload);
}

function lookupByCode_(rawId) {
  const id = String(rawId || '').replace(/[٠-٩]/g, d => d.charCodeAt(0) - 1632).trim();
  if (!id || id.length > 20) return { ok: false, error: 'Invalid ID.' };
  const data = getSheetDataCached_();
  if (!data || !data.ok) return data;
  const low = data.headers.map(h => String(h).toLowerCase());
  const codeIdx = low.findIndex(h => h.includes('code') || h.includes('كود'));
  if (codeIdx < 0) return { ok: false, error: 'Code column not found.' };
  const keep = low.map((h, i) => (h.includes('email') || h.includes('البريد') || h === SUBMISSION_ID_HEADER.toLowerCase()) ? -1 : i).filter(i => i >= 0);
  const rows = data.rows
    .filter(r => String(r[codeIdx]).replace(/[٠-٩]/g, d => d.charCodeAt(0) - 1632).trim() === id)
    .map(r => keep.map(i => r[i]));
  return { ok: true, headers: keep.map(i => data.headers[i]), rows };
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

// Reads the header row and the raw data rows (no formatting applied).
function readSheetRaw_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return null;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  return { headers, values };
}

// Reads the sheet directly via SpreadsheetApp, which returns the exact stored
// cell values with no per-column type inference (so Arabic-Indic digits in a
// mostly-numeric column are not nulled out like the gviz endpoint did).
// If any row is still missing its submission ID (e.g. the form trigger has not
// run yet), IDs are assigned first so the page always receives them.
function getSheetData() {
  try {
    const sheet = getResponsesSheet_();
    let raw = readSheetRaw_(sheet);
    if (!raw) return { ok: true, headers: [], rows: [] };

    const idIdx = findSubmissionIdColumn_(raw.headers);
    const missing = idIdx < 0 || raw.values.some(r => !String(r[idIdx]).trim());
    if (missing) {
      const lock = LockService.getScriptLock();
      if (lock.tryLock(10000)) {
        try {
          if (ensureSubmissionIds_(sheet)) SpreadsheetApp.flush();
        } finally {
          lock.releaseLock();
        }
        raw = readSheetRaw_(sheet) || raw;
      }
    }

    return { ok: true, headers: raw.headers, rows: raw.values.map(row => row.map(formatCell_)) };
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
    const decision = String(params.isReserved || '').toLowerCase();
    // Optional: admin-picked date for rows that originally had "اقرب وقت"
    const dateValue = params.date ? String(params.date).trim() : '';
    const sid = String(params.sid || '').trim();
    const expectedTs = String(params.ts || '').trim();

    if (!['y', 'n'].includes(decision)) {
      return { ok: false, error: 'Invalid decision.' };
    }
    if (!sid || sid.length > 100) {
      return { ok: false, error: 'Missing submission ID. Please refresh the page and try again.' };
    }

    lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, error: 'The sheet is busy. Please try again.' };

    const sheet = getResponsesSheet_();
    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const reservedColumn = headers.findIndex(header => String(header).trim() === 'isReserved') + 1;
    if (!reservedColumn) {
      return { ok: false, error: 'isReserved column was not found.' };
    }
    const idColumn = findSubmissionIdColumn_(headers) + 1;
    if (!idColumn || lastRow < 2) {
      clearClinicCache();
      return { ok: false, error: 'The sheet changed since it was loaded. Please refresh the page.' };
    }

    // Locate the row by its unique submission ID. Row numbers shift when old
    // records are purged and timestamps can repeat, so neither is trusted to
    // pick the row. Row 1 is the header and is never a candidate.
    const ids = sheet.getRange(2, idColumn, lastRow - 1, 1).getValues();
    const matches = [];
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === sid) matches.push(i + 2);
    }
    if (matches.length !== 1) {
      clearClinicCache();
      return {
        ok: false,
        error: matches.length === 0
          ? 'This record no longer exists. Please refresh the page.'
          : 'Duplicate submission ID found. Please refresh the page and try again.'
      };
    }
    const rowNumber = matches[0];

    // Extra safety net: when the page also sent the timestamp, it must match.
    if (expectedTs) {
      const tsColumn = findTimestampColumn_(headers) + 1;
      const actualTs = formatCell_(sheet.getRange(rowNumber, tsColumn).getValue()).trim();
      if (actualTs !== expectedTs) {
        clearClinicCache();
        return { ok: false, error: 'The sheet changed since it was loaded. Please refresh the page.' };
      }
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
