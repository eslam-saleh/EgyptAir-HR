const CLINIC_SHEET_ID = '1LCk4J_McBFhDuRf4s-5HIirck3uwfZzCFeI3084dGig';

function doGet(e) {
  const params = e.parameter || {};
  const callback = String(params.callback || '');

  const payload = params.action === 'getData'
    ? getSheetData()
    : updateDecision(params);

  if (/^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(payload)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return jsonResponse(payload);
}

// Reads the sheet directly via the Sheets service (SpreadsheetApp), which
// returns the exact stored cell values with no per-column type inference.
// This avoids the gviz/tq endpoint's behavior of nulling out cells whose
// value (e.g. Arabic-Indic digits in a mostly-numeric column) doesn't match
// the type it auto-detected for that column.
function getSheetData() {
  try {
    const ss = SpreadsheetApp.openById(CLINIC_SHEET_ID);
    const sheet = ss.getSheets()[0];
    const timeZone = ss.getSpreadsheetTimeZone();
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
      rows = values.map(row => row.map(cell => {
        if (Object.prototype.toString.call(cell) === '[object Date]') {
          // Keep the same "M/d/yyyy HH:mm:ss" shape the client already parses.
          return Utilities.formatDate(cell, timeZone, 'M/d/yyyy HH:mm:ss');
        }
        return cell === null || cell === undefined ? '' : String(cell);
      }));
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

function updateDecision(params) {
  try {
    const rowNumber = Number(params.rowNumber);
    const decision = String(params.isReserved || '').toLowerCase();

    if (!rowNumber || !['y', 'n'].includes(decision)) {
      return { ok: false, error: 'Invalid row or decision.' };
    }

    const sheet = SpreadsheetApp.openById(CLINIC_SHEET_ID).getSheets()[0];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const reservedColumn = headers.findIndex(header => String(header).trim() === 'isReserved') + 1;

    if (!reservedColumn) {
      return { ok: false, error: 'isReserved column was not found.' };
    }

    sheet.getRange(rowNumber, reservedColumn).setValue(decision);
    return { ok: true };
  } catch (error) {
    return { ok: false, error: error.message };
  }
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}