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
    // Optional: admin-picked date for rows that originally had "اقرب وقت"
    const dateValue = params.date ? String(params.date).trim() : '';

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

    // When confirming/denying an "اقرب وقت" row, also replace the placeholder
    // with the real date so later inquiries and reloads show the chosen date.
    if (dateValue) {
      const dateColumn = headers.findIndex(header => {
        const label = String(header).trim().toLowerCase();
        return label.includes('date') || label.includes('التاريخ');
      }) + 1;

      if (dateColumn > 0) {
        sheet.getRange(rowNumber, dateColumn).setValue(dateValue);
      }
    }

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

// ==========================================================
// AUTOMATIC PURGE FUNCTION (UPDATED FOR TEXT DATE FORMATS)
// ==========================================================
function deleteOldRecords() {
  try {
    const ss = SpreadsheetApp.openById(CLINIC_SHEET_ID);
    const sheet = ss.getSheets()[0];
    const DATE_COLUMN_INDEX = 1; // 1 = Column A, 2 = Column B, etc. Change if your date isn't in Column A.
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 2 || lastCol < 1) return; // Empty sheet or header only
    
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = range.getValues();
    
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Normalize time to midnight for accurate comparison
    const cutoffTime = today.getTime() - (40 * 24 * 60 * 60 * 1000); // 40 days ago in milliseconds
    
    // Loop backwards so row deletions don't skip rows or break indices
    for (let i = values.length - 1; i >= 0; i--) {
      let cellValue = String(values[i][DATE_COLUMN_INDEX - 1]).trim();
      
      // Skip "اقرب وقت" and empty cells
      if (cellValue === "اقرب وقت" || cellValue === "") continue;
      
      // Extract the YYYY-MM-DD part before the dash (e.g., "2026-08-16 - الأحد" becomes "2026-08-16")
      let dateString = cellValue.split(" - ")[0];
      let parsedDate = new Date(dateString);
      
      if (!isNaN(parsedDate.getTime())) {
        if (parsedDate.getTime() < cutoffTime) {
          let rowToDelete = i + 2; // +2 for 0-index offset and header row
          sheet.deleteRow(rowToDelete);
        }
      }
    }
    Logger.log("Cleanup check completed successfully.");
  } catch (error) {
    Logger.log("Error running cleanup: " + error.message);
  }
}
