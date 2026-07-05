const CLINIC_SHEET_ID = '1LCk4J_McBFhDuRf4s-5HIirck3uwfZzCFeI3084dGig';

function doGet(e) {
  const payload = updateDecision(e.parameter || {});
  const callback = String(e.parameter.callback || '');

  if (/^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(payload)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return jsonResponse(payload);
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
