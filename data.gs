/** بيانات الموظفين — backend for the existing Employees sheet. */
const SPREADSHEET_ID = '1merEUtN-JlAFsxjffXpoqvS3SIlwaP65E27n_1fyI7Y';
const SHEET_NAME = 'Employees';
const FIELDS = [
  ['code', 'كود'],
  ['name', 'الاســـــــــــم '],
  ['originalCompany', ' الشركة '],
  ['currentCompany', 'الشركة الحالية'],
  ['currentSector', ' القطاع  '],
  ['currentGeneralDepartment', 'الاداره العامه'],
  ['currentSubDepartment', ' الاداره الفرعيه'],
  ['jobTitle', 'مسمى الوظيفة '],
  ['actualJob', 'الوظيفة فعلى '],
  ['statisticsJob', 'مسمى الوظيفة للاحصائيات'],
  ['uniform', 'الزى'],
  ['gender', 'النـــــــــــــوع'],
  ['birthDate', 'تاريخ الميلاد'],
  ['retirementDate', 'تاريخ المعاش'],
  ['appointmentDate', 'تاريخ التعيين'], // fixed typo: was تاريح التعيين
  ['age', 'السن الان '],
  ['ageGroup', 'الفئة العمرية '],
  ['ageAtAppointment', 'السن عند التعيين '],
  ['totalService', 'مدة الخدمة الكلية '],
  ['serviceType', 'نوع الخدمه'],
  ['workType', 'نوع العمل'],
  ['currentGrade', 'الدرجة الحالية'],
  ['gradeDate', 'تاريخ الدرجة '],
  ['appointmentJob', 'الوظيفة عند التعيين '],
  ['jobDate', 'تاريخ الوظيفة '],
  ['studySpecialization', 'التخصص الدراسى'],
  ['qualificationDate', 'تاريخ المؤهل '],
  ['currentQualificationTitle', 'مسمى المؤهل الحالى '],
  ['currentQualificationType', 'نوع المؤهل الحالى '],
  ['appointmentQualificationTitle2', 'مسمى مؤهل التعيين'],
  ['appointmentQualificationType2', 'نوع مؤهل التعيين'],
  ['appointmentQualificationTitle', 'مسمى المؤهل عند التعيين'],
  ['appointmentQualificationType', 'نوع المؤهل عند التعيين '],
  ['disabilityType', 'نوع الاعاقة'],
  ['newJobGroups', 'المجموعات النوعية  '],
  ['secondedJob', 'الوظيفة المنتدب عليها ( ان وجد )'],
  ['secondedJob2', 'الوظيفة المنتدب عليها ( ان وجد )2'],
  ['secondedCompany', 'شركة منتدب عليها'],
  ['secondedSector', 'قطاع المنتدب عليه'],
  ['secondedGeneralDepartment', 'ادارة عامة منتدب عليها'],
  ['secondedDepartment', 'الادارة المنتدب عليها'],
  ['secondmentStart', 'تاريخ الانتداب '],
  ['secondmentEnd', 'تاريخ نهاية الانتداب '],
  ['secondedJobGroup', 'المجموعة النوعية الجديدة المنتدب عليها'],
  ['unpaidLeaveType', 'نوع الاجازة بدون مرتب  ( ان وجد ) '],
  ['leaveMandatory', 'وجوبية الاجازة '],
  ['leaveStart', 'تاريخ البداية'],
  ['leaveEnd', 'تاريخ النهاية '],
  ['militaryStatus', 'نوع المعاملة العسكرية '],
  ['exemptionReason', 'سبب الاعفاء ( نهائي - مؤقت )'],
  ['contractStart', 'تاريخ بداية التعاقد'],
  ['contractEnd', 'تاريخ نهاية التعاقد'],
  ['appointmentType', 'نوع التعيين'],
  ['status', 'status'],
  ['address', 'العنوان'],
  ['phone', 'التليفون'],
  ['governorate', 'المحافظة'],
  ['addressExtra', 'الموقع'],
  ['originalCompanyName', 'مسمى الشركة الاساسية'],
  ['originalGeneralDepartment', 'الاداره العامة الأساسية'],
  ['originalSubDepartment', 'مسمى الادارة الفرعية الاساسية'],
  ['workSchedule', 'توقيت العمل']
];
const HEADER_KEYS = FIELDS.map(x => x[1]);

// Keys written by sheet formulas (not sent from the client).
// السن الان ← تاريخ الميلاد
// تاريخ المعاش ← تاريخ الميلاد + 60 سنة (سن المعاش القياسي)
// الفئة العمرية ← تاريخ الميلاد (شرائح كل 5 سنوات)
// السن عند التعيين ← تاريخ الميلاد + تاريخ التعيين
// مدة الخدمة الكلية ← تاريخ التعيين
const COMPUTED = new Set(['retirementDate', 'age', 'ageGroup', 'ageAtAppointment', 'totalService']);

// Date-only fields — any time component is stripped before write.
const DATE_FIELDS = new Set([
  'birthDate', 'retirementDate', 'appointmentDate', 'gradeDate', 'jobDate',
  'qualificationDate', 'secondmentStart', 'secondmentEnd', 'leaveStart',
  'leaveEnd', 'contractStart', 'contractEnd'
]);

// Preferred access control: set EMPLOYEE_DATA_ALLOWED_EMAILS to a
// comma-separated allowlist and deploy as "User accessing the web app" to
// your organization. Apps Script must be able to identify the signed-in user.
// EMPLOYEE_DATA_TOKEN is only a compatibility fallback for private testing;
// a token embedded in a browser page is not strong protection for HR data.
const TOKEN_PROPERTY = 'EMPLOYEE_DATA_TOKEN';
const ALLOWED_EMAILS_PROPERTY = 'EMPLOYEE_DATA_ALLOWED_EMAILS';

// Standard Egyptian civil-service retirement age used for تاريخ المعاش.
const RETIREMENT_AGE_YEARS = 60;

function setupEmployeesData() {
  const sh = getSheet();
  ensureSheetShape(sh);
  // Ensure header row matches FIELDS (fixes old typo تاريح → تاريخ on col O).
  sh.getRange(1, 1, 1, HEADER_KEYS.length).setValues([HEADER_KEYS]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEADER_KEYS.length)
    .setFontWeight('bold')
    .setBackground('#09243f')
    .setFontColor('#ffffff')
    .setWrap(true);
  return {
    ok: true,
    sheet: SHEET_NAME,
    rows: Math.max(0, sh.getLastRow() - 1),
    columns: HEADER_KEYS.length
  };
}

function getSheet() {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet Employees was not found');
  return sh;
}

function ensureSheetShape(sh) {
  if (sh.getLastColumn() < FIELDS.length) {
    throw new Error(
      `Sheet ${SHEET_NAME} has ${sh.getLastColumn()} columns; ${FIELDS.length} are required.`
    );
  }
}

function isAuthorized(token) {
  const properties = PropertiesService.getScriptProperties();
  const allowed = String(properties.getProperty(ALLOWED_EMAILS_PROPERTY) || '')
    .split(',')
    .map(email => email.trim().toLowerCase())
    .filter(Boolean);
  if (allowed.length) {
    const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    return Boolean(email) && allowed.indexOf(email) !== -1;
  }
  const required = properties.getProperty(TOKEN_PROPERTY);
  return !required || String(token || '') === required;
}

function doGet(e) {
  try {
    const parameters = (e && e.parameter) || {};
    if (!isAuthorized(parameters.token)) return json({ ok: false, error: 'Unauthorized request.' });
    const action = parameters.action || 'list';
    if (action === 'download') return csvDownload();
    if (action !== 'list') return json({ ok: false, error: 'Unknown action' });
    return json({ ok: true, employees: readEmployees() });
  } catch (err) {
    return json({ ok: false, error: errorMessage(err) });
  }
}

function doPost(e) {
  let lock;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!isAuthorized(body.token)) return json({ ok: false, error: 'Unauthorized request.' });
    if (body.action !== 'update' && body.action !== 'create') {
      return json({ ok: false, error: 'Unknown action' });
    }
    lock = LockService.getScriptLock();
    lock.waitLock(30000);
    const employee =
      body.action === 'update'
        ? updateEmployee(body.employee || {})
        : createEmployee(body.employee || {});
    return json({ ok: true, employee: employee });
  } catch (err) {
    return json({ ok: false, error: errorMessage(err) });
  } finally {
    if (lock) lock.releaseLock();
  }
}

function errorMessage(err) {
  return err && err.message ? err.message : String(err);
}

function readEmployees() {
  const sh = getSheet();
  ensureSheetShape(sh);
  const values = sh.getDataRange().getDisplayValues();
  return values
    .slice(1)
    .filter(row => row.some(Boolean))
    .map((row, index) => rowToObject(row, index + 2));
}

function rowToObject(row, rowNumber) {
  const employee = { rowNumber: rowNumber };
  FIELDS.forEach((field, index) => {
    employee[field[0]] = row[index] || '';
  });
  employee.id = employee.code || String(rowNumber);
  return employee;
}

function normalizeCode(value) {
  return String(value || '').trim().toLowerCase();
}

/**
 * Strip any time component from a date-like value.
 * Accepts Date objects, "yyyy/mm/dd", "yyyy-mm-dd", ISO strings with time, etc.
 * Returns plain "yyyy/mm/dd" text, or the original value if it is not a date.
 */
function normalizeDateOnly(value) {
  if (value === null || value === undefined || value === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    const y = value.getFullYear();
    const m = value.getMonth() + 1;
    const d = value.getDate();
    const pad = n => (n < 10 ? '0' + n : String(n));
    return y + '/' + pad(m) + '/' + pad(d);
  }
  let s = String(value).trim();
  // Drop time portion: "2020-05-12 00:00:00", "2020/05/12T14:30:00", etc.
  s = s.replace(/[T\s].*$/, '').trim();
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (!m) return value; // not a date — leave unchanged
  const pad = n => (n.length < 2 ? '0' + n : n);
  return m[1] + '/' + pad(m[2]) + '/' + pad(m[3]);
}

/**
 * Normalize status values: activ → active, non active / nonactive → inactive.
 * Other known Arabic labels are left as-is.
 */
function normalizeStatus(value) {
  const s = String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
  if (!s) return '';
  if (s === 'activ' || s === 'active') return 'active';
  if (
    s === 'non active' ||
    s === 'nonactive' ||
    s === 'non-active' ||
    s === 'inactive'
  ) {
    return 'inactive';
  }
  // Preserve original casing for Arabic / other labels
  return String(value).trim();
}

function validateEmployee(sh, input, rowNumber) {
  const code = String(input.code || '').trim();
  const name = String(input.name || '').trim();
  if (!code || !name) throw new Error('Employee code and name are required.');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const codes = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  const duplicate = codes.some(
    (entry, index) =>
      normalizeCode(entry[0]) === normalizeCode(code) && index + 2 !== rowNumber
  );
  if (duplicate) throw new Error('An employee with this code already exists.');
}

function valuesForWrite(input) {
  return FIELDS.map(field => {
    const key = field[0];
    if (COMPUTED.has(key)) return null; // formulas own these cells
    let value = input[key] === undefined ? '' : input[key];
    if (DATE_FIELDS.has(key)) value = normalizeDateOnly(value);
    if (key === 'status') value = normalizeStatus(value);
    return value;
  });
}

function updateEmployee(input) {
  const sh = getSheet();
  const row = Number(input.rowNumber);
  ensureSheetShape(sh);
  if (!Number.isInteger(row) || row < 2 || row > sh.getLastRow()) {
    throw new Error('The employee row was not found.');
  }
  validateEmployee(sh, input, row);
  const values = valuesForWrite(input);
  sh.getRange(row, 1, 1, FIELDS.length).setValues([values]);
  // Force plain-text / date format on date columns so Sheets does not
  // re-attach a time component.
  applyDateFormats(sh, row);
  writeDerivedFormulas(sh, row);
  SpreadsheetApp.flush();
  return rowToObject(sh.getRange(row, 1, 1, FIELDS.length).getDisplayValues()[0], row);
}

function createEmployee(input) {
  const sh = getSheet();
  ensureSheetShape(sh);
  validateEmployee(sh, input, 0);
  const values = valuesForWrite(input);
  sh.appendRow(values);
  const row = sh.getLastRow();
  applyDateFormats(sh, row);
  writeDerivedFormulas(sh, row);
  SpreadsheetApp.flush();
  return rowToObject(sh.getRange(row, 1, 1, FIELDS.length).getDisplayValues()[0], row);
}

/** Force date cells to plain text so Sheets never stores/shows a time component. */
function applyDateFormats(sh, row) {
  DATE_FIELDS.forEach(key => {
    const col = FIELDS.findIndex(f => f[0] === key) + 1;
    if (col > 0) {
      const cell = sh.getRange(row, col);
      const parsed = parseSheetDate(cell.getValue());
      cell.setNumberFormat('@');
      cell.setValue(parsed ? formatDateYmd(parsed) : '');
    }
  });
}

/**
 * Parse a cell value into a Date at local midnight, or null.
 * Handles Date objects, "yyyy-mm-dd", "yyyy/mm/dd", and values with a time part.
 */
function parseSheetDate(value) {
  if (value === null || value === undefined || value === '') return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  const s = String(value).trim().replace(/[T\s].*$/, '');
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(d.getTime()) ? null : d;
}

function ageInYears(fromDate, toDate) {
  if (!fromDate || !toDate || toDate < fromDate) return '';
  let age = toDate.getFullYear() - fromDate.getFullYear();
  if (
    toDate.getMonth() < fromDate.getMonth() ||
    (toDate.getMonth() === fromDate.getMonth() && toDate.getDate() < fromDate.getDate())
  ) {
    age--;
  }
  return age >= 0 && age < 120 ? age : '';
}

function ageGroupLabel(age) {
  if (age === '' || age == null) return '';
  const n = Number(age);
  if (isNaN(n)) return '';
  if (n < 20) return 'أقل من 20';
  if (n < 25) return '20-24';
  if (n < 30) return '25-29';
  if (n < 35) return '30-34';
  if (n < 40) return '35-39';
  if (n < 45) return '40-44';
  if (n < 50) return '45-49';
  if (n < 55) return '50-54';
  if (n < 60) return '55-59';
  if (n < 65) return '60-64';
  return '65 فأكثر';
}

function serviceDurationLabel(appointmentDate, today) {
  if (!appointmentDate || appointmentDate > today) return '';
  let months =
    (today.getFullYear() - appointmentDate.getFullYear()) * 12 +
    (today.getMonth() - appointmentDate.getMonth());
  if (today.getDate() < appointmentDate.getDate()) months--;
  if (months < 0) return '';
  const years = Math.floor(months / 12);
  const rest = months % 12;
  if (years === 0 && rest === 0) return 'أقل من شهر';
  if (years === 0) return rest + ' شهر';
  if (rest === 0) return years + ' سنة';
  return years + ' سنة و ' + rest + ' شهر';
}

function formatDateYmd(date) {
  if (!date) return '';
  const pad = n => (n < 10 ? '0' + n : String(n));
  return date.getFullYear() + '/' + pad(date.getMonth() + 1) + '/' + pad(date.getDate());
}

/**
 * Compute derived values from birth (col M) and appointment (col O).
 * Writes plain VALUES (not formulas) into N, P, Q, R, S so Arabic-locale
 * sheets never hit #ERROR! from comma-based formula syntax.
 *
 *   N تاريخ المعاش      = birth + 60 years
 *   P السن الان         = years from birth to today
 *   Q الفئة العمرية     = 5-year band from current age
 *   R السن عند التعيين  = years from birth to appointment
 *   S مدة الخدمة الكلية = years + months from appointment to today
 */
function writeDerivedFormulas(sh, row) {
  const birth = parseSheetDate(sh.getRange(row, 13).getValue()); // M
  const appointment = parseSheetDate(sh.getRange(row, 15).getValue()); // O
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const age = ageInYears(birth, today);
  const pension = birth
    ? new Date(birth.getFullYear() + RETIREMENT_AGE_YEARS, birth.getMonth(), birth.getDate())
    : null;
  const ageAtAppt = ageInYears(birth, appointment);
  const group = ageGroupLabel(age);
  const service = serviceDurationLabel(appointment, today);

  // N P Q R S = columns 14,16,17,18,19 — plain values only (no formulas)
  const pensionCell = sh.getRange(row, 14);
  pensionCell.setNumberFormat('@');
  pensionCell.setValue(formatDateYmd(pension));
  sh.getRange(row, 16).setValue(age === '' ? '' : age);
  sh.getRange(row, 17).setValue(group);
  sh.getRange(row, 18).setValue(ageAtAppt === '' ? '' : ageAtAppt);
  sh.getRange(row, 19).setValue(service);
}

/**
 * One-time backfill for EVERY existing row.
 *
 * ONLY touches specific columns — never rewrites a whole row —
 * so other fields cannot shift or pick up values from the wrong column.
 *
 *  - N/P/Q/R/S recalculated as plain values (fixes #ERROR! and bad ages like 126)
 *  - all date columns forced to plain-text yyyy/mm/dd (no 00:00:00)
 *  - status: activ → active
 *  - header labels refreshed (تاريح → تاريخ)
 *
 * Run once: select backfillDerivedFormulas → Run. Safe to re-run.
 */
function backfillDerivedFormulas() {
  const sh = getSheet();
  ensureSheetShape(sh);
  // Header labels only — does not move data
  sh.getRange(1, 1, 1, HEADER_KEYS.length).setValues([HEADER_KEYS]);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return { ok: true, updated: 0, message: 'No data rows to update.' };
  }

  const numRows = lastRow - 1;
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Read ONLY the source columns we need
  const codeVals = sh.getRange(2, 1, numRows, 1).getDisplayValues();
  const birthVals = sh.getRange(2, 13, numRows, 1).getValues(); // M
  const apptVals = sh.getRange(2, 15, numRows, 1).getValues(); // O
  const statusVals = sh.getRange(2, 54, numRows, 1).getValues(); // status

  const pensionOut = [];
  const ageOut = [];
  const groupOut = [];
  const ageAtOut = [];
  const serviceOut = [];
  const statusOut = [];
  let updated = 0;

  for (let i = 0; i < numRows; i++) {
    if (!String(codeVals[i][0] || '').trim()) {
      pensionOut.push(['']);
      ageOut.push(['']);
      groupOut.push(['']);
      ageAtOut.push(['']);
      serviceOut.push(['']);
      statusOut.push([statusVals[i][0]]);
      continue;
    }

    const birth = parseSheetDate(birthVals[i][0]);
    const appointment = parseSheetDate(apptVals[i][0]);
    const age = ageInYears(birth, today);
    const pension = birth
      ? new Date(birth.getFullYear() + RETIREMENT_AGE_YEARS, birth.getMonth(), birth.getDate())
      : null;
    const ageAt = ageInYears(birth, appointment);

    pensionOut.push([formatDateYmd(pension)]);
    ageOut.push([age === '' ? '' : age]);
    groupOut.push([ageGroupLabel(age)]);
    ageAtOut.push([ageAt === '' ? '' : ageAt]);
    serviceOut.push([serviceDurationLabel(appointment, today)]);
    statusOut.push([normalizeStatus(statusVals[i][0])]);
    updated++;
  }

  // Write ONLY derived + status columns (in-place, no full-row rewrite)
  const nRange = sh.getRange(2, 14, numRows, 1);
  nRange.setNumberFormat('@');
  nRange.setValues(pensionOut);

  sh.getRange(2, 16, numRows, 1).setValues(ageOut);
  sh.getRange(2, 17, numRows, 1).setValues(groupOut);
  sh.getRange(2, 18, numRows, 1).setValues(ageAtOut);
  sh.getRange(2, 19, numRows, 1).setValues(serviceOut);
  sh.getRange(2, 54, numRows, 1).setValues(statusOut);

  // Strip time from every date column: store as plain text yyyy/mm/dd
  DATE_FIELDS.forEach(key => {
    const col = FIELDS.findIndex(f => f[0] === key) + 1;
    if (col < 1) return;
    const r = sh.getRange(2, col, numRows, 1);
    const raw = r.getValues();
    const cleaned = raw.map(cell => {
      const parsed = parseSheetDate(cell[0]);
      if (parsed) {
        const year = parsed.getFullYear();
        // Drop nonsense dates (e.g. 1111-11-11)
        if (year < 1920 || year > 2100) return [''];
        return [formatDateYmd(parsed)];
      }
      return [''];
    });
    r.setNumberFormat('@'); // plain text — Sheets will not re-add 00:00:00
    r.setValues(cleaned);
  });

  SpreadsheetApp.flush();
  return {
    ok: true,
    updated: updated,
    message:
      'Recalculated ' +
      updated +
      ' row(s) in-place (derived cols only). Dates stored as plain yyyy/mm/dd text with no time. Status activ→active.'
  };
}

function json(x) {
  return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function csvDownload() {
  const values = getSheet().getDataRange().getDisplayValues();
  const csv =
    '\uFEFF' +
    values
      .map(row =>
        row.map(v => '"' + String(v ?? '').replace(/"/g, '""') + '"').join(',')
      )
      .join('\r\n');
  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
}
