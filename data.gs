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

/** Apply yyyy/mm/dd display format (no time) to every date column on a row. */
function applyDateFormats(sh, row) {
  DATE_FIELDS.forEach(key => {
    const col = FIELDS.findIndex(f => f[0] === key) + 1;
    if (col > 0) {
      sh.getRange(row, col).setNumberFormat('yyyy/mm/dd');
    }
  });
}

/**
 * Sheet formulas for computed columns (1-based):
 *   M = birthDate, N = retirementDate, O = appointmentDate
 *   P = age, Q = ageGroup, R = ageAtAppointment, S = totalService
 *
 * السن الان          = whole years from birth to today
 * تاريخ المعاش       = birth + 60 years (standard pension age)
 * الفئة العمرية      = 5-year bands from current age
 * السن عند التعيين   = whole years from birth to appointment
 * مدة الخدمة الكلية  = years + months from appointment to today
 */
function writeDerivedFormulas(sh, row) {
  // N — تاريخ المعاش (birth + 60)
  sh.getRange(row, 14).setFormula(
    '=IFERROR(DATE(YEAR(M' + row + ')+' + RETIREMENT_AGE_YEARS +
      ',MONTH(M' + row + '),DAY(M' + row + ')),"")'
  );
  // P — السن الان
  sh.getRange(row, 16).setFormula(
    '=IFERROR(DATEDIF(M' + row + ',TODAY(),"Y"),"")'
  );
  // Q — الفئة العمرية (every 5 years)
  sh.getRange(row, 17).setFormula(
    '=IF(P' + row + '="","",' +
      'IF(P' + row + '<20,"أقل من 20",' +
      'IF(P' + row + '<25,"20–24",' +
      'IF(P' + row + '<30,"25–29",' +
      'IF(P' + row + '<35,"30–34",' +
      'IF(P' + row + '<40,"35–39",' +
      'IF(P' + row + '<45,"40–44",' +
      'IF(P' + row + '<50,"45–49",' +
      'IF(P' + row + '<55,"50–54",' +
      'IF(P' + row + '<60,"55–59",' +
      'IF(P' + row + '<65,"60–64","65 فأكثر"))))))))))'
  );
  // R — السن عند التعيين
  sh.getRange(row, 18).setFormula(
    '=IFERROR(DATEDIF(M' + row + ',O' + row + ',"Y"),"")'
  );
  // S — مدة الخدمة الكلية
  sh.getRange(row, 19).setFormula(
    '=IFERROR(DATEDIF(O' + row + ',TODAY(),"Y")&" سنة و "&' +
      'DATEDIF(O' + row + ',TODAY(),"YM")&" شهر","")'
  );
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
