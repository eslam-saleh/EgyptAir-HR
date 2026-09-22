/** بيانات الموظفين — backend for the existing Employees + Retired sheets. */
const SPREADSHEET_ID = '1merEUtN-JlAFsxjffXpoqvS3SIlwaP65E27n_1fyI7Y';
const SHEET_NAME = 'Employees';
const RETIRED_SHEET_NAME = 'Retired';
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
  const sh = getSheet(SHEET_NAME);
  ensureSheetShape(sh, SHEET_NAME);
  // Ensure header row matches FIELDS (fixes old typo تاريح → تاريخ on col O).
  sh.getRange(1, 1, 1, HEADER_KEYS.length).setValues([HEADER_KEYS]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEADER_KEYS.length)
    .setFontWeight('bold')
    .setBackground('#09243f')
    .setFontColor('#ffffff')
    .setWrap(true);
  // Ensure Retired sheet exists with the same shape/headers.
  ensureRetiredSheet();
  return {
    ok: true,
    sheet: SHEET_NAME,
    retiredSheet: RETIRED_SHEET_NAME,
    rows: Math.max(0, sh.getLastRow() - 1),
    columns: HEADER_KEYS.length
  };
}

function getSheet(name) {
  const sheetName = name || SHEET_NAME;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(sheetName);
  if (!sh && sheetName === RETIRED_SHEET_NAME) {
    sh = ensureRetiredSheet();
  }
  if (!sh) throw new Error('Sheet ' + sheetName + ' was not found');
  return sh;
}

function ensureSheetShape(sh, sheetName) {
  if (sh.getLastColumn() < FIELDS.length) {
    throw new Error(
      'Sheet ' + (sheetName || sh.getName()) + ' has ' + sh.getLastColumn() +
      ' columns; ' + FIELDS.length + ' are required.'
    );
  }
}

/** Create the Retired sheet (same columns/headers as Employees) if missing. */
function ensureRetiredSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(RETIRED_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(RETIRED_SHEET_NAME);
  }
  if (sh.getLastRow() < 1 || sh.getLastColumn() < FIELDS.length) {
    // Ensure at least header width
    if (sh.getMaxColumns() < FIELDS.length) {
      sh.insertColumnsAfter(sh.getMaxColumns(), FIELDS.length - sh.getMaxColumns());
    }
  }
  sh.getRange(1, 1, 1, HEADER_KEYS.length).setValues([HEADER_KEYS]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEADER_KEYS.length)
    .setFontWeight('bold')
    .setBackground('#5c2b2b')
    .setFontColor('#ffffff')
    .setWrap(true);
  return sh;
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
    // Lean mapped records for Document Center only (active Employees only).
    if (action === 'directory') {
      return json({ ok: true, kind: 'directory', employees: readDirectoryEmployees() });
    }
    // Full sheet records for Employee Data.
    // ?action=list          → active Employees
    // ?action=list&sheet=retired  or  ?action=listRetired → Retired sheet
    if (action === 'list' || action === 'listRetired') {
      const sheetKey =
        action === 'listRetired' || String(parameters.sheet || '').toLowerCase() === 'retired'
          ? RETIRED_SHEET_NAME
          : SHEET_NAME;
      return json({
        ok: true,
        kind: sheetKey === RETIRED_SHEET_NAME ? 'listRetired' : 'list',
        sheet: sheetKey,
        employees: readEmployees(sheetKey)
      });
    }
    return json({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json({ ok: false, error: errorMessage(err) });
  }
}

function doPost(e) {
  let lock;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!isAuthorized(body.token)) return json({ ok: false, error: 'Unauthorized request.' });
    const action = body.action;
    if (
      action !== 'update' &&
      action !== 'create' &&
      action !== 'retire' &&
      action !== 'moveToRetired'
    ) {
      return json({ ok: false, error: 'Unknown action' });
    }
    lock = LockService.getScriptLock();
    if (!lock.tryLock(25000)) {
      lock = null;
      throw new Error('The employee sheet is busy. Please try again.');
    }
    let employee;
    if (action === 'retire' || action === 'moveToRetired') {
      employee = moveEmployeeToRetired(body.employee || {});
    } else if (action === 'update') {
      // Updates only apply to the active Employees sheet.
      employee = updateEmployee(body.employee || {});
    } else {
      employee = createEmployee(body.employee || {});
    }
    invalidateDirectoryCache();
    return json({ ok: true, saved: true, employee: employee });
  } catch (err) {
    return json({ ok: false, error: errorMessage(err) });
  } finally {
    if (lock) lock.releaseLock();
  }
}

function errorMessage(err) {
  return err && err.message ? err.message : String(err);
}

const DIRECTORY_CACHE_KEY = 'emp_dir_v2';
// Keep direct spreadsheet edits visible quickly; cache is only a short-lived optimization.
const DIRECTORY_CACHE_TTL = 15;
const DIRECTORY_JOB_GROUPS = [
  'مجموعة الوظائف الفنية والمكتبية',
  'مجموعة الوظائف التخصصية',
  'مجموعة الوظائف الحرفية والخدمات المعاونة'
];

function readEmployees(sheetName) {
  const name = sheetName || SHEET_NAME;
  const sh = getSheet(name);
  ensureSheetShape(sh, name);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2, 1, lastRow - 1, FIELDS.length).getDisplayValues();
  // Number rows from their REAL sheet position (i + 2, since values[0] is
  // sheet row 2), and only SKIP blank rows rather than filtering-then-
  // reindexing. The previous version filtered blanks out first and then
  // numbered what was left, so every employee after a blank row got a
  // rowNumber that pointed at the wrong sheet row. That's what let
  // updateEmployee()/resolveEmployeeRow() reject good edits with "The
  // employee row changed" for any employee sitting after a blank row.
  const employees = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (!row.some(Boolean)) continue; // skip fully blank rows, keep true row numbers
    const obj = rowToObject(row, i + 2);
    obj._sheet = name;
    employees.push(obj);
  }
  return employees;
}

function fieldIndex(key) {
  return FIELDS.findIndex(field => field[0] === key);
}

function cellAt(row, key) {
  const index = fieldIndex(key);
  return index < 0 ? '' : row[index] || '';
}

function firstNonEmpty_() {
  for (let i = 0; i < arguments.length; i++) {
    const s = String(arguments[i] == null ? '' : arguments[i]).trim();
    if (s) return s;
  }
  return '';
}

function foldDirectoryText(value) {
  return String(value || '')
    .replace(/[\u064B-\u0652\u0670\u0640]/g, '')
    .replace(/[\u002D\u2010-\u2015_]+/g, ' ')
    .replace(/[إأآا]/g, 'ا')
    .replace(/[يى]/g, 'ي')
    .replace(/[هة]/g, 'ه')
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/^ال/, '');
}

function isValidDirectoryJobGroup(value) {
  if (!value) return false;
  const key = foldDirectoryText(value);
  return DIRECTORY_JOB_GROUPS.some(group => foldDirectoryText(group) === key);
}

function isExternalSecondmentLabel(text) {
  const n = String(text || '')
    .replace(/[\u064B-\u0652\u0670\u0640]/g, '')
    .replace(/[يى]/g, 'ي')
    .toLowerCase();
  return n.indexOf('ندب خارجي') !== -1 || n.indexOf('ندب خارجى') !== -1;
}

function appendCompanyIfExternal(label, company) {
  const base = String(label || '').trim();
  if (!base || !isExternalSecondmentLabel(base)) return base;
  const co = String(company || '').trim();
  if (!co) return base;
  if (base.indexOf('(') !== -1 && base.indexOf(co) !== -1) return base;
  return base + ' (' + co + ')';
}

function directoryDate(value) {
  return formatDateYmd(parseSheetDate(value));
}

function mapDirectoryEmployee(row) {
  const codeRaw = normalizeCode(cellAt(row, 'code'));
  if (!codeRaw) return null;
  const currentCompany = firstNonEmpty_(
    cellAt(row, 'currentCompany'),
    cellAt(row, 'originalCompanyName'),
    cellAt(row, 'originalCompany')
  );
  let job = firstNonEmpty_(
    cellAt(row, 'secondedJob'),
    cellAt(row, 'secondedJob2'),
    cellAt(row, 'jobTitle'),
    cellAt(row, 'actualJob'),
    cellAt(row, 'statisticsJob')
  );
  let sector = firstNonEmpty_(cellAt(row, 'secondedSector'), cellAt(row, 'currentSector'));
  let department = firstNonEmpty_(
    cellAt(row, 'secondedGeneralDepartment'),
    cellAt(row, 'currentGeneralDepartment')
  );
  let subDepartment = firstNonEmpty_(
    cellAt(row, 'secondedDepartment'),
    cellAt(row, 'currentSubDepartment')
  );
  job = appendCompanyIfExternal(job, currentCompany);
  department = appendCompanyIfExternal(department, currentCompany);
  subDepartment = appendCompanyIfExternal(subDepartment, currentCompany);

  let group = '';
  const secondedGroup = cellAt(row, 'secondedJobGroup');
  const newGroups = cellAt(row, 'newJobGroups');
  if (isValidDirectoryJobGroup(secondedGroup)) group = String(secondedGroup).trim();
  else if (isValidDirectoryJobGroup(newGroups)) group = String(newGroups).trim();
  if (group) {
    const key = foldDirectoryText(group);
    for (let i = 0; i < DIRECTORY_JOB_GROUPS.length; i++) {
      if (foldDirectoryText(DIRECTORY_JOB_GROUPS[i]) === key) {
        group = DIRECTORY_JOB_GROUPS[i];
        break;
      }
    }
  }

  return {
    code: /^\d+$/.test(codeRaw) ? Number(codeRaw) : codeRaw,
    name: String(cellAt(row, 'name') || '').trim(),
    hireDate: directoryDate(cellAt(row, 'appointmentDate')),
    dateOfBirth: directoryDate(cellAt(row, 'birthDate')),
    sector: sector || '',
    department: department || '',
    subDepartment: subDepartment || '',
    job: job || '',
    group: group || '',
    gender: String(cellAt(row, 'gender') || '').trim(),
    currentCompany: currentCompany || ''
  };
}

function getCachedDirectory() {
  try {
    const cache = CacheService.getScriptCache();
    const n = Number(cache.get(DIRECTORY_CACHE_KEY + '_n') || 0);
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(DIRECTORY_CACHE_KEY + '_' + i);
    const fetched = cache.getAll(keys); // one round trip instead of n
    const parts = [];
    for (let i = 0; i < n; i++) {
      const part = fetched[DIRECTORY_CACHE_KEY + '_' + i];
      if (part == null) return null;
      parts.push(part);
    }
    return JSON.parse(parts.join(''));
  } catch (err) {
    return null;
  }
}

function putCachedDirectory(employees) {
  try {
    const cache = CacheService.getScriptCache();
    const str = JSON.stringify(employees);
    // CacheService limits each value to 100 KB; use a conservative character
    // size because JSON may contain multi-byte Arabic UTF-8 text.
    const chunk = 60000;
    const n = Math.ceil(str.length / chunk);
    if (!n || n > 40) return;
    const payload = {};
    payload[DIRECTORY_CACHE_KEY + '_n'] = String(n);
    for (let i = 0; i < n; i++) {
      payload[DIRECTORY_CACHE_KEY + '_' + i] = str.substr(i * chunk, chunk);
    }
    cache.putAll(payload, DIRECTORY_CACHE_TTL);
  } catch (err) {
    // Cache is optional; the live sheet read still succeeds.
  }
}

function invalidateDirectoryCache() {
  try {
    const cache = CacheService.getScriptCache();
    const n = Number(cache.get(DIRECTORY_CACHE_KEY + '_n') || 0);
    const keys = [DIRECTORY_CACHE_KEY + '_n'];
    for (let i = 0; i < n; i++) keys.push(DIRECTORY_CACHE_KEY + '_' + i);
    cache.removeAll(keys);
  } catch (err) {
    // ignore
  }
}

function readDirectoryEmployees() {
  const cached = getCachedDirectory();
  if (cached && Array.isArray(cached)) return cached;
  // Directory is always the active Employees sheet only.
  const sh = getSheet(SHEET_NAME);
  ensureSheetShape(sh, SHEET_NAME);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2, 1, lastRow - 1, FIELDS.length).getDisplayValues();
  const byCode = {};
  const employees = [];
  values.forEach(row => {
    const mapped = mapDirectoryEmployee(row);
    if (!mapped) return;
    const key = String(mapped.code);
    if (byCode[key]) return;
    byCode[key] = true;
    employees.push(mapped);
  });
  putCachedDirectory(employees);
  return employees;
}

function rowToObject(row, rowNumber) {
  const employee = { rowNumber: rowNumber };
  FIELDS.forEach((field, index) => {
    employee[field[0]] = row[index] || '';
  });
  const normalizedCode = normalizeCode(employee.code);
  if (normalizedCode) employee.code = normalizedCode;
  employee.id = employee.code || String(rowNumber);
  employee.version = rowVersion(row);
  return employee;
}

function rowVersion(row) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(row.map(value => value == null ? '' : String(value)))
  );
  return bytes.map(byte => {
    const n = byte < 0 ? byte + 256 : byte;
    return ('0' + n.toString(16)).slice(-2);
  }).join('');
}

function normalizeCode(value) {
  return String(value || '')
    .trim()
    .replace(/[٠-٩]/g, digit => String('٠١٢٣٤٥٦٧٨٩'.indexOf(digit)))
    .replace(/[۰-۹]/g, digit => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(digit)))
    .toLowerCase();
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
 * Normalize status to canonical spreadsheet values:
 *   active | inactive | leave | seconded
 * Maps legacy / typo variants from the sheet and UI:
 *   activ, على رأس العمل → active
 *   non activ, non active, غير نشط → inactive
 */
function normalizeStatus(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const s = raw.toLowerCase().replace(/\s+/g, ' ');
  // Active variants (including legacy Arabic label)
  if (
    s === 'activ' ||
    s === 'active' ||
    s.indexOf('راس العمل') !== -1 ||
    s.indexOf('رأس العمل') !== -1 ||
    s.indexOf('على رأس') !== -1 ||
    s.indexOf('على راس') !== -1
  ) {
    return 'active';
  }
  // Inactive variants (including "non activ" typo in the sheet)
  if (
    s === 'non active' ||
    s === 'non activ' ||
    s === 'non-active' ||
    s === 'nonactiv' ||
    s === 'inactive' ||
    s.indexOf('غير نشط') !== -1
  ) {
    return 'inactive';
  }
  if (s.indexOf('leave') !== -1 || s.indexOf('اجاز') !== -1 || s.indexOf('إجاز') !== -1) {
    return 'leave';
  }
  if (s.indexOf('منتدب') !== -1 || s.indexOf('second') !== -1) {
    return 'seconded';
  }
  return raw;
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

/**
 * Move an employee from the active Employees sheet to the Retired sheet.
 * Copies the full row (display values), appends to Retired, then deletes
 * the source row. Directory cache is invalidated by the caller.
 */
function moveEmployeeToRetired(input) {
  ensureRetiredSheet();
  const activeSh = getSheet(SHEET_NAME);
  ensureSheetShape(activeSh, SHEET_NAME);
  const row = resolveEmployeeRow(activeSh, input);
  const displayRow = activeSh.getRange(row, 1, 1, FIELDS.length).getDisplayValues()[0];
  if (!input.version || input.version !== rowVersion(displayRow)) {
    throw new Error(
      'This employee record changed since it was loaded. Reload the list and try again.'
    );
  }
  // Optional: normalize status on archive if client sent a reason/status.
  const values = displayRow.slice();
  const statusIdx = fieldIndex('status');
  if (statusIdx >= 0 && input.status) {
    values[statusIdx] = normalizeStatus(input.status) || values[statusIdx];
  }
  const retiredSh = getSheet(RETIRED_SHEET_NAME);
  ensureSheetShape(retiredSh, RETIRED_SHEET_NAME);
  // Append as plain values (same column order).
  retiredSh.appendRow(values);
  const newRow = retiredSh.getLastRow();
  // Keep date columns as plain text on the Retired side too.
  applyDateFormats(retiredSh, newRow);
  // Remove from active sheet (shifts rows below).
  activeSh.deleteRow(row);
  SpreadsheetApp.flush();
  const result = rowToObject(
    retiredSh.getRange(newRow, 1, 1, FIELDS.length).getDisplayValues()[0],
    newRow
  );
  result._sheet = RETIRED_SHEET_NAME;
  return result;
}

function valuesForWrite(input, existing) {
  return FIELDS.map((field, index) => {
    const key = field[0];
    if (COMPUTED.has(key)) return existing && existing[index] != null ? existing[index] : '';
    let value = input[key] === undefined || input[key] === null ? '' : input[key];
    if (key === 'code') value = normalizeCode(value);
    if (DATE_FIELDS.has(key)) value = normalizeDateOnly(value);
    if (key === 'status') value = normalizeStatus(value);
    return value;
  });
}

function findRowByCode(sh, code) {
  const target = normalizeCode(code);
  if (!target) return 0;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const codes = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < codes.length; i++) {
    if (normalizeCode(codes[i][0]) === target) return i + 2;
  }
  return 0;
}

function resolveEmployeeRow(sh, input) {
  const lastRow = sh.getLastRow();
  const requested = Number(String(input.rowNumber == null ? '' : input.rowNumber).trim());
  const rowOk = Number.isInteger(requested) && requested >= 2 && requested <= lastRow;
  if (rowOk) return requested;
  // No usable rowNumber (missing, or out of range because rows were added/
  // removed since the client loaded the list) — fall back to locating the
  // row by the employee's code.
  //
  // Deliberately NOT also checking "does this row's current code match the
  // submitted code" here: that would misfire the moment someone edits an
  // employee's own code field (a normal, supported edit) — the row is
  // still the right one, it just no longer matches the code it had at
  // load time. Whether the row changed underneath the client at all is
  // updateEmployee()'s job, via the full-row rowVersion hash check right
  // after this call, which covers a code edit correctly instead of
  // mistaking it for "wrong row".
  const code = String(input.code || '').trim();
  const byCode = findRowByCode(sh, code);
  if (byCode) return byCode;
  throw new Error('The employee row was not found.');
}

function updateEmployee(input) {
  // Edits only apply to the active Employees sheet.
  const sh = getSheet(SHEET_NAME);
  ensureSheetShape(sh, SHEET_NAME);
  const row = resolveEmployeeRow(sh, input);
  const existing = sh.getRange(row, 1, 1, FIELDS.length).getValues()[0];
  const existingDisplay = sh.getRange(row, 1, 1, FIELDS.length).getDisplayValues()[0];
  if (!input.version || input.version !== rowVersion(existingDisplay)) {
    throw new Error('This employee record changed since it was loaded. Reload the list and try again.');
  }
  validateEmployee(sh, input, row);
  const values = valuesForWrite(input, existing);
  sh.getRange(row, 1, 1, FIELDS.length).setValues([values]);
  // Force plain-text / date format on date columns so Sheets does not
  // re-attach a time component.
  applyDateFormats(sh, row);
  writeDerivedFormulas(sh, row);
  SpreadsheetApp.flush();
  const result = rowToObject(sh.getRange(row, 1, 1, FIELDS.length).getDisplayValues()[0], row);
  result._sheet = SHEET_NAME;
  return result;
}

function createEmployee(input) {
  // New employees are always created on the active Employees sheet.
  const sh = getSheet(SHEET_NAME);
  ensureSheetShape(sh, SHEET_NAME);
  validateEmployee(sh, input, 0);
  const values = valuesForWrite(input);
  sh.appendRow(values);
  const row = sh.getLastRow();
  applyDateFormats(sh, row);
  writeDerivedFormulas(sh, row);
  SpreadsheetApp.flush();
  const result = rowToObject(sh.getRange(row, 1, 1, FIELDS.length).getDisplayValues()[0], row);
  result._sheet = SHEET_NAME;
  return result;
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
  const sh = getSheet(SHEET_NAME);
  ensureSheetShape(sh, SHEET_NAME);
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
  const values = getSheet(SHEET_NAME).getDataRange().getDisplayValues();
  const csv =
    '\uFEFF' +
    values
      .map(row =>
        row.map(v => '"' + String(v ?? '').replace(/"/g, '""') + '"').join(',')
      )
      .join('\r\n');
  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
}
