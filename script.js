const CALENDAR_YEAR = new Date().getFullYear();
const DEFAULT_BALANCE = 14;
const DEFAULT_EMERGENCY = 7;
const STORAGE_KEY = `holidayPlanner:${CALENDAR_YEAR}`;

let currentLanguage = localStorage.getItem('egyptairPortal:language') || localStorage.getItem('holidayPlanner:language') || 'en';
let annualBalance = DEFAULT_BALANCE;
let emergencyBalance = DEFAULT_EMERGENCY;
/** @type {Map<string, 'annual'|'emergency'|'absence'>} */
let selectedDays = new Map();

const translations = {
  en: {
    appTitle: 'Employee Holiday Planner',
    eyebrow: 'Egypt full-year planner',
    title: 'Employee Holiday Planner',
    intro: 'Choose personal leave days from January 1 to December 31 of the current year. Annual leave is used first, then emergency (عارضة), then absence (غياب).',
    balanceLabel: 'Annual holiday balance',
    balanceHint: 'Type your annual leave days.',
    emergencyLabel: 'Emergency leave (عارضة)',
    emergencyHint: 'Used after annual leave is finished.',
    takenLabel: 'Annual used',
    remainingLabel: 'Annual remaining',
    emergencyRemainingLabel: 'Emergency remaining',
    absenceLabel: 'Absence (غياب)',
    reset: 'Reset selection',
    regularDay: 'Regular day',
    officialHoliday: 'Official holiday',
    personalHoliday: 'Annual leave',
    emergencyLeave: 'Emergency (عارضة)',
    absenceDay: 'Absence (غياب)',
    weekend: 'Weekend',
    today: 'Today',
    summaryLabel: 'Leave summary',
    downloadSummary: 'Download as text',
    noDays: 'No leave days selected.',
    selectedSummary: count => `Selected leave days (${count}):`,
    annualSection: 'Annual leave:',
    emergencySection: 'Emergency (عارضة):',
    absenceSection: 'Absence (غياب):',
    balanceSummary: (annualTaken, annualTotal, annualRem, emergTaken, emergTotal, emergRem, absence) =>
      `Annual: ${annualTaken}/${annualTotal} (remaining ${annualRem}) | Emergency: ${emergTaken}/${emergTotal} (remaining ${emergRem}) | Absence: ${absence}`,
    blockedHoliday: name => `You cannot select this day because it is an official holiday: ${name}.`,
    blockedWeekend: 'You cannot select leave on Friday or Saturday.',
    resetConfirm: 'Are you sure you want to reset all selected days?',
    dateLocale: 'en-GB',
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
    weekdays: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
    holidayNames: {
      copticChristmas: 'Coptic Christmas',
      policeDay: 'January 25 Revolution / Police Day',
      policeDayObserved: 'Day off for January 25 Revolution / Police Day',
      eidFitr: 'Eid al-Fitr',
      eidFitrHoliday: 'Eid al-Fitr holiday',
      shamElNessim: 'Sham el-Nessim',
      sinaiLiberation: 'Sinai Liberation Day',
      laborDay: 'Labor Day',
      arafatDay: 'Arafat Day',
      eidAdha: 'Eid al-Adha',
      eidAdhaHoliday: 'Eid al-Adha holiday',
      hijriNewYear: 'Hijri New Year',
      june30: 'June 30 Revolution',
      june30Observed: 'Day off for June 30 Revolution',
      julyRevolution: 'July 23 Revolution Day',
      prophetBirthday: "Prophet Muhammad's Birthday",
      armedForces: 'Armed Forces Day',
      armedForcesObserved: 'Day off for Armed Forces Day',
      sinaiLiberationObserved: 'Day off for Sinai Liberation Day',
      julyRevolutionObserved: 'Day off for July 23 Revolution Day',
      laborDayObserved: 'Day off for Labor Day'
    }
  },
  ar: {
    appTitle: 'مخطط إجازات الموظفين',
    eyebrow: 'مخطط سنوي كامل لمصر',
    title: 'مخطط إجازات الموظفين',
    intro: 'اختر أيام الإجازة الشخصية من 1 يناير إلى 31 ديسمبر في السنة الحالية. تُستخدم الإجازة السنوية أولاً ثم العارضة ثم الغياب.',
    balanceLabel: 'رصيد الإجازات السنوي',
    balanceHint: 'اكتب عدد أيام الإجازة السنوية.',
    emergencyLabel: 'رصيد العارضة',
    emergencyHint: 'تُستخدم بعد انتهاء الإجازة السنوية.',
    takenLabel: 'الإجازة المستخدمة',
    remainingLabel: 'رصيد الإجازة المتبقي',
    emergencyRemainingLabel: 'رصيد العارضة المتبقي',
    absenceLabel: 'الغياب',
    reset: 'إعادة ضبط الاختيار',
    regularDay: 'يوم عادي',
    officialHoliday: 'إجازة رسمية',
    personalHoliday: 'إجازة سنوية',
    emergencyLeave: 'عارضة',
    absenceDay: 'غياب',
    weekend: 'عطلة نهاية الأسبوع',
    today: 'اليوم',
    summaryLabel: 'ملخص الإجازات',
    downloadSummary: 'تحميل كنص',
    noDays: 'لا توجد أيام إجازة مختارة.',
    selectedSummary: count => `أيام الإجازة المختارة (${count}):`,
    annualSection: 'الإجازة السنوية:',
    emergencySection: 'العارضة:',
    absenceSection: 'الغياب:',
    balanceSummary: (annualTaken, annualTotal, annualRem, emergTaken, emergTotal, emergRem, absence) =>
      `السنوية: ${annualTaken}/${annualTotal} (متبقي ${annualRem}) | العارضة: ${emergTaken}/${emergTotal} (متبقي ${emergRem}) | الغياب: ${absence}`,
    blockedHoliday: name => `لا يمكن اختيار هذا اليوم لأنه إجازة رسمية: ${name}.`,
    blockedWeekend: 'لا يمكن اختيار إجازة يوم الجمعة أو السبت.',
    resetConfirm: 'هل تريد إعادة ضبط كل الأيام المختارة؟',
    dateLocale: 'ar-EG',
    months: ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'],
    weekdays: ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'],
    holidayNames: {
      copticChristmas: 'عيد الميلاد المجيد',
      policeDay: 'ثورة 25 يناير / عيد الشرطة',
      policeDayObserved: 'بدل إجازة ثورة 25 يناير / عيد الشرطة',
      eidFitr: 'عيد الفطر',
      eidFitrHoliday: 'إجازة عيد الفطر',
      shamElNessim: 'شم النسيم',
      sinaiLiberation: 'عيد تحرير سيناء',
      laborDay: 'عيد العمال',
      arafatDay: 'وقفة عرفات',
      eidAdha: 'عيد الأضحى',
      eidAdhaHoliday: 'إجازة عيد الأضحى',
      hijriNewYear: 'رأس السنة الهجرية',
      june30: 'ثورة 30 يونيو',
      june30Observed: 'بدل إجازة ثورة 30 يونيو',
      julyRevolution: 'ثورة 23 يوليو',
      prophetBirthday: 'المولد النبوي الشريف',
      armedForces: 'عيد القوات المسلحة',
      armedForcesObserved: 'بدل إجازة عيد القوات المسلحة',
      sinaiLiberationObserved: 'بدل إجازة عيد تحرير سيناء',
      julyRevolutionObserved: 'بدل إجازة ثورة 23 يوليو',
      laborDayObserved: 'بدل إجازة عيد العمال'
    }
  }
};

const holidayDefinitions = {
  2025: [
    ['2025-01-07', 'copticChristmas'],
    ['2025-01-26', 'policeDayObserved'],
    ['2025-03-29', 'eidFitrHoliday'],
    ['2025-03-30', 'eidFitrHoliday'],
    ['2025-03-31', 'eidFitr'],
    ['2025-04-01', 'eidFitrHoliday'],
    ['2025-04-02', 'eidFitrHoliday'],
    ['2025-04-21', 'shamElNessim'],
    ['2025-04-24', 'sinaiLiberationObserved'],
    ['2025-04-25', 'sinaiLiberation'],
    ['2025-05-01', 'laborDay'],
    ['2025-06-05', 'arafatDay'],
    ['2025-06-06', 'eidAdha'],
    ['2025-06-07', 'eidAdhaHoliday'],
    ['2025-06-08', 'eidAdhaHoliday'],
    ['2025-06-09', 'eidAdhaHoliday'],
    ['2025-06-26', 'hijriNewYear'],
    ['2025-07-03', 'june30Observed'],
    ['2025-07-24', 'julyRevolutionObserved'],
    ['2025-09-04', 'prophetBirthday'],
    ['2025-10-09', 'armedForcesObserved']
  ],
  2026: [
    ['2026-01-07', 'copticChristmas'],
    ['2026-01-29', 'policeDayObserved'],
    ['2026-03-19', 'eidFitrHoliday'],
    ['2026-03-20', 'eidFitr'],
    ['2026-03-21', 'eidFitrHoliday'],
    ['2026-03-22', 'eidFitrHoliday'],
    ['2026-03-23', 'eidFitrHoliday'],
    ['2026-04-13', 'shamElNessim'],
    ['2026-04-25', 'sinaiLiberation'],
    ['2026-05-01', 'laborDay'],
    ['2026-05-26', 'arafatDay'],
    ['2026-05-27', 'eidAdha'],
    ['2026-05-28', 'eidAdhaHoliday'],
    ['2026-05-29', 'eidAdhaHoliday'],
    ['2026-05-30', 'eidAdhaHoliday'],
    ['2026-05-31', 'eidAdhaHoliday'],
    ['2026-06-16', 'hijriNewYear'],
    ['2026-07-02', 'june30Observed'],
    ['2026-07-23', 'julyRevolution'],
    ['2026-08-26', 'prophetBirthday'],
    ['2026-10-08', 'armedForcesObserved']
  ],
  2027: [
    ['2027-01-07', 'copticChristmas'],
    ['2027-01-28', 'policeDayObserved'],
    ['2027-03-09', 'eidFitr'],
    ['2027-03-10', 'eidFitrHoliday'],
    ['2027-03-11', 'eidFitrHoliday'],
    ['2027-04-25', 'sinaiLiberation'],
    ['2027-05-01', 'laborDay'],
    ['2027-05-03', 'shamElNessim'],
    ['2027-05-15', 'arafatDay'],
    ['2027-05-16', 'eidAdha'],
    ['2027-05-17', 'eidAdhaHoliday'],
    ['2027-05-18', 'eidAdhaHoliday'],
    ['2027-05-19', 'eidAdhaHoliday'],
    ['2027-06-18', 'hijriNewYear'],
    ['2027-07-01', 'june30Observed'],
    ['2027-07-23', 'julyRevolution'],
    ['2027-08-27', 'prophetBirthday'],
    ['2027-10-08', 'armedForcesObserved']
  ]
};

document.addEventListener('DOMContentLoaded', () => {
  loadState();
  initializeBalanceControls();
  initializeResetButton();
  initializeDownloadButton();
  applyLanguage(currentLanguage);
  generateCalendar();
  updateAll();
});

window.addEventListener('portal-language-change', event => {
  applyLanguage(event.detail.language);
  generateCalendar();
  updateAll();
});

let holidayMapCache = null;

function getHolidayMap() {
  if (!holidayMapCache) {
    const definitions = holidayDefinitions[CALENDAR_YEAR] || getFixedHolidayDefinitions(CALENDAR_YEAR);
    holidayMapCache = definitions.reduce((map, [date, key]) => {
      map[date] = key;
      return map;
    }, {});
  }
  return holidayMapCache;
}

function getFixedHolidayDefinitions(year) {
  const fixed = [
    [`${year}-01-07`, 'copticChristmas'],
    [`${year}-01-25`, 'policeDay'],
    [`${year}-04-25`, 'sinaiLiberation'],
    [`${year}-05-01`, 'laborDay'],
    [`${year}-06-30`, 'june30'],
    [`${year}-07-23`, 'julyRevolution'],
    [`${year}-10-06`, 'armedForces']
  ];

  const observedMap = {
    policeDay: 'policeDayObserved',
    sinaiLiberation: 'sinaiLiberationObserved',
    laborDay: 'laborDayObserved',
    june30: 'june30Observed',
    julyRevolution: 'julyRevolutionObserved',
    armedForces: 'armedForcesObserved'
  };

  const observed = [];
  fixed.forEach(([dateStr, key]) => {
    const [y, m, d] = dateStr.split('-').map(Number);
    const date = new Date(y, m - 1, d);
    const day = date.getDay();
    if (day === 5 || day === 6) {
      const offset = day === 5 ? 2 : 1;
      const observedDate = new Date(y, m - 1, d + offset);
      const oy = observedDate.getFullYear();
      const om = String(observedDate.getMonth() + 1).padStart(2, '0');
      const od = String(observedDate.getDate()).padStart(2, '0');
      observed.push([`${oy}-${om}-${od}`, observedMap[key]]);
    }
  });

  return [...fixed, ...observed];
}

function initializeBalanceControls() {
  const annualInput = document.getElementById('balanceInput');
  const emergencyInput = document.getElementById('emergencyInput');

  annualInput.value = annualBalance;
  emergencyInput.value = emergencyBalance;

  annualInput.addEventListener('input', () => setAnnualBalance(Number(annualInput.value)));
  emergencyInput.addEventListener('input', () => setEmergencyBalance(Number(emergencyInput.value)));
}

function initializeResetButton() {
  document.getElementById('resetBtn').addEventListener('click', () => {
    if (selectedDays.size === 0 || confirm(t('resetConfirm'))) {
      selectedDays.clear();
      saveState();
      generateCalendar();
      updateAll();
    }
  });
}

function initializeDownloadButton() {
  const btn = document.getElementById('downloadSummaryBtn');
  if (!btn) return;
  btn.addEventListener('click', downloadSummary);
}

function setAnnualBalance(value) {
  annualBalance = Math.max(0, Math.min(365, Number.isFinite(value) ? Math.floor(value) : 0));
  document.getElementById('balanceInput').value = annualBalance;
  reclassifySelections();
  saveState();
  generateCalendar();
  updateAll();
}

function setEmergencyBalance(value) {
  emergencyBalance = Math.max(0, Math.min(365, Number.isFinite(value) ? Math.floor(value) : 0));
  document.getElementById('emergencyInput').value = emergencyBalance;
  reclassifySelections();
  saveState();
  generateCalendar();
  updateAll();
}

/** Re-assign day types when balances change (chronological order). */
function reclassifySelections() {
  const keys = [...selectedDays.keys()].sort();
  selectedDays.clear();
  let annualUsed = 0;
  let emergUsed = 0;
  for (const key of keys) {
    if (annualUsed < annualBalance) {
      selectedDays.set(key, 'annual');
      annualUsed++;
    } else if (emergUsed < emergencyBalance) {
      selectedDays.set(key, 'emergency');
      emergUsed++;
    } else {
      selectedDays.set(key, 'absence');
    }
  }
}

function applyLanguage(language) {
  currentLanguage = language === 'ar' ? 'ar' : 'en';
  localStorage.setItem('holidayPlanner:language', currentLanguage);
  const dictionary = translations[currentLanguage];

  document.querySelectorAll('[data-i18n]').forEach(element => {
    element.textContent = dictionary[element.dataset.i18n];
  });
}

function generateCalendar() {
  const container = document.getElementById('calendarContainer');
  if (!container) return;
  container.innerHTML = '';

  for (let month = 0; month < 12; month++) {
    container.appendChild(createMonthCalendar(CALENDAR_YEAR, month));
  }
}

function createMonthCalendar(year, month) {
  const monthDiv = document.createElement('section');
  monthDiv.className = 'month-container';

  const header = document.createElement('div');
  header.className = 'month-header';
  header.innerHTML = `<div class="month-title">${t('months')[month]}</div><div class="month-year">${year}</div>`;
  monthDiv.appendChild(header);

  const weekdaysDiv = document.createElement('div');
  weekdaysDiv.className = 'weekdays';
  t('weekdays').forEach(day => {
    const dayEl = document.createElement('div');
    dayEl.className = 'weekday';
    dayEl.textContent = day;
    weekdaysDiv.appendChild(dayEl);
  });
  monthDiv.appendChild(weekdaysDiv);

  const daysGrid = document.createElement('div');
  daysGrid.className = 'days-grid';

  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  for (let i = 0; i < firstDay; i++) {
    daysGrid.appendChild(document.createElement('div'));
  }

  for (let day = 1; day <= daysInMonth; day++) {
    daysGrid.appendChild(createDayButton(year, month, day));
  }

  monthDiv.appendChild(daysGrid);
  return monthDiv;
}

function createDayButton(year, month, day) {
  const date = new Date(year, month, day);
  const dateKey = toDateKey(date);
  const holidayMap = getHolidayMap();
  const holidayKey = holidayMap[dateKey];
  const isWeekend = date.getDay() === 5 || date.getDay() === 6;
  const dayType = selectedDays.get(dateKey);
  const isToday = toDateKey(new Date()) === dateKey;
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'day';
  button.textContent = day;
  button.dataset.date = dateKey;

  if (holidayKey) {
    button.classList.add('official-holiday');
    button.title = holidayName(holidayKey);
  } else if (isWeekend) {
    button.classList.add('weekend');
  }

  if (dayType === 'annual') {
    button.classList.add('personal-holiday');
    button.title = t('personalHoliday');
  } else if (dayType === 'emergency') {
    button.classList.add('emergency-leave');
    button.title = t('emergencyLeave');
  } else if (dayType === 'absence') {
    button.classList.add('absence-day');
    button.title = t('absenceDay');
  }

  if (isToday) {
    button.classList.add('today');
    button.setAttribute('data-today', t('today'));
  }

  button.addEventListener('click', () => toggleDay(dateKey, holidayKey, isWeekend));
  return button;
}

function countByType(type) {
  let n = 0;
  for (const t of selectedDays.values()) {
    if (t === type) n++;
  }
  return n;
}

function toggleDay(dateKey, holidayKey, isWeekend) {
  if (holidayKey) {
    window.showToast(t('blockedHoliday')(holidayName(holidayKey)), 'warning');
    return;
  }
  if (isWeekend) {
    window.showToast(t('blockedWeekend'), 'warning');
    return;
  }

  if (selectedDays.has(dateKey)) {
    selectedDays.delete(dateKey);
    reclassifySelections();
  } else {
    const annualUsed = countByType('annual');
    const emergUsed = countByType('emergency');
    if (annualUsed < annualBalance) {
      selectedDays.set(dateKey, 'annual');
    } else if (emergUsed < emergencyBalance) {
      selectedDays.set(dateKey, 'emergency');
    } else {
      selectedDays.set(dateKey, 'absence');
    }
  }

  saveState();
  generateCalendar();
  updateAll();
}

function updateAll() {
  const annualTaken = countByType('annual');
  const emergTaken = countByType('emergency');
  const absenceTaken = countByType('absence');
  const annualRem = Math.max(annualBalance - annualTaken, 0);
  const emergRem = Math.max(emergencyBalance - emergTaken, 0);

  const calcVal = document.getElementById('calculatorValue');
  const currBal = document.getElementById('currentBalance');
  const emergEl = document.getElementById('emergencyRemaining');
  const absEl = document.getElementById('absenceCount');

  if (calcVal) calcVal.textContent = annualTaken;
  if (currBal) {
    currBal.textContent = annualRem;
    currBal.classList.toggle('low', annualRem === 0);
  }
  if (emergEl) {
    emergEl.textContent = emergRem;
    emergEl.classList.toggle('low', emergRem === 0);
  }
  if (absEl) {
    absEl.textContent = absenceTaken;
    absEl.classList.toggle('low', absenceTaken > 0);
  }
  updateSummary(annualTaken, annualRem, emergTaken, emergRem, absenceTaken);
}

function updateSummary(annualTaken, annualRem, emergTaken, emergRem, absenceTaken) {
  const summary = document.getElementById('summaryText');
  if (!summary) return;

  const annualKeys = [];
  const emergKeys = [];
  const absKeys = [];
  for (const [key, type] of selectedDays) {
    if (type === 'annual') annualKeys.push(key);
    else if (type === 'emergency') emergKeys.push(key);
    else absKeys.push(key);
  }
  annualKeys.sort();
  emergKeys.sort();
  absKeys.sort();

  const total = annualKeys.length + emergKeys.length + absKeys.length;
  if (total === 0) {
    summary.value = `${t('noDays')}\n${t('balanceSummary')(0, annualBalance, annualRem, 0, emergencyBalance, emergRem, 0)}`;
    return;
  }

  const lines = [t('selectedSummary')(total), ''];

  if (annualKeys.length) {
    lines.push(t('annualSection'));
    annualKeys.forEach(k => lines.push('  ' + formatDate(k)));
    lines.push('');
  }
  if (emergKeys.length) {
    lines.push(t('emergencySection'));
    emergKeys.forEach(k => lines.push('  ' + formatDate(k)));
    lines.push('');
  }
  if (absKeys.length) {
    lines.push(t('absenceSection'));
    absKeys.forEach(k => lines.push('  ' + formatDate(k)));
    lines.push('');
  }

  lines.push(t('balanceSummary')(annualTaken, annualBalance, annualRem, emergTaken, emergencyBalance, emergRem, absenceTaken));
  summary.value = lines.join('\n');
}

function downloadSummary() {
  const summary = document.getElementById('summaryText');
  if (!summary) return;
  const text = summary.value || '';
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth() + 1;
  const d = now.getDate();
  const filename = `رصيد-${y}-${m}-${d}.txt`;
  const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function loadState() {
  try {
    const state = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
    annualBalance = Number.isFinite(state.annualBalance) ? state.annualBalance : DEFAULT_BALANCE;
    emergencyBalance = Number.isFinite(state.emergencyBalance) ? state.emergencyBalance : DEFAULT_EMERGENCY;

    selectedDays = new Map();
    if (Array.isArray(state.selectedDays)) {
      // Legacy: plain array of date keys → treat as annual first
      state.selectedDays.forEach(key => {
        if (typeof key === 'string') selectedDays.set(key, 'annual');
      });
      reclassifySelections();
    } else if (state.selectedDays && typeof state.selectedDays === 'object') {
      Object.entries(state.selectedDays).forEach(([key, type]) => {
        if (type === 'annual' || type === 'emergency' || type === 'absence') {
          selectedDays.set(key, type);
        }
      });
      reclassifySelections();
    }
  } catch {
    annualBalance = DEFAULT_BALANCE;
    emergencyBalance = DEFAULT_EMERGENCY;
    selectedDays = new Map();
  }
}

function saveState() {
  const obj = {};
  for (const [k, v] of selectedDays) obj[k] = v;
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    annualBalance,
    emergencyBalance,
    selectedDays: obj
  }));
}

function t(key) {
  return translations[currentLanguage][key];
}

function holidayName(key) {
  return translations[currentLanguage].holidayNames[key] || key;
}

function toDateKey(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function formatDate(dateKey) {
  const [year, month, day] = dateKey.split('-').map(Number);
  return new Date(year, month - 1, day).toLocaleDateString(t('dateLocale'), {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
}