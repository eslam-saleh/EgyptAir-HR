const CALENDAR_YEAR = new Date().getFullYear();
const DEFAULT_BALANCE = 14;
const STORAGE_KEY = `holidayPlanner:${CALENDAR_YEAR}`;

let currentLanguage = localStorage.getItem('egyptairPortal:language') || localStorage.getItem('holidayPlanner:language') || 'en';
let annualBalance = DEFAULT_BALANCE;
let selectedDays = new Set();

const translations = {
  en: {
    appTitle: 'Employee Holiday Planner',
    eyebrow: 'Egypt full-year planner',
    title: 'Employee Holiday Planner',
    intro: 'Choose personal leave days from January 1 to December 31 of the current year.',
    balanceLabel: 'Annual holiday balance',
    balanceHint: 'Pick from the list or type a custom number.',
    takenLabel: 'Days selected',
    remainingLabel: 'Remaining balance',
    reset: 'Reset selection',
    regularDay: 'Regular day',
    officialHoliday: 'Official holiday',
    personalHoliday: 'Personal holiday',
    weekend: 'Weekend',
    today: 'Today',
    summaryLabel: 'Leave summary',
    noDays: 'No personal holidays selected.',
    selectedSummary: count => `Selected personal holidays (${count}):`,
    balanceSummary: (taken, total, remaining) => `Taken: ${taken} day(s) | Balance: ${total} day(s) | Remaining: ${remaining} day(s)`,
    blockedHoliday: name => `You cannot select this day because it is an official holiday: ${name}.`,
    blockedWeekend: 'You cannot select personal holidays on Friday or Saturday.',
    limitReached: total => `You have reached your annual balance of ${total} day(s).`,
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
    intro: 'اختر أيام الإجازة الشخصية من 1 يناير إلى 31 ديسمبر في السنة الحالية.',
    balanceLabel: 'رصيد الإجازات السنوي',
    balanceHint: 'اختر من القائمة أو اكتب رقما مخصصا.',
    takenLabel: 'الأيام المختارة',
    remainingLabel: 'الرصيد المتبقي',
    reset: 'إعادة ضبط الاختيار',
    regularDay: 'يوم عادي',
    officialHoliday: 'إجازة رسمية',
    personalHoliday: 'إجازة شخصية',
    weekend: 'عطلة نهاية الأسبوع',
    today: 'اليوم',
    summaryLabel: 'ملخص الإجازات',
    noDays: 'لا توجد إجازات شخصية مختارة.',
    selectedSummary: count => `الإجازات الشخصية المختارة (${count}):`,
    balanceSummary: (taken, total, remaining) => `المستخدم: ${taken} يوم | الرصيد: ${total} يوم | المتبقي: ${remaining} يوم`,
    blockedHoliday: name => `لا يمكن اختيار هذا اليوم لأنه إجازة رسمية: ${name}.`,
    blockedWeekend: 'لا يمكن اختيار إجازة شخصية يوم الجمعة أو السبت.',
    limitReached: total => `لقد وصلت إلى رصيدك السنوي وهو ${total} يوم.`,
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
    ['2025-01-25', 'policeDay'],
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
    ['2025-06-30', 'june30'],
    ['2025-07-03', 'june30Observed'],
    ['2025-07-23', 'julyRevolution'],
    ['2025-07-24', 'julyRevolutionObserved'],
    ['2025-09-04', 'prophetBirthday'],
    ['2025-10-06', 'armedForces'],
    ['2025-10-09', 'armedForcesObserved']
  ],
  2026: [
    ['2026-01-07', 'copticChristmas'],
    ['2026-01-25', 'policeDay'],
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
    ['2026-06-30', 'june30'],
    ['2026-07-02', 'june30Observed'],
    ['2026-07-23', 'julyRevolution'],
    ['2026-08-26', 'prophetBirthday'],
    ['2026-10-06', 'armedForces'],
    ['2026-10-08', 'armedForcesObserved']
  ],
  2027: [
    ['2027-01-07', 'copticChristmas'],
    ['2027-01-25', 'policeDay'],
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
    ['2027-06-30', 'june30'],
    ['2027-07-01', 'june30Observed'],
    ['2027-07-23', 'julyRevolution'],
    ['2027-08-27', 'prophetBirthday'],
    ['2027-10-06', 'armedForces'],
    ['2027-10-08', 'armedForcesObserved']
  ]
};

document.addEventListener('DOMContentLoaded', () => {
  loadState();
  initializeBalanceControls();
  initializeResetButton();
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
    if (day === 5 || day === 6) { // Friday or Saturday
      const offset = day === 5 ? 2 : 1; // next Sunday
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
  const select = document.getElementById('balanceSelect');
  const input = document.getElementById('balanceInput');
  const balanceOptions = [7, 10, 14, 15, 18, 21, 24, 30];

  select.innerHTML = '';
  balanceOptions.forEach(value => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });

  select.value = balanceOptions.includes(annualBalance) ? String(annualBalance) : '';
  input.value = annualBalance;

  select.addEventListener('change', () => setAnnualBalance(Number(select.value)));
  input.addEventListener('input', () => setAnnualBalance(Number(input.value)));
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

function setAnnualBalance(value) {
  annualBalance = Math.max(0, Math.min(365, Number.isFinite(value) ? Math.floor(value) : 0));
  document.getElementById('balanceInput').value = annualBalance;
  const select = document.getElementById('balanceSelect');
  if (select) {
    select.value = [...select.options].some(o => o.value === String(annualBalance)) ? String(annualBalance) : '';
  }
  saveState();
  updateAll();
}

function applyLanguage(language) {
  currentLanguage = language === 'ar' ? 'ar' : 'en';
  localStorage.setItem('holidayPlanner:language', currentLanguage);
  const dictionary = translations[currentLanguage];
  // document.title = dictionary.appTitle; // Handled by portal header

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
  const isSelected = selectedDays.has(dateKey);
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
  
  if (isSelected) button.classList.add('personal-holiday');
  if (isToday) {
    button.classList.add('today');
    button.setAttribute('data-today', t('today'));
  }

  button.addEventListener('click', () => toggleDay(dateKey, holidayKey, isWeekend));
  return button;
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
  } else {
    if (selectedDays.size >= annualBalance) {
      window.showToast(t('limitReached')(annualBalance), 'warning');
      return;
    }
    selectedDays.add(dateKey);
  }

  saveState();
  generateCalendar();
  updateAll();
}

function updateAll() {
  const taken = selectedDays.size;
  const remaining = Math.max(annualBalance - taken, 0);
  const calcVal = document.getElementById('calculatorValue');
  const currBal = document.getElementById('currentBalance');
  
  if (calcVal) calcVal.textContent = taken;
  if (currBal) {
    currBal.textContent = remaining;
    currBal.classList.toggle('low', remaining === 0);
  }
  updateSummary(taken, remaining);
}

function updateSummary(taken, remaining) {
  const summary = document.getElementById('summaryText');
  if (!summary) return;
  const selected = [...selectedDays].sort();
  if (!selected.length) {
    summary.value = `${t('noDays')}\n${t('balanceSummary')(taken, annualBalance, remaining)}`;
    return;
  }

  const lines = [
    t('selectedSummary')(selected.length),
    ...selected.map(dateKey => formatDate(dateKey)),
    '',
    t('balanceSummary')(taken, annualBalance, remaining)
  ];
  summary.value = lines.join('\n');
}

function loadState() {
  try {
    const state = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
    annualBalance = Number.isFinite(state.annualBalance) ? state.annualBalance : DEFAULT_BALANCE;
    selectedDays = new Set(Array.isArray(state.selectedDays) ? state.selectedDays : []);
  } catch {
    annualBalance = DEFAULT_BALANCE;
    selectedDays = new Set();
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    annualBalance,
    selectedDays: [...selectedDays]
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
