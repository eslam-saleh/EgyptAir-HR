(function () {
  const LANGUAGE_KEY = 'egyptairPortal:language';
  const THEME_KEY = 'egyptairPortal:theme';

  const savedLanguage = localStorage.getItem(LANGUAGE_KEY) || localStorage.getItem('holidayPlanner:language') || 'en';
  const savedTheme = localStorage.getItem(THEME_KEY) || 'light';

  function setLanguage(language) {
    const lang = language === 'ar' ? 'ar' : 'en';
    localStorage.setItem(LANGUAGE_KEY, lang);
    localStorage.setItem('holidayPlanner:language', lang);
    document.documentElement.lang = lang;
    document.documentElement.dir = lang === 'ar' ? 'rtl' : 'ltr';

    document.querySelectorAll('[data-en][data-ar]').forEach((element) => {
      const value = element.getAttribute(`data-${lang}`);
      if (element.matches('input, textarea')) {
        element.placeholder = value;
      } else {
        element.textContent = value;
      }
    });

    document.querySelectorAll('[data-lang-option]').forEach((button) => {
      button.classList.toggle('active', button.dataset.langOption === lang);
    });

    window.dispatchEvent(new CustomEvent('portal-language-change', { detail: { language: lang } }));
  }

  function setTheme(theme) {
    const nextTheme = theme === 'dark' ? 'dark' : 'light';
    localStorage.setItem(THEME_KEY, nextTheme);
    document.documentElement.dataset.theme = nextTheme;

    document.querySelectorAll('[data-theme-toggle]').forEach((button) => {
      const lang = localStorage.getItem(LANGUAGE_KEY) || 'en';
      const lightLabel = lang === 'ar' ? 'فاتح' : 'Light';
      const darkLabel = lang === 'ar' ? 'داكن' : 'Dark';
      button.textContent = nextTheme === 'dark' ? lightLabel : darkLabel;
    });
  }

  document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('[data-lang-option]').forEach((button) => {
      button.addEventListener('click', () => {
        setLanguage(button.dataset.langOption);
        setTheme(localStorage.getItem(THEME_KEY) || 'light');
      });
    });

    document.querySelectorAll('[data-theme-toggle]').forEach((button) => {
      button.addEventListener('click', () => {
        setTheme((localStorage.getItem(THEME_KEY) || 'light') === 'dark' ? 'light' : 'dark');
      });
    });

    setTheme(savedTheme);
    setLanguage(savedLanguage);
  });

  window.EgyptAirPortal = {
    get language() {
      return localStorage.getItem(LANGUAGE_KEY) || 'en';
    },
    setLanguage,
    setTheme
  };
}());
