(function () {
  // localStorage can throw (private mode, blocked cookies, old browsers): never let that stop the page.
  const store = {
    get(key) {
      try { return window.localStorage.getItem(key); } catch (error) { return null; }
    },
    set(key, value) {
      try { window.localStorage.setItem(key, value); } catch (error) { /* ignore */ }
    }
  };

  // CustomEvent constructor fallback for older browsers.
  function emit(name, detail) {
    let event;
    if (typeof window.CustomEvent === 'function') {
      event = new CustomEvent(name, { detail: detail });
    } else {
      event = document.createEvent('CustomEvent');
      event.initCustomEvent(name, false, false, detail);
    }
    window.dispatchEvent(event);
  }

  const LANGUAGE_KEY = 'egyptairPortal:language';
  const THEME_KEY = 'egyptairPortal:theme';

  const savedLanguage = store.get(LANGUAGE_KEY) || store.get('holidayPlanner:language') || 'en';
  const savedTheme = store.get(THEME_KEY) || 'light';

  function setLanguage(language) {
    const lang = language === 'ar' ? 'ar' : 'en';
    store.set(LANGUAGE_KEY, lang);
    store.set('holidayPlanner:language', lang);
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

    document.querySelectorAll('[data-en-placeholder][data-ar-placeholder]').forEach((element) => {
      element.placeholder = element.getAttribute(`data-${lang}-placeholder`) || '';
    });

    document.querySelectorAll('[data-lang-option]').forEach((button) => {
      button.classList.toggle('active', button.dataset.langOption === lang);
    });

    emit('portal-language-change', { language: lang });
  }

  function setTheme(theme) {
    const nextTheme = theme === 'dark' ? 'dark' : 'light';
    store.set(THEME_KEY, nextTheme);
    document.documentElement.dataset.theme = nextTheme;

    document.querySelectorAll('[data-theme-toggle]').forEach((button) => {
      const lang = store.get(LANGUAGE_KEY) || 'en';
      const lightLabel = lang === 'ar' ? 'فاتح' : 'Light';
      const darkLabel = lang === 'ar' ? 'داكن' : 'Dark';
      button.textContent = nextTheme === 'dark' ? lightLabel : darkLabel;
    });
  }

  document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('[data-lang-option]').forEach((button) => {
      button.addEventListener('click', () => {
        setLanguage(button.dataset.langOption);
        setTheme(store.get(THEME_KEY) || 'light');
      });
    });

    document.querySelectorAll('[data-theme-toggle]').forEach((button) => {
      button.addEventListener('click', () => {
        setTheme((store.get(THEME_KEY) || 'light') === 'dark' ? 'light' : 'dark');
      });
    });

    // setLanguage first, then setTheme: the theme button label depends on the language,
    // and setLanguage() rewrites every [data-en][data-ar] label (including the theme button).
    setLanguage(savedLanguage);
    setTheme(savedTheme);
  });

  window.EgyptAirPortal = {
    get language() {
      return store.get(LANGUAGE_KEY) || 'en';
    },
    setLanguage,
    setTheme
  };

  /* ── Toast helper ── */
  function ensureToastContainer() {
    let container = document.getElementById('portal-toast-container');
    if (!container) {
      container = document.createElement('div');
      container.id = 'portal-toast-container';
      container.className = 'toast-container';
      document.body.appendChild(container);
    }
    return container;
  }

  window.showToast = function(message, type = 'info') {
    const container = ensureToastContainer();
    const toast = document.createElement('div');
    toast.className = `toast toast--${type}`;
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 4000);
  };
}());
