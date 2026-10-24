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

  const savedLanguage = store.get(LANGUAGE_KEY) || store.get('holidayPlanner:language') || 'ar';
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
      const lang = store.get(LANGUAGE_KEY) || 'ar';
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
      return store.get(LANGUAGE_KEY) || 'ar';
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

/* ── Fun plane game (ES5, no dependencies, works on old browsers) ── */
(function () {
  var SVG_PLANE = '<svg viewBox="0 0 64 64" width="100%" height="100%" aria-hidden="true">' +
    '<path fill="currentColor" d="M62 30c0-2-3-4-6-4H41L26 6c-.4-.6-1-1-1.8-1h-3.6c-1 0-1.600.9-1.300 1.800L25 26H12l-5-7c-.4-.6-1-1-1.800-1H3c-1 0-1.500.9-1.200 1.800L5 30l-3.200 10.200C1.500 41.100 2 42 3 42h2.200c.7 0 1.400-.4 1.800-1l5-7h13l-4.700 19.200c-.3.9.3 1.800 1.300 1.800h3.600c.7 0 1.400-.4 1.800-1l15-20h15c3 0 6-2 6-4z"/></svg>';

  var raf = window.requestAnimationFrame || window.webkitRequestAnimationFrame ||
    window.mozRequestAnimationFrame || function (cb) { return window.setTimeout(function () { cb(new Date().getTime()); }, 16); };

  function now() { return new Date().getTime(); }
  function rand(min, max) { return min + Math.random() * (max - min); }

  var sky = null;

  function getSky() {
    if (!sky) {
      sky = document.createElement('div');
      sky.className = 'plane-sky';
      document.body.appendChild(sky);
    }
    return sky;
  }

  function flyPlane(delay) {
    var layer = getSky();
    var vw = document.documentElement.clientWidth || window.innerWidth || 1000;
    var vh = document.documentElement.clientHeight || window.innerHeight || 700;

    var size = Math.round(rand(40, 90));
    var baseTop = rand(30, Math.max(60, vh - size - 40));
    var wobble = rand(8, 28);
    var duration = rand(3200, 5200);
    var startX = -(size + 100);
    var endX = vw + 40;

    var el = document.createElement('div');
    el.className = 'plane-fly';
    el.style.width = size + 'px';
    el.style.height = size + 'px';
    el.style.left = startX + 'px';
    el.style.top = baseTop + 'px';
    el.style.display = 'none';
    el.innerHTML = '<span class="plane-trail"></span>' + SVG_PLANE;
    layer.appendChild(el);

    var startTime = now() + delay;

    function step() {
      var t = now() - startTime;
      if (t < 0) { raf(step); return; }
      if (t === 0 || el.style.display === 'none') { el.style.display = 'block'; }
      var p = t / duration;
      if (p >= 1) {
        if (el.parentNode) { el.parentNode.removeChild(el); }
        return;
      }
      el.style.left = Math.round(startX + (endX - startX) * p) + 'px';
      el.style.top = Math.round(baseTop + Math.sin(p * Math.PI * 3) * wobble) + 'px';
      raf(step);
    }
    raf(step);
  }

  function launchPlanes() {
    var count = 6;
    for (var i = 0; i < count; i++) {
      flyPlane(i * 280 + rand(0, 200));
    }
  }

  function init() {
    var btn = document.getElementById('planeBtn');
    if (!btn) { return; }
    var handler = function (e) {
      if (e && e.preventDefault) { e.preventDefault(); }
      launchPlanes();
    };
    if (btn.addEventListener) { btn.addEventListener('click', handler, false); }
    else if (btn.attachEvent) { btn.attachEvent('onclick', handler); }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
}());
