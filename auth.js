/*
 * auth.js – one shared sign-in for the whole portal.
 *
 * Load it FIRST in <head> of every page:
 *   <script src="auth.js"></script>                      (any valid code)
 *   <script src="auth.js" data-page="clinics"></script>  (needs the clinics right)
 *
 * Codes and their rights now live in the "auth" sheet of the متنوعات spreadsheet (see misc.gs).
 * The browser never sees the list: it sends the typed code to misc.gs, which answers with the rights.
 *
 * Rights (letters in the sheet's rights column):
 *   c = Clinics     t = Countdown     v = Employee Data (view)     e = Employee Data (edit)
 *   m = Meetings
 *   s = Document Center "Select" button (fill-in-the-form for unpinned documents)
 * A code that is listed in the sheet can open every normal page; the letters add the extra pages.
 *
 * To add/remove a person: add/delete a row in the "auth" sheet with the SHA-256 of the code. Generate it with
 *   echo -n 12345 | sha256sum          (Linux / Mac / Git Bash)
 * NOTE: this is still a client-side gate. It keeps casual visitors out, it is not real security.
 */
(function () {
  'use strict';

  var SESSION_KEY = 'egyptairPortal:auth';
  var FAIL_KEY = 'egyptairPortal:authFails';
  var SESSION_HOURS = 8;       // login lapses after this long
  var MAX_FAILS = 5;           // wrong tries before a short wait
  var WAIT_SECONDS = 60;

  // Server: misc.gs (web app URL + optional token are set once in misc-config.js).
  var CFG = window.MISC_CONFIG || {};
  var LOGIN_TIMEOUT_MS = 20000;   // Apps Script can take a few seconds to wake up

  var PAGE_RIGHT = { clinics: 'c', tagging: 't', employees: 'v', meetings: 'm' };

  /* ---------- tiny safe storage (never throws) ---------- */
  var mem = {};
  function sget(k) { try { return window.localStorage.getItem(k); } catch (e) { return mem[k] || null; } }
  function sset(k, v) { try { window.localStorage.setItem(k, v); } catch (e) { mem[k] = v; } }
  function sdel(k) { try { window.localStorage.removeItem(k); } catch (e) { delete mem[k]; } }

  /* ---------- helpers ---------- */
  function normalizeDigits(s) {
    return String(s || '')
      .replace(/[\u0660-\u0669]/g, function (d) { return String(d.charCodeAt(0) - 0x0660); })
      .replace(/[\u06F0-\u06F9]/g, function (d) { return String(d.charCodeAt(0) - 0x06F0); })
      .replace(/\s+/g, '');
  }

  // Asks misc.gs to check the code. It answers { ok, rights } and, on success, writes the sign-in
  // to the "visitors" sheet. Rejects (network / blocked) so the login page can say so.
  function checkCode(code) {
    return new Promise(function (resolve, reject) {
      if (!CFG.url) { reject(new Error('MISC_CONFIG.url is empty')); return; }
      var ctl = window.AbortController ? new AbortController() : null;
      var timer = setTimeout(function () { if (ctl) ctl.abort(); reject(new Error('timeout')); }, LOGIN_TIMEOUT_MS);
      fetch(CFG.url, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'login', code: code, token: CFG.token || undefined }),
        signal: ctl ? ctl.signal : undefined
      }).then(function (r) { return r.json(); }).then(function (j) {
        clearTimeout(timer); resolve(j || {});
      }, function (e) { clearTimeout(timer); reject(e); });
    });
  }

  function failState() {
    try { return JSON.parse(sget(FAIL_KEY)) || { n: 0, until: 0 }; } catch (e) { return { n: 0, until: 0 }; }
  }

  function readSession() {
    try {
      var s = JSON.parse(sget(SESSION_KEY));
      if (s && typeof s.r === 'string' && Date.now() - s.at < SESSION_HOURS * 3600000) return s;
    } catch (e) { /* fall through */ }
    sdel(SESSION_KEY);
    return null;
  }

  var session = readSession();
  var rights = session ? session.r : '';

  function can(letter) { return rights.indexOf(letter) !== -1; }

  /* ---------- public API ---------- */
  window.PortalAuth = {
    signedIn: function () { return !!session; },
    can: can,                                   // PortalAuth.can('e')
    code: function () { return session ? session.c : ''; },
    logout: function () { sdel(SESSION_KEY); window.location.href = 'login.html'; },
    login: function (rawCode) {
      var fs = failState();
      var left = Math.ceil((fs.until - Date.now()) / 1000);
      if (left > 0) return Promise.resolve({ ok: false, wait: left });
      var code = normalizeDigits(rawCode);
      if (!code) return Promise.resolve({ ok: false, wait: 0 });
      return checkCode(code).then(function (res) {
        if (!res.ok) {
          if (res.error) throw new Error(res.error);   // server problem, not a wrong code
          fs.n = (fs.n || 0) + 1;
          if (fs.n >= MAX_FAILS) fs = { n: 0, until: Date.now() + WAIT_SECONDS * 1000 };
          sset(FAIL_KEY, JSON.stringify(fs));
          return { ok: false, wait: fs.until > Date.now() ? WAIT_SECONDS : 0 };
        }
        sset(SESSION_KEY, JSON.stringify({ c: code, r: String(res.rights || ''), at: Date.now() }));
        sdel(FAIL_KEY);
        return { ok: true, wait: 0 };
      });
    }
  };

  /* ---------- which page is this? ---------- */
  var me = document.currentScript;
  var pageName = me && me.getAttribute('data-page');
  if (pageName === 'login') return;            // login.html is the sign-in page itself
  var needed = pageName ? PAGE_RIGHT[pageName] : '';
  var allowed = !!session && (!needed || can(needed));
  window.PortalAuth.allowed = allowed;

  /* ---------- allowed: hide the links this person may not use ---------- */
  if (allowed) {
    var hide = '';
    ['c', 't', 'v', 'm'].forEach(function (r) {
      if (!can(r)) hide += '[data-requires~="' + r + '"]{display:none !important}';
    });
    if (hide) {
      var st = document.createElement('style');
      st.appendChild(document.createTextNode(hide));
      document.head.appendChild(st);
    }
    return;
  }

  /* ---------- NOT allowed: send the visitor to login.html and keep this page inert ---------- */
  function blockedCall() { return Promise.reject(new Error('Sign-in required')); }
  try { window.fetch = blockedCall; } catch (e) { /* ignore */ }
  try {
    if (window.XMLHttpRequest) {
      window.XMLHttpRequest.prototype.send = function () { throw new Error('Sign-in required'); };
    }
  } catch (e) { /* ignore */ }

  var hideStyle = document.createElement('style');
  hideStyle.appendChild(document.createTextNode('html{visibility:hidden !important}'));
  document.head.appendChild(hideStyle);

  var file = window.location.pathname.split('/').pop() || 'index.html';
  window.location.replace('login.html?next=' + encodeURIComponent(file) + (session ? '&denied=1' : ''));
}());
