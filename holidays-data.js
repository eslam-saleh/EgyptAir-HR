/*
 * holidays-data.js – official holidays now live in the "holidays" sheet of the متنوعات spreadsheet
 * (columns: date | key | name_ar | name_en, date as yyyy-mm-dd). Shared by calendar-calculator.html
 * and holiday-planner.html. Load it AFTER misc-config.js.
 *
 *   HolidaysData.byYear   { '2026': [ ['2026-01-07', 'copticChristmas'], ... ] }
 *   HolidaysData.sets     { '2026': Set(['2026-01-07', ...]) }
 *   HolidaysData.names    { copticChristmas: { ar: '…', en: '…' } }
 *   HolidaysData.ready    Promise (never rejects) – resolves once the refresh finished or failed
 *   HolidaysData.onUpdate(fn)  – fn() runs when fresher data than the saved copy arrives
 *
 * The objects are filled from the copy saved in this browser straight away, then refreshed from
 * misc.gs (at most once every 5 minutes) and updated in place. A year with no rows is simply absent,
 * and holiday-planner.html then falls back to its fixed national holidays.
 */
(function () {
  'use strict';
  var CFG = window.MISC_CONFIG || {};
  var KEY = 'egyptairPortal:holidays';
  var TTL_MS = 5 * 60 * 1000;
  var TIMEOUT_MS = 15000;
  var listeners = [];

  var H = window.HolidaysData = {
    byYear: {}, sets: {}, names: {},
    onUpdate: function (fn) { if (typeof fn === 'function') listeners.push(fn); }
  };

  function apply(list) {
    [H.byYear, H.sets, H.names].forEach(function (o) { Object.keys(o).forEach(function (k) { delete o[k]; }); });
    (list || []).forEach(function (r) {
      if (!r || !/^\d{4}-\d{2}-\d{2}$/.test(r.date)) return;
      var y = r.date.slice(0, 4);
      (H.byYear[y] = H.byYear[y] || []).push([r.date, r.key || '']);
      (H.sets[y] = H.sets[y] || new Set()).add(r.date);
      if (r.key) H.names[r.key] = { ar: r.ar || '', en: r.en || '' };
    });
  }

  var saved = null;
  try { saved = JSON.parse(window.localStorage.getItem(KEY)); } catch (e) { /* ignore */ }
  if (saved && saved.list) apply(saved.list);

  var fresh = saved && saved.at && (Date.now() - saved.at) < TTL_MS;
  H.ready = new Promise(function (resolve) {
    if (fresh || !CFG.url || typeof fetch !== 'function') { resolve(H); return; }
    var timer = setTimeout(function () { resolve(H); }, TIMEOUT_MS);
    var url = CFG.url + (CFG.url.indexOf('?') === -1 ? '?' : '&') + 'action=holidays' +
      (CFG.token ? '&token=' + encodeURIComponent(CFG.token) : '');
    fetch(url).then(function (r) { return r.json(); }).then(function (j) {
      if (j && j.ok && Array.isArray(j.holidays)) {
        var changed = JSON.stringify(j.holidays) !== JSON.stringify(saved && saved.list);
        apply(j.holidays);
        try { window.localStorage.setItem(KEY, JSON.stringify({ at: Date.now(), list: j.holidays })); } catch (e) { /* ignore */ }
        if (changed) listeners.forEach(function (fn) { try { fn(); } catch (e) { /* ignore */ } });
      }
    }).catch(function () { /* keep the saved copy */ }).then(function () { clearTimeout(timer); resolve(H); });
  });
}());
