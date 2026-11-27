/*
 * site-data.js – the values now live in the "site-data" sheet of the متنوعات spreadsheet
 * (columns: key | value). Add as many rows as you like — key names are matched to placeholder
 * text case-insensitively and ignoring spaces/dashes/underscores, so "currentYear" will match
 * {current year}, {Current-Year}, {CURRENT_YEAR}, etc.
 *
 * This file only loads them: window.SiteData is filled from the last copy kept in this browser
 * straight away, then refreshed from misc.gs (at most once every 5 minutes). Anything that needs
 * the fresh copy can wait for  window.SiteDataReady  (a Promise that never rejects).
 */
(function () {
  'use strict';
  var CFG = window.MISC_CONFIG || {};
  var KEY = 'egyptairPortal:siteData';
  var TTL_MS = 5 * 60 * 1000;
  var TIMEOUT_MS = 15000;

  function read() { try { return JSON.parse(window.localStorage.getItem(KEY)) || null; } catch (e) { return null; } }
  function write(o) { try { window.localStorage.setItem(KEY, JSON.stringify(o)); } catch (e) { /* ignore */ } }

  var saved = read();
  var data = {};
  if (saved && saved.data) Object.keys(saved.data).forEach(function (k) { data[k] = saved.data[k]; });
  window.SiteData = data;   // one object for the page's lifetime; it is updated in place

  var fresh = saved && saved.at && (Date.now() - saved.at) < TTL_MS;
  window.SiteDataReady = new Promise(function (resolve) {
    if (fresh || !CFG.url || typeof fetch !== 'function') { resolve(data); return; }
    var timer = setTimeout(function () { resolve(data); }, TIMEOUT_MS);
    var url = CFG.url + (CFG.url.indexOf('?') === -1 ? '?' : '&') + 'action=sitedata' +
      (CFG.token ? '&token=' + encodeURIComponent(CFG.token) : '');
    fetch(url).then(function (r) { return r.json(); }).then(function (j) {
      if (j && j.ok && j.data) {
        Object.keys(data).forEach(function (k) { delete data[k]; });
        Object.keys(j.data).forEach(function (k) { data[k] = j.data[k]; });
        write({ at: Date.now(), data: j.data });
        try { window.dispatchEvent(new Event('sitedata:ready')); } catch (e) { /* ignore */ }
      }
    }).catch(function () { /* keep the saved copy */ }).then(function () { clearTimeout(timer); resolve(data); });
  });
}());
