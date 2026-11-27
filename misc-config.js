/*
 * misc-config.js – where the متنوعات web app lives. Load it BEFORE auth.js on every page:
 *   <script src="misc-config.js"></script>
 *   <script src="auth.js"></script>
 *
 * Paste the "Web app" URL you get after deploying misc.gs (Deploy ▸ New deployment ▸ Web app,
 * Execute as: Me, Who has access: Anyone). If you set the MISC_TOKEN script property, put the same
 * value in token below; otherwise leave it empty.
 */
window.MISC_CONFIG = {
  url: 'https://script.google.com/macros/s/AKfycbypFSSQ6coD1PeyF5h6UMxJdnaNQWKDVQAOA0fQX4aHITf45Hnca2hJwZv9HvlAgu0n/exec',
  token: ''
};
