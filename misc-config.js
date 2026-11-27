/*
 * misc-config.js – where the متنوعات web app lives. Load it AFTER secure.js and BEFORE auth.js on every page:
 *   <script src="secure.js"></script>
 *   <script src="misc-config.js"></script>
 *   <script src="auth.js"></script>
 *
 * The web app URL is stored SEALED (see secure.js). It is decoded in memory only when a script reads
 * MISC_CONFIG.url, i.e. at the moment of the request, and never kept anywhere.
 *
 * To change it, seal the new "Web app" URL you get after deploying misc.gs:
 *   node seal-tool.js "https://script.google.com/macros/s/.../exec"
 * and paste the result into SEALED_URL below. If you set the MISC_TOKEN script property, seal that value the
 * same way and paste it into SEALED_TOKEN (leave it '' when there is no token).
 */
(function () {
  var SEALED_URL = 's1.dece62f5224f89e7.3eb459e8.yc-mqW53F0m94SC5LCdLp96D3HUkErja487O_IpJh1mfu5Tucrp8t1Sk089KNtY0Sx_dFolw84QxUFe5J3fxPTSAQYXjrPSYpdBjhNhHIwRKTj0pOlhXWa4nom2Xwf5MJ7KHYP2v8EyfT8wk2E73Cg';
  var SEALED_TOKEN = '';
  window.MISC_CONFIG = {
    get url() { return window.secure.unseal(SEALED_URL); },
    get token() { return window.secure.unseal(SEALED_TOKEN); }
  };
}());
