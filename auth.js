/*
 * auth.js – one shared sign-in for the whole portal.
 *
 * Load it FIRST in <head> of every page:
 *   <script src="auth.js"></script>                      (any valid code)
 *   <script src="auth.js" data-page="clinics"></script>  (needs the clinics right)
 *
 * Rights (letters in USERS below):
 *   c = Clinics     t = Countdown     v = Employee Data (view)     e = Employee Data (edit)
 * A code that is listed in USERS can open every normal page; the letters add the extra pages.
 *
 * To add/remove a person: put the SHA-256 of the code in USERS. Generate it with
 *   echo -n 12345 | sha256sum          (Linux / Mac / Git Bash)
 * NOTE: this is a client-side gate. It keeps casual visitors out, it is not real security.
 */
(function () {
  'use strict';

  var SESSION_KEY = 'egyptairPortal:auth';
  var FAIL_KEY = 'egyptairPortal:authFails';
  var SESSION_HOURS = 8;       // login lapses after this long
  var MAX_FAILS = 5;           // wrong tries before a short wait
  var WAIT_SECONDS = 60;

  var USERS = {
    '40dcbc1c876fc50a6f1fada833a91038ef62a21666820629ee391b3b5feb3086': '',
    '44d71d355722bb0c6d04c8f221030c14ff89148551cd4fc39ce73942365ce8aa': '',
    '28c35700be5bd86f80aa8ba82632926734c73c2f471ab2923a6287a56eaf223a': 't',
    '3f64a7f4dd53bb20cfbc148d34d8bf866c5f8728da9dce2066be0e3e408c612e': 'tv',
    '2a6269286b00572886518126ed10de661ce4ae7256598ac70bfa56d2da14d161': '',
    'b3c24a5e33aa0057eae207a469be9c909ea962c379ca21092aa691bb09faa11f': 'c',
    '578508930eda772818694b02e79c545916bbf06428f58c559c750c566259532a': '',
    '62378664e2fb341da0c26e963c7631339357bb7390a83f136cb76f2645af95a6': '',
    'de40ee7ec6e2296125fb4a775262bbecf2fdbcc740cb81266fe3f7195ef2ad28': '',
    'a6b4a8573f97200bbba71127b0db313fb21824ff8390932590817e12ac2de906': '',
    '519a44856bd3de3985730aa8fc3b2e89910d6859df109c3a5fe288193587dbb5': '',
    'a44ed711645ec0f6b68109b35d88e3644b8b59b3c4a96a491f06b0e8d95e42f2': '',
    '9b4d605eedadee411de809393139385c3a06073dd9a052151d671308f6cd5b73': '',
    '5385a24c075400f76f2c028fd341ef2647e17fe386f3546a455d8cdb3fd2450e': '',
    'a26ceb16dadb19e0e066c113d2b84eaae118b69eafa9b405d843d203927078c5': '',
    '91ebcf68e80ce90efd50bb91a26470823cb89f50ae96c2236d254cbe7f90e09f': 'tv',
    'b04e39aa824b7a2d31c939b0e924d8233f0778dd902a3c5c0613fdf270761c2a': '',
    '3ba999816de46a27d66e582d1ef84692fc7ced9ea3848490b9241176ed591a16': 't',
    '29a70687c36ffcfc53be32d7aa62ac7dbc1c5e3e0782f8daa40893e67ffec5a7': '',
    'dc6631d46f359cdce2418b4c3ee5ef56ac2c6e4782d0383c2fcc48979ceaa107': 'v',
    '8325e50c76becc029e8b4c420a23dd2eb21ef54e26a82c8f11605ae0637e1adf': 'v',
    '7cb5076073ffc517d32f725305235ac140edcfece64dfe808d3afb7708108ffa': '',
    '3962ab4167f6d0d2712605ae96cb4584e8548ecd2d1e92e9b9d5a0073a8002bb': 'tve',
    'b041879ed6c2172fe9945360bea1a4502aa462236aa45dd6df30e75e14e73238': '',
    '3ea82a5852dd322aea1f0dc8e26275d6ab047c0ad8e3abf5d917160d325e5fb5': 'tv',
    'e1296ec40d9705193ee08cf5e65d9bbcc2a62e1209568b1bbb1d780f3d9a016f': '',
    '140da1171aa5c3052eda5e1facaadc542d19ad2213f2c7abdabb81d16a1c5e21': '',
    '5cffa0319559271572b16c67459ba31d1600e54d3f834e06d58c8c3d8433c1f7': '',
    '4066f060269c78cfcac71c7ad7db3c506ea8f36d9d28a59747d08e15096db6fa': '',
    '1dde0d1afffbeb04755818ecac8c353b5cc98459bc28793e9728d600de1ffc40': '',
    '5499d97bf22e193e9f5e7f84ef00a7984805f8824a7516a89aa675957facd89a': 'tv',
    '3bfd18d5b0a3bf82bcee673777d2683a41a21f09dd7c9f033abf9c6559c5660c': 'c',
    '0b04d7953771c6cb396d9436b5143aee4482b36d2b4bc9b95c14dc31df778187': 'tv',
    'bc1295b20a169139f17f78e1edfd76963b7b4b5f4d879f07bb9453a32431a96a': '',
    'f2cc0df5d8a663f56ebc753b6b1058e65175d6954e9e691094a1c2a9b207fcd6': '',
    '27ca4a53b4217f9d49b534cdd7328647cb607b2f88293f651322baae2d46e5c7': '',
    'd073e341e454ce18f401d4dc19e7c72661eb45dd8bf1bc802cb8416ef71aa971': 'tv',
    'd0857331c68dcca55e8a03ce1e3b9db98e1fc1b12e82872497b473df1de342b1': 'tv',
    '32bbfa9456562c2c9da1474fc961019c3473ed1eb80f68cdc25c9c82dfad816c': '',
    'c0f2a3b76fc1e9bf848e82af6de70a384091b465aa0561105d47f1e3f71ca851': '',
    '4f3dc22203978f1ef0dfbf3344c1c10f9bd3af9f547d528467e6a70f2ca702de': '',
    'd9d04476ba7871797298e1034a1104c67a3c9546381f2c57019703aba732bab0': '',
    '707ce489cc327373d501548e92bdaa24b12058ee8bfa7923849fc327ca6f663d': '',
    '9e44f14ce7f9474ab38a58870797821b758c6539a37b0e6271490ca641f69a4b': 'tv',
    'c3c048b2d83efb39b9737ce09df3ef3c2e382f7d4e0c2228b3a10259c589e1df': '',
    'fb99c2a25567005e010d1f833279a441b22bda2abd697fd46cad1eac91c7ccff': '',
    '08ecfb97885e4f6c54d328f0320a8bf3536b2e9c4dcb9e460f1b1369a294b018': '',
    '35444bace3288370c5e40ec05f645737ac4aa9200a31056cde636c41865fc24d': '',
    '11afcd456cc2a4eafd8c5134bcd4710e565f92af086473d937ede871f44cbe39': '',
    'c0a33cafa0967d46ed22ab30a84d2565ffe7f012aa9c50767eba8276143d88a3': '',
    '03d7be2755145563ffb1edd37de0f1d667fbfd26ac9d58f9b3309b5e05084cad': '',
    'a22dac8b32dea97c06226a1c163d738027bef582d8d564f8af04130990416d1f': '',
    'b6b343fed92fa3ebb38aebdf06e58982e6617906c72cb840e912c8e5d59019c8': '',
    '918a15fdc6826ac0e34b14a8c897cb7b315881bc365871687dd23a41cb4f37e8': '',
    '29f7e33c2656e36b348f68f78950f7c87638e93e8e81dedf4bad98bd8c40c816': '',
    'b08f9aadb5f34a8231db3e08af2fde5028832263fb13d11e9c60ad06f8eab295': 'tv',
    '726e5d45caaabaa026b256b007ed46b4db3415e45a8ca28c753ac9341be0080a': '',
    '753d4392e8972570ff23f29a62470e0ebb8bf530b03a9798b6f0eedb3c00b1e0': '',
    '43b4d56eff360052eee0595236ebc9f4e66c3d4244016b1009dfd7bd3fc0a91e': '',
    '5c387ae5b889281d1fb831fc693dd945e1f54f1587b6310170678def2f1c3672': '',
    '60dc13afd9db46855db7cb4b46b2429c473ae61d93c867bb13914cc0ee137a78': '',
    '770dd25925e4bbc1e09d89d1412be7b2d41c535a9370d58895a960e0162f7927': '',
    '0c2022d32040fa888913fb6d6e7e38d714207bad63ae648250ac43236f0639d2': '',
    'af9c4738ec118556c938831c6727cf3501dd5b9b1740f643b4740b3f9f05011f': '',
    'b204eadc21e568cf6948e6cc7bfee0dab69caed46d18b624ef1fd28b90f54b1b': '',
    '2614839e24e309b17d33185f18b0ed23f7a54c91f762c018ebf55b9652bd2c1a': 'tv',
    'd2e3ed1682b47a449e0625660782a042ec442ff6e9934e6f7646630fa8563ec6': '',
    'd24492e6d2975efb4164ba2b6953cf30c839417c98e8838f136924f28a4fd219': '',
    '44d724a47a714cbbca08f325a6fc85fefde03ae77413a3af161f4ed3ceebee1e': '',
    'a8c855fbaeaf16e3f04d4352ff0e576a32a3dc4cbb07e310f6fdfb61f6c28cce': 'tve',
    '24bd79f4173351729adbc261a88702b88b7e56fef8b55d2c4a1f800f8363370b': '',
    'e85db97e0af29710a13e4fe619497617801de31f2898160761d7ef3e52c748aa': '',
    '492f431bae35265f2e5f4ed49bd8c58dda912431be561504846988d00d05d117': '',
    'a591c960cc7c01ab484aa964b338b9297e9d05f5ffe7330dce50967345b70f64': 'ctv'
  };

  var PAGE_RIGHT = { clinics: 'c', tagging: 't', employees: 'v' };

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

  function sha256Hex(text) {
    var bytes = new TextEncoder().encode(text);
    return crypto.subtle.digest('SHA-256', bytes).then(function (buf) {
      var out = '', arr = new Uint8Array(buf);
      for (var i = 0; i < arr.length; i++) out += (arr[i] < 16 ? '0' : '') + arr[i].toString(16);
      return out;
    });
  }

  function failState() {
    try { return JSON.parse(sget(FAIL_KEY)) || { n: 0, until: 0 }; } catch (e) { return { n: 0, until: 0 }; }
  }

  function readSession() {
    try {
      var s = JSON.parse(sget(SESSION_KEY));
      if (s && USERS.hasOwnProperty(s.h) && Date.now() - s.at < SESSION_HOURS * 3600000) return s;
    } catch (e) { /* fall through */ }
    sdel(SESSION_KEY);
    return null;
  }

  var session = readSession();
  var rights = session ? USERS[session.h] : '';

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
      return sha256Hex(code).then(function (h) {
        if (!USERS.hasOwnProperty(h)) {
          fs.n = (fs.n || 0) + 1;
          if (fs.n >= MAX_FAILS) fs = { n: 0, until: Date.now() + WAIT_SECONDS * 1000 };
          sset(FAIL_KEY, JSON.stringify(fs));
          return { ok: false, wait: fs.until > Date.now() ? WAIT_SECONDS : 0 };
        }
        sset(SESSION_KEY, JSON.stringify({ h: h, c: code, at: Date.now() }));
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
    ['c', 't', 'v'].forEach(function (r) {
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
