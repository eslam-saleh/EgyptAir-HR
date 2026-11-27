/*
 * secure.js – keeps links and API keys out of readable page source.
 *
 * Load it FIRST on every page (before misc-config.js / auth.js):
 *   <script src="secure.js"></script>
 *
 * The sealed blobs live in each page / config file, the SALT lives only here:
 * neither one alone is enough to read a value.
 *
 * ── Hashing + sealing ───────────────────────────────────────────────────
 *   sha256Hex : pure JS (works on plain http:// where crypto.subtle does not)
 *               and UTF-8 safe, so Arabic text hashes correctly.
 *
 *   seal / unseal : SHA-256 used as a stream cipher (counter mode), so the
 *               values never appear as readable text in View Source or the
 *               Elements panel. This stops casual snooping; it is obfuscation,
 *               not a vault: the browser must be able to decode it, so a
 *               developer still can.
 *
 * Values are decoded in memory ONLY at the moment they are used. Nothing is
 * cached: call secure.unseal(blob) right where the value is needed.
 *
 * To seal a new value (Node, from this folder):
 *   node seal-tool.js "https://example.com/exec"
 */
(function (root) {
  'use strict';

  // Same salt style as the rest of the portal. Change it and re-seal
  // (every blob must be re-sealed after a change).
  var SALT = ['ea', 'clinics', '\u2708', 'v1'].join('\u00b7');

  /* ---------- UTF-8 (manual, so it works everywhere and for Arabic) ---------- */
  function utf8Encode(str) {
    var out = [];
    str = String(str);
    for (var i = 0; i < str.length; i++) {
      var c = str.charCodeAt(i);
      if (c >= 0xD800 && c <= 0xDBFF && i + 1 < str.length) {
        var d = str.charCodeAt(i + 1);
        if (d >= 0xDC00 && d <= 0xDFFF) { c = 0x10000 + ((c - 0xD800) << 10) + (d - 0xDC00); i++; }
      }
      if (c < 0x80) out.push(c);
      else if (c < 0x800) out.push(0xC0 | (c >> 6), 0x80 | (c & 63));
      else if (c < 0x10000) out.push(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
      else out.push(0xF0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
    }
    return out;
  }

  function utf8Decode(bytes) {
    var s = '';
    for (var i = 0; i < bytes.length;) {
      var b = bytes[i++], c;
      if (b < 0x80) c = b;
      else if (b < 0xE0) c = ((b & 31) << 6) | (bytes[i++] & 63);
      else if (b < 0xF0) { c = ((b & 15) << 12) | ((bytes[i++] & 63) << 6); c |= bytes[i++] & 63; }
      else {
        c = ((b & 7) << 18) | ((bytes[i++] & 63) << 12); c |= (bytes[i++] & 63) << 6; c |= bytes[i++] & 63;
        c -= 0x10000; s += String.fromCharCode(0xD800 + (c >> 10), 0xDC00 + (c & 1023)); continue;
      }
      s += String.fromCharCode(c);
    }
    return s;
  }

  /* ---------- SHA-256 (pure JS) ---------- */
  var K = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
  ];

  function rotr(x, n) { return (x >>> n) | (x << (32 - n)); }

  function sha256Bytes(bytes) {
    var l = bytes.length, bitHi = Math.floor(l / 0x20000000), bitLo = (l << 3) >>> 0;
    var m = bytes.slice();
    m.push(0x80);
    while ((m.length % 64) !== 56) m.push(0);
    m.push((bitHi >>> 24) & 255, (bitHi >>> 16) & 255, (bitHi >>> 8) & 255, bitHi & 255,
           (bitLo >>> 24) & 255, (bitLo >>> 16) & 255, (bitLo >>> 8) & 255, bitLo & 255);
    var H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
    var w = new Array(64);
    for (var off = 0; off < m.length; off += 64) {
      var t;
      for (t = 0; t < 16; t++) {
        w[t] = ((m[off + 4 * t] << 24) | (m[off + 4 * t + 1] << 16) | (m[off + 4 * t + 2] << 8) | m[off + 4 * t + 3]) | 0;
      }
      for (t = 16; t < 64; t++) {
        var s0 = rotr(w[t - 15], 7) ^ rotr(w[t - 15], 18) ^ (w[t - 15] >>> 3);
        var s1 = rotr(w[t - 2], 17) ^ rotr(w[t - 2], 19) ^ (w[t - 2] >>> 10);
        w[t] = (w[t - 16] + s0 + w[t - 7] + s1) | 0;
      }
      var a = H[0], b = H[1], c = H[2], d = H[3], e = H[4], f = H[5], g = H[6], h = H[7];
      for (t = 0; t < 64; t++) {
        var S1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25);
        var ch = (e & f) ^ (~e & g);
        var t1 = (h + S1 + ch + K[t] + w[t]) | 0;
        var S0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22);
        var maj = (a & b) ^ (a & c) ^ (b & c);
        var t2 = (S0 + maj) | 0;
        h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
      }
      H[0] = (H[0] + a) | 0; H[1] = (H[1] + b) | 0; H[2] = (H[2] + c) | 0; H[3] = (H[3] + d) | 0;
      H[4] = (H[4] + e) | 0; H[5] = (H[5] + f) | 0; H[6] = (H[6] + g) | 0; H[7] = (H[7] + h) | 0;
    }
    var out = [];
    for (var i = 0; i < 8; i++) out.push((H[i] >>> 24) & 255, (H[i] >>> 16) & 255, (H[i] >>> 8) & 255, H[i] & 255);
    return out;
  }

  function toHex(bytes) {
    var s = '';
    for (var i = 0; i < bytes.length; i++) s += (bytes[i] < 16 ? '0' : '') + bytes[i].toString(16);
    return s;
  }

  function sha256Hex(text) { return toHex(sha256Bytes(utf8Encode(text))); }

  /* ---------- base64url (no padding) ---------- */
  var B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
  function b64Encode(bytes) {
    var s = '';
    for (var i = 0; i < bytes.length; i += 3) {
      var n = (bytes[i] << 16) | ((bytes[i + 1] || 0) << 8) | (bytes[i + 2] || 0);
      s += B64[(n >> 18) & 63] + B64[(n >> 12) & 63];
      if (i + 1 < bytes.length) s += B64[(n >> 6) & 63];
      if (i + 2 < bytes.length) s += B64[n & 63];
    }
    return s;
  }
  function b64Decode(str) {
    var out = [], buf = 0, bits = 0;
    for (var i = 0; i < str.length; i++) {
      var v = B64.indexOf(str.charAt(i));
      if (v < 0) return null;
      buf = (buf << 6) | v; bits += 6;
      if (bits >= 8) { bits -= 8; out.push((buf >> bits) & 255); buf &= (1 << bits) - 1; }
    }
    return out;
  }

  /* ---------- stream cipher: SHA-256(SALT | nonce | counter) as keystream ---------- */
  function keystream(nonce, length) {
    var out = [];
    for (var counter = 0; out.length < length; counter++) {
      var block = sha256Bytes(utf8Encode(SALT + '|' + nonce + '|' + counter));
      for (var i = 0; i < block.length && out.length < length; i++) out.push(block[i]);
    }
    return out;
  }

  function randomNonce() {
    var bytes = [], i;
    try {
      var c = root.crypto || root.msCrypto;
      if (c && c.getRandomValues) { var u = new Uint8Array(8); c.getRandomValues(u); for (i = 0; i < 8; i++) bytes.push(u[i]); }
    } catch (e) { bytes = []; }
    if (bytes.length !== 8) { bytes = []; for (i = 0; i < 8; i++) bytes.push(Math.floor(Math.random() * 256)); }
    return toHex(bytes);
  }

  function tagOf(nonce, plain) { return sha256Hex(SALT + '|tag|' + nonce + '|' + plain).slice(0, 8); }

  // seal(text) -> "s1.<nonce>.<tag>.<cipher>"
  function seal(plain) {
    plain = String(plain);
    var nonce = randomNonce();
    var data = utf8Encode(plain), ks = keystream(nonce, data.length), enc = [];
    for (var i = 0; i < data.length; i++) enc.push(data[i] ^ ks[i]);
    return 's1.' + nonce + '.' + tagOf(nonce, plain) + '.' + b64Encode(enc);
  }

  // unseal(blob) -> original text, or '' when the blob is empty, damaged or sealed with another salt.
  function unseal(blob) {
    if (typeof blob !== 'string' || !blob) return '';
    var p = blob.split('.');
    if (p.length !== 4 || p[0] !== 's1') return '';
    var enc = b64Decode(p[3]);
    if (!enc) return '';
    var ks = keystream(p[1], enc.length), data = [];
    for (var i = 0; i < enc.length; i++) data.push(enc[i] ^ ks[i]);
    var plain = utf8Decode(data);
    return tagOf(p[1], plain) === p[2] ? plain : '';
  }

  // unsealJSON(blob) -> parsed value (for sealed objects/arrays such as the EmailJS accounts), or null.
  function unsealJSON(blob) {
    var s = unseal(blob);
    if (!s) return null;
    try { return JSON.parse(s); } catch (e) { return null; }
  }

  /* ---------- opening a sealed link ---------- */
  // Builds a throw-away <a>, clicks it and removes it, so the address exists in the DOM only for that click.
  function open(blob, opts) {
    opts = opts || {};
    var url = unseal(blob);
    if (!url || !root.document) return false;
    var a = root.document.createElement('a');
    a.href = url;
    a.target = opts.target || '_blank';
    a.rel = 'noopener noreferrer';
    if (opts.download) a.setAttribute('download', opts.download);
    a.style.display = 'none';
    root.document.body.appendChild(a);
    a.click();
    root.document.body.removeChild(a);
    return true;
  }

  // Any <a data-secure-href="BLOB"> is decoded only when it is clicked (also middle-click / keyboard).
  function wireLinks() {
    var doc = root.document;
    if (!doc || !doc.addEventListener) return;
    function handler(ev) {
      var el = ev.target && ev.target.closest ? ev.target.closest('a[data-secure-href]') : null;
      if (!el) return;
      if (ev.type === 'auxclick' && ev.button !== 1) return;
      ev.preventDefault();
      open(el.getAttribute('data-secure-href'), {
        target: el.getAttribute('target') || '_blank',
        download: el.getAttribute('download') || ''
      });
    }
    doc.addEventListener('click', handler);
    doc.addEventListener('auxclick', handler);
  }

  var secure = {
    sha256Hex: sha256Hex,
    seal: seal,
    unseal: unseal,
    unsealJSON: unsealJSON,
    open: open
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = secure;
  else { root.secure = secure; wireLinks(); }
}(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this)));
