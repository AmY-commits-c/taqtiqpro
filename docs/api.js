/**
 * Taqti' Pro — penghubung frontend GitHub Pages ↔ Google Apps Script.
 *
 * LANGKAH WAJIB: tampal URL Web app GAS (berakhir dengan /exec) di bawah.
 * Fail ini menggantikan `google.script.run`, jadi kod App/Latihan/Sets/Akses
 * tidak perlu diubah langsung.
 */
const GAS_URL = 'TAMPAL_URL_WEB_APP_DI_SINI';

(function () {
  function hantar(fn, args, ok, fail) {
    const gagal = function (msg) { if (fail) fail(new Error(msg)); else console.error(msg); };
    if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(GAS_URL)) {
      gagal('URL pelayan belum ditetapkan dalam api.js.');
      return;
    }
    fetch(GAS_URL, {
      method: 'POST',
      // text/plain = "simple request" → tiada CORS preflight yang GAS tidak sokong
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ fn: fn, args: args }),
      redirect: 'follow'
    })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (j && j.ok) { if (ok) ok(j.data); }
        else gagal((j && j.error) || 'Ralat pelayan.');
      })
      .catch(function (e) { gagal('Tidak dapat menghubungi pelayan. Semak sambungan internet. (' + e.message + ')'); });
  }

  // Tiruan google.script.run: .withSuccessHandler(fn).withFailureHandler(fn).namaFungsi(...args)
  function runner(ok, fail) {
    return new Proxy({}, {
      get: function (_, name) {
        if (name === 'withSuccessHandler') return function (h) { return runner(h, fail); };
        if (name === 'withFailureHandler') return function (h) { return runner(ok, h); };
        if (typeof name !== 'string') return undefined;
        return function () { hantar(name, Array.prototype.slice.call(arguments), ok, fail); };
      }
    });
  }

  window.google = window.google || {};
  window.google.script = { run: runner(null, null) };
})();
