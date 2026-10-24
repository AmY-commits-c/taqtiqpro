/**
 * Taqti' Pro — Latihan & Semakan Syair 'Arudh & Qafiyah
 * Bahagian pelayan (Google Apps Script, terikat pada Google Sheet).
 *
 * Tab yang digunakan:
 *   Tetapan, Rujukan_BukuTeks, Bank_Contoh_Syair, Rekod_Pelajar, Set_Latihan, Jawapan_Latihan,
 *   Permohonan_Akses, Kod_Akses (akses berbayar — bayaran manual, kod dijana bila guru luluskan)
 * Semua analisis (padanan & taqti') berlaku dalam pelayar — pelayan hanya baca/tulis Sheet.
 */

const APP_VERSION = 'v2.9 (2026-10-24)'; // + API doPost untuk frontend GitHub Pages

const SH = {
  TETAPAN: 'Tetapan',
  RUJUKAN: 'Rujukan_BukuTeks',
  BANK: 'Bank_Contoh_Syair',
  REKOD: 'Rekod_Pelajar',
  SET: 'Set_Latihan',
  JAWAPAN: 'Jawapan_Latihan',
  MOHON: 'Permohonan_Akses',
  KOD: 'Kod_Akses'
};

const HEADERS = {
  Rujukan_BukuTeks: ['ID', 'Bahar', 'Laqob', 'Arudh', 'Dharb', 'Teks', 'Halaman', 'Catatan'],
  Bank_Contoh_Syair: ['ID', 'Bahar', 'Laqob', 'Arudh', 'Dharb', 'Teks', 'Bil_Bait', 'Penyair', 'Zaman', 'Sumber', 'Skor_Sistem', 'Status', 'Catatan'],
  Rekod_Pelajar: ['ID', 'Tahun', 'Nama_Pelajar', 'Kelas', 'Teks', 'Bahar', 'Laqob', 'Arudh', 'Dharb', 'Penyair_Sumber', 'Skor_BukuTeks', 'Padanan_BukuTeks', 'Amaran_Pernah_Guna', 'Tarikh', 'Guru', 'Catatan'],
  Set_Latihan: ['ID', 'Nama_Set', 'Tahun', 'Kelas_Sasaran', 'Status', 'Bantuan', 'Bil_Item', 'Item_JSON', 'Dicipta', 'Dikemaskini', 'Guru', 'Catatan', 'Susunan'],
  Jawapan_Latihan: ['ID', 'Set_ID', 'Tahun', 'Nama_Pelajar', 'Kelas', 'Item_No', 'Teks', 'Bahar_Sebenar', 'Laqob_Sebenar', 'Bahar_Jawapan', 'Laqob_Jawapan',
    'Bahar_OK', 'Laqob_OK', 'Taqti_Betul', 'Taqti_Jumlah', 'Taf_Betul', 'Taf_Jumlah', 'Hal_Betul', 'Hal_Jumlah', 'Skor', 'Butiran_JSON', 'Tarikh', 'Kod'],
  Permohonan_Akses: ['ID', 'Tarikh', 'Nama', 'Kelas_Sekolah', 'Telefon', 'Emel', 'Pakej_Bulan', 'Jumlah', 'Rujukan_Bayaran', 'Resit_URL', 'Status', 'Kod', 'Diproses', 'Catatan'],
  Kod_Akses: ['ID', 'Kod', 'Nama', 'Kelas_Sekolah', 'Telefon', 'Bulan', 'Mula', 'Tamat', 'Status', 'Peranti', 'Permohonan_ID', 'Dicipta', 'Catatan']
};

const DEFAULT_SETTINGS = [
  ['Kunci', 'Nilai', 'Keterangan'],
  ['Nama_Sekolah', '', 'Dipaparkan di bahagian atas aplikasi'],
  ['Ambang_Tolak', 70, 'Skor padanan buku teks (%) yang menyebabkan syair DITOLAK'],
  ['Ambang_Ragu', 40, 'Skor padanan (%) yang dianggap MERAGUKAN — guru perlu semak'],
  ['Min_Huruf', 15, 'Bilangan huruf minimum untuk semakan yang bermakna'],
  ['PIN_Guru', '', 'WAJIB untuk akses guru (Rekod, Bank, Buku teks, simpan & padam). Kosong = akses guru dikunci sepenuhnya.'],
  ['Mod_Akses', 'Percuma', 'Percuma = Latihan terbuka kepada semua. Berbayar = Latihan perlu kod akses (Semak syair kekal percuma).'],
  ['Pakej_Harga', '1:10, 3:25, 6:40', 'Pakej akses: bulan:harga RM, dipisah koma. Cth. 1:10, 3:25, 6:40'],
  ['Maklumat_Bayaran', 'Bank: ____ | No. akaun: ____ | Nama: ____', 'Arahan bayaran yang dipaparkan kepada pembeli'],
  ['QR_Bayaran_URL', '', 'Pilihan: pautan gambar QR DuitNow (fail Drive dikongsi "Sesiapa yang ada pautan")'],
  ['Had_Peranti', 2, 'Bilangan peranti maksimum bagi satu kod akses'],
  ['Emel_Notifikasi', '', 'Pilihan: emel guru untuk dimaklumkan bila ada permohonan baharu']
];

/* ───────────── Menu & persediaan ───────────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Taqti' Pro")
    .addItem('Sediakan / baiki tab', 'setupSheets')
    .addToUi();
}

function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  let t = ss.getSheetByName(SH.TETAPAN);
  if (!t) {
    t = ss.insertSheet(SH.TETAPAN);
    t.getRange(1, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    t.getRange(1, 1, 1, 3).setFontWeight('bold');
    t.setColumnWidth(3, 420);
  } else {
    const keys = t.getRange(1, 1, Math.max(1, t.getLastRow()), 1).getValues().map(function (r) { return String(r[0]); });
    DEFAULT_SETTINGS.slice(1).forEach(function (row) { if (keys.indexOf(row[0]) < 0) t.appendRow(row); });
  }
  Object.keys(HEADERS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const h = HEADERS[name];
    const first = sh.getRange(1, 1, 1, h.length).getValues()[0];
    if (first.join('') === '') {
      sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
      sh.setFrozenRows(1);
    } else {
      // tambah lajur baharu (cth. Susunan) di hujung tanpa mengusik data sedia ada
      h.forEach(function (name, k) { if (String(first[k] || '') === '') sh.getRange(1, k + 1).setValue(name).setFontWeight('bold'); });
    }
    const teksCol = h.indexOf('Teks') + 1;
    if (teksCol > 0) sh.setColumnWidth(teksCol, 520);
  });
  return 'Tab sudah sedia.';
}

/* ───────────── Web app ───────────── */

function doGet() {
  setupSheets();
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle("Taqti' Pro — Latihan 'Arudh & Qafiyah")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ───────────── API untuk frontend luar (GitHub Pages) ─────────────
 * Frontend menghantar POST (Content-Type: text/plain) dengan badan JSON:
 *   { "fn": "getBootData", "args": [ ... ] }
 * Jawapan: { ok: true, data: ... }  atau  { ok: false, error: "mesej" }
 * Hanya fungsi dalam senarai di bawah boleh dipanggil. Semakan PIN guru & kod akses
 * kekal di pelayan seperti biasa.
 */
const API_FNS = {
  getBootData: getBootData, getGuruData: getGuruData,
  saveRekod: saveRekod, saveBank: saveBank, saveRujukan: saveRujukan, updateField: updateField, deleteById: deleteById,
  saveSet: saveSet, saveSetOrder: saveSetOrder, submitJawapan: submitJawapan,
  mohonAkses: mohonAkses, semakPermohonan: semakPermohonan, semakKod: semakKod,
  luluskanPermohonan: luluskanPermohonan, tolakPermohonan: tolakPermohonan, janaKodManual: janaKodManual,
  urusKod: urusKod, setModAkses: setModAkses
};

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API_FNS[String(req.fn || '')];
    if (!fn) throw new Error('Fungsi tidak dikenali: ' + req.fn);
    if (req.fn === 'getBootData') setupSheets(); // sama seperti doGet
    const args = Array.isArray(req.args) ? req.args : [];
    const data = fn.apply(null, args);
    out = { ok: true, data: data === undefined ? null : data };
  } catch (err) {
    out = { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────────── Baca data ───────────── */

function readTable_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const head = values.shift().map(String);
  return values
    .filter(function (r) { return r.join('').trim() !== ''; })
    .map(function (r) {
      const o = {};
      head.forEach(function (h, i) {
        let v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        o[h] = v === null || v === undefined ? '' : v;
      });
      return o;
    });
}

function readSettings_() {
  const s = { Nama_Sekolah: '', Ambang_Tolak: 70, Ambang_Ragu: 40, Min_Huruf: 15 };
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.TETAPAN);
  if (!sh) return s;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[0] !== '' && r[1] !== '') s[String(r[0]).trim()] = r[1];
  });
  return s;
}

function pinOk_(pin) {
  const want = String(readSettings_().PIN_Guru || '').trim();
  return !!want && String(pin || '').trim() === want; // PIN kosong = tiada akses guru
}
function checkPin_(pin) { if (!pinOk_(pin)) throw new Error('Akses guru diperlukan (PIN tidak sah).'); }

/** Data awam untuk paparan pelajar: Semak & Latihan. Tiada nama pelajar dihantar. */
function getBootData() {
  const st = readSettings_();
  return {
    version: APP_VERSION,
    settings: { Nama_Sekolah: st.Nama_Sekolah, Ambang_Tolak: st.Ambang_Tolak, Ambang_Ragu: st.Ambang_Ragu, Min_Huruf: st.Min_Huruf },
    akses: aksesInfo_(st),
    pinSet: !!String(st.PIN_Guru || '').trim(),
    rujukan: readTable_(SH.RUJUKAN),
    bank: readTable_(SH.BANK).filter(function (r) { return r.Status !== 'Ditolak'; }),
    rekodLite: readTable_(SH.REKOD).map(function (r) { return { Tahun: r.Tahun, Teks: r.Teks }; }),
    sets: readTable_(SH.SET).filter(function (r) { return r.Status === 'Aktif'; }).map(publicSet_)
  };
}

/** Data penuh untuk guru — hanya dengan PIN yang betul. */
function getGuruData(pin) {
  checkPin_(pin);
  return {
    rekod: readTable_(SH.REKOD),
    bank: readTable_(SH.BANK),
    setsAll: readTable_(SH.SET).map(publicSet_),
    jawapan: readTable_(SH.JAWAPAN),
    mohon: readTable_(SH.MOHON),
    kod: readTable_(SH.KOD),
    guru: Session.getActiveUser().getEmail() || ''
  };
}

function publicSet_(r) {
  let items = [];
  try { items = JSON.parse(r.Item_JSON || '[]'); } catch (e) { items = []; }
  return { ID: r.ID, Nama_Set: r.Nama_Set, Tahun: r.Tahun, Kelas_Sasaran: r.Kelas_Sasaran, Status: r.Status,
    Bantuan: r.Bantuan, Bil_Item: r.Bil_Item, items: items, Dicipta: r.Dicipta, Dikemaskini: r.Dikemaskini, Catatan: r.Catatan, Susunan: r.Susunan };
}

/* ───────────── Tulis data ───────────── */

function nextId_(sh, prefix) {
  const last = sh.getLastRow();
  let max = 0;
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      const m = String(r[0]).match(/(\d+)$/);
      if (m) max = Math.max(max, Number(m[1]));
    });
  }
  return prefix + String(max + 1).padStart(3, '0');
}

function appendRow_(name, prefix, obj) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const h = HEADERS[name];
  obj.ID = nextId_(sh, prefix);
  sh.appendRow(h.map(function (k) { return obj[k] === undefined ? '' : obj[k]; }));
  return obj;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function clean_(s) { return String(s || '').replace(/[\t\r\n]+/g, ' ').trim(); }

/** Simpan syair pelajar. alsoBank = true → salin juga ke Bank_Contoh_Syair. */
function saveRekod(rec, alsoBank, pin) {
  checkPin_(pin);
  return withLock_(function () {
    const guru = Session.getActiveUser().getEmail() || '';
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const row = {
      Tahun: Number(rec.Tahun) || new Date().getFullYear(),
      Nama_Pelajar: clean_(rec.Nama_Pelajar), Kelas: clean_(rec.Kelas), Teks: clean_(rec.Teks),
      Bahar: clean_(rec.Bahar), Laqob: clean_(rec.Laqob), Arudh: clean_(rec.Arudh), Dharb: clean_(rec.Dharb),
      Penyair_Sumber: clean_(rec.Penyair_Sumber), Skor_BukuTeks: rec.Skor_BukuTeks,
      Padanan_BukuTeks: clean_(rec.Padanan_BukuTeks), Amaran_Pernah_Guna: clean_(rec.Amaran_Pernah_Guna),
      Tarikh: now, Guru: guru, Catatan: clean_(rec.Catatan)
    };
    if (!row.Teks || !row.Nama_Pelajar || !row.Bahar) throw new Error('Teks, nama pelajar dan bahar wajib diisi.');
    const saved = appendRow_(SH.REKOD, 'RP-', row);
    let bank = null;
    if (alsoBank) {
      bank = appendRow_(SH.BANK, 'BC-', {
        Bahar: row.Bahar, Laqob: row.Laqob, Arudh: row.Arudh, Dharb: row.Dharb, Teks: row.Teks,
        Bil_Bait: rec.Bil_Bait || '', Penyair: row.Penyair_Sumber, Zaman: '', Sumber: 'Kerja kursus ' + row.Tahun,
        Skor_Sistem: rec.Skor_Sistem || '', Status: 'Perlu Semak', Catatan: 'Dari ' + saved.ID
      });
    }
    return { rekod: saved, bank: bank };
  });
}

function saveBank(rec, pin) {
  checkPin_(pin);
  return withLock_(function () {
    if (!clean_(rec.Teks) || !clean_(rec.Bahar)) throw new Error('Teks dan bahar wajib diisi.');
    return appendRow_(SH.BANK, 'BC-', {
      Bahar: clean_(rec.Bahar), Laqob: clean_(rec.Laqob), Arudh: clean_(rec.Arudh), Dharb: clean_(rec.Dharb),
      Teks: clean_(rec.Teks), Bil_Bait: rec.Bil_Bait || '', Penyair: clean_(rec.Penyair), Zaman: clean_(rec.Zaman),
      Sumber: clean_(rec.Sumber), Skor_Sistem: rec.Skor_Sistem || '', Status: rec.Status || 'Perlu Semak',
      Catatan: clean_(rec.Catatan)
    });
  });
}

function saveRujukan(rec, pin) {
  checkPin_(pin);
  return withLock_(function () {
    if (!clean_(rec.Teks) || !clean_(rec.Bahar)) throw new Error('Teks dan bahar wajib diisi.');
    return appendRow_(SH.RUJUKAN, 'BT-', {
      Bahar: clean_(rec.Bahar), Laqob: clean_(rec.Laqob), Arudh: clean_(rec.Arudh), Dharb: clean_(rec.Dharb),
      Teks: clean_(rec.Teks), Halaman: clean_(rec.Halaman), Catatan: clean_(rec.Catatan)
    });
  });
}

/** Kemas kini satu medan pada baris yang dikenal pasti dengan ID. */
function updateField(sheetName, id, field, value, pin) {
  checkPin_(pin);
  if ([SH.BANK, SH.REKOD, SH.RUJUKAN, SH.SET, SH.JAWAPAN, SH.MOHON, SH.KOD].indexOf(sheetName) < 0) throw new Error('Tab tidak sah.');
  return withLock_(function () {
    const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
    const h = HEADERS[sheetName];
    const col = h.indexOf(field) + 1;
    if (col < 1) throw new Error('Medan tidak sah: ' + field);
    const ids = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sh.getRange(i + 2, col).setValue(clean_(value));
        return true;
      }
    }
    throw new Error('ID ' + id + ' tidak dijumpai.');
  });
}

function deleteById(sheetName, id, pin) {
  checkPin_(pin);
  if ([SH.BANK, SH.REKOD, SH.RUJUKAN, SH.SET, SH.JAWAPAN, SH.MOHON, SH.KOD].indexOf(sheetName) < 0) throw new Error('Tab tidak sah.');
  return withLock_(function () {
    const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
    const ids = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) { sh.deleteRow(i + 2); return true; }
    }
    throw new Error('ID ' + id + ' tidak dijumpai.');
  });
}


/* ───────────── Set latihan ───────────── */

/** Cipta (tiada ID) atau kemas kini (ada ID) set latihan. items: [{bait:[sadr, ajuz], src:'BT-005#1'}] */
function saveSet(set, pin) {
  checkPin_(pin);
  return withLock_(function () {
    const items = (set.items || []).filter(function (it) { return it && it.bait && clean_(it.bait[0]); })
      .map(function (it) { return { bait: [clean_(it.bait[0]), clean_(it.bait[1] || '')], src: clean_(it.src || '') }; });
    if (!clean_(set.Nama_Set)) throw new Error('Nama set wajib diisi.');
    if (!items.length) throw new Error('Tambah sekurang-kurangnya satu item.');
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const row = {
      Nama_Set: clean_(set.Nama_Set), Tahun: Number(set.Tahun) || new Date().getFullYear(), Kelas_Sasaran: clean_(set.Kelas_Sasaran),
      Status: set.Status === 'Arkib' ? 'Arkib' : 'Aktif', Bantuan: set.Bantuan === 'Tidak' ? 'Tidak' : 'Ya',
      Bil_Item: items.length, Item_JSON: JSON.stringify(items), Dikemaskini: now, Guru: Session.getActiveUser().getEmail() || '', Catatan: clean_(set.Catatan)
    };
    const sh = SpreadsheetApp.getActive().getSheetByName(SH.SET);
    const h = HEADERS[SH.SET];
    if (set.ID) {
      const ids = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(set.ID)) {
          const cur = sh.getRange(i + 2, 1, 1, h.length).getValues()[0];
          row.ID = set.ID; row.Dicipta = cur[h.indexOf('Dicipta')]; row.Susunan = cur[h.indexOf('Susunan')];
          sh.getRange(i + 2, 1, 1, h.length).setValues([h.map(function (k) { return row[k] === undefined ? '' : row[k]; })]);
          return publicSet_(row);
        }
      }
      throw new Error('Set ' + set.ID + ' tidak dijumpai.');
    }
    row.Dicipta = now;
    return publicSet_(appendRow_(SH.SET, 'SL-', row));
  });
}

/** Simpan susunan set (ids mengikut urutan paparan). Lajur Susunan = 1, 2, 3 … */
function saveSetOrder(ids, pin) {
  checkPin_(pin);
  return withLock_(function () {
    const sh = SpreadsheetApp.getActive().getSheetByName(SH.SET);
    const h = HEADERS[SH.SET], col = h.indexOf('Susunan') + 1;
    const n = Math.max(0, sh.getLastRow() - 1);
    if (!n) return true;
    const cur = sh.getRange(2, 1, n, 1).getValues();
    const pos = {}; (ids || []).forEach(function (id, i) { pos[String(id)] = i + 1; });
    const out = cur.map(function (r) { const p = pos[String(r[0])]; return [p ? p : '']; });
    sh.getRange(2, col, n, 1).setValues(out);
    return true;
  });
}

/** Rekod jawapan pelajar bagi satu item. Awam (tanpa PIN) — hanya untuk set yang Aktif. */
function submitJawapan(j) {
  return withLock_(function () {
    const set = readTable_(SH.SET).filter(function (r) { return String(r.ID) === String(j.Set_ID); })[0];
    if (!set || set.Status !== 'Aktif') throw new Error('Set latihan ini tidak aktif.');
    const nama = clean_(j.Nama_Pelajar).slice(0, 80), kelas = clean_(j.Kelas).slice(0, 30);
    if (!nama) throw new Error('Nama pelajar wajib diisi.');
    let kod = '';
    if (berbayar_()) {
      const k = cariKod_(j.Kod);
      if (!k || !kodSah_(k)) throw new Error('Kod akses tidak sah atau telah tamat. Masukkan kod yang sah untuk menghantar jawapan.');
      kod = k.Kod;
    }
    const n = function (v) { v = Number(v); return isFinite(v) ? v : 0; };
    let butiran = '[]';
    try { butiran = JSON.stringify(JSON.parse(j.Butiran_JSON || '[]')).slice(0, 20000); } catch (e) { butiran = '[]'; }
    return appendRow_(SH.JAWAPAN, 'JL-', {
      Set_ID: set.ID, Tahun: n(set.Tahun) || new Date().getFullYear(), Nama_Pelajar: nama, Kelas: kelas,
      Item_No: n(j.Item_No), Teks: clean_(j.Teks).slice(0, 500), Bahar_Sebenar: clean_(j.Bahar_Sebenar), Laqob_Sebenar: clean_(j.Laqob_Sebenar),
      Bahar_Jawapan: clean_(j.Bahar_Jawapan), Laqob_Jawapan: clean_(j.Laqob_Jawapan), Bahar_OK: j.Bahar_OK ? 1 : 0, Laqob_OK: j.Laqob_OK ? 1 : 0,
      Taqti_Betul: n(j.Taqti_Betul), Taqti_Jumlah: n(j.Taqti_Jumlah), Taf_Betul: n(j.Taf_Betul), Taf_Jumlah: n(j.Taf_Jumlah),
      Hal_Betul: n(j.Hal_Betul), Hal_Jumlah: n(j.Hal_Jumlah), Skor: n(j.Skor), Butiran_JSON: butiran,
      Tarikh: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), Kod: kod
    });
  });
}


/* ───────────── Akses berbayar (bayaran manual + kod akses) ───────────── */

const KOD_ABJAD = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // tanpa I, O, 0, 1
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function now_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
function berbayar_() { return String(readSettings_().Mod_Akses || '').trim().toLowerCase() === 'berbayar'; }

function pakej_(st) {
  return String(st.Pakej_Harga || '1:10, 3:25, 6:40').split(',').map(function (x) {
    const m = String(x).match(/(\d+)\s*:\s*([\d.]+)/); return m ? { bulan: Number(m[1]), harga: Number(m[2]) } : null;
  }).filter(function (p) { return p && p.bulan > 0 && p.bulan <= 24; });
}
function aksesInfo_(st) {
  return { berbayar: String(st.Mod_Akses || '').trim().toLowerCase() === 'berbayar', pakej: pakej_(st),
    bayaran: String(st.Maklumat_Bayaran || ''), qr: String(st.QR_Bayaran_URL || ''), hadPeranti: Number(st.Had_Peranti) || 2 };
}

function janaKodBaharu_() {
  const ada = {}; readTable_(SH.KOD).forEach(function (r) { ada[String(r.Kod)] = 1; });
  for (let t = 0; t < 50; t++) {
    const b = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    let s = '';
    for (let i = 0; i < 8; i++) s += KOD_ABJAD[parseInt(b.substr(i * 2, 2), 16) % KOD_ABJAD.length];
    const kod = 'AR-' + s.slice(0, 4) + '-' + s.slice(4);
    if (!ada[kod]) return kod;
  }
  throw new Error('Gagal menjana kod unik. Cuba lagi.');
}
function normKod_(k) { return String(k || '').toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/^AR/, ''); }
function cariKod_(kod) {
  const n = normKod_(kod); if (n.length !== 8) return null;
  return readTable_(SH.KOD).filter(function (r) { return normKod_(r.Kod) === n; })[0] || null;
}
function kodSah_(k) { return k && k.Status === 'Aktif' && String(k.Tamat).slice(0, 10) >= today_(); }
function tambahBulan_(ymd, n) {
  const p = ymd.split('-').map(Number); const d = new Date(p[0], p[1] - 1 + n, p[2]);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function setCell_(sheet, id, field, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheet), h = HEADERS[sheet];
  const ids = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) { sh.getRange(i + 2, h.indexOf(field) + 1).setValue(value); return true; }
  throw new Error('ID ' + id + ' tidak dijumpai.');
}
function buatKod_(o) {
  const mula = today_();
  return appendRow_(SH.KOD, 'KA-', { Kod: janaKodBaharu_(), Nama: clean_(o.Nama).slice(0, 80), Kelas_Sekolah: clean_(o.Kelas_Sekolah).slice(0, 80),
    Telefon: o.Telefon ? "'" + clean_(o.Telefon).replace(/^'/, '') : '', Bulan: Number(o.Bulan), Mula: mula, Tamat: tambahBulan_(mula, Number(o.Bulan)),
    Status: 'Aktif', Peranti: '', Permohonan_ID: o.Permohonan_ID || '', Dicipta: now_(), Catatan: clean_(o.Catatan) });
}
function kodAwam_(k) { return { Kod: k.Kod, Nama: k.Nama, Kelas_Sekolah: k.Kelas_Sekolah, Tamat: String(k.Tamat).slice(0, 10), Bulan: k.Bulan }; }

/** Pelajar: hantar permohonan beli + resit (base64, ≤ 5 MB). Awam. */
function mohonAkses(f) {
  const st = readSettings_();
  const nama = clean_(f.Nama).slice(0, 80), tel = clean_(f.Telefon).replace(/[^\d+]/g, '').slice(0, 15);
  if (nama.length < 3) throw new Error('Tulis nama penuh.');
  if (tel.length < 9) throw new Error('Tulis nombor telefon yang sah.');
  const pk = pakej_(st).filter(function (p) { return p.bulan === Number(f.Pakej_Bulan); })[0];
  if (!pk) throw new Error('Pakej tidak sah.');
  if (!clean_(f.Rujukan_Bayaran) && !f.resit) throw new Error('Masukkan no. rujukan bayaran atau muat naik resit.');
  let url = '';
  if (f.resit && f.resit.data) {
    const mime = String(f.resit.type || '');
    if (!/^image\/|^application\/pdf$/.test(mime)) throw new Error('Resit mesti gambar atau PDF.');
    const bytes = Utilities.base64Decode(f.resit.data);
    if (bytes.length > 5 * 1024 * 1024) throw new Error('Saiz resit melebihi 5 MB.');
    const it = DriveApp.getFoldersByName('Resit Akses Latihan Arudh');
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder('Resit Akses Latihan Arudh');
    const file = folder.createFile(Utilities.newBlob(bytes, mime, now_().replace(/[: ]/g, '-') + '_' + nama.replace(/[^\w\- ]/g, '').slice(0, 30) + '_' + (f.resit.name || 'resit')));
    url = file.getUrl();
  }
  const row = withLock_(function () {
    return appendRow_(SH.MOHON, 'PA-', { Tarikh: now_(), Nama: nama, Kelas_Sekolah: clean_(f.Kelas_Sekolah).slice(0, 80), Telefon: "'" + tel,
      Emel: clean_(f.Emel).slice(0, 120), Pakej_Bulan: pk.bulan, Jumlah: pk.harga, Rujukan_Bayaran: clean_(f.Rujukan_Bayaran).slice(0, 60),
      Resit_URL: url, Status: 'Menunggu', Kod: '', Diproses: '', Catatan: '' });
  });
  const to = String(st.Emel_Notifikasi || '').trim();
  if (to) { try { MailApp.sendEmail(to, "Permohonan akses Taqti' Pro: " + row.ID, nama + ' (' + tel + ') memohon pakej ' + pk.bulan + ' bulan, RM' + pk.harga + '.\nRujukan: ' + row.Rujukan_Bayaran + '\nResit: ' + (url || '-') + '\n\nLuluskan di tab Akses (mod guru).'); } catch (e) {} }
  return { ID: row.ID, Status: row.Status, Pakej_Bulan: pk.bulan, Jumlah: pk.harga };
}

/** Pelajar: semak status permohonan (ID + 4 digit akhir telefon). Kod dipaparkan bila diluluskan. */
function semakPermohonan(id, tel4) {
  const r = readTable_(SH.MOHON).filter(function (x) { return String(x.ID) === String(id).trim().toUpperCase(); })[0];
  const t = String(tel4 || '').replace(/\D/g, '');
  if (!r || t.length < 4 || String(r.Telefon).replace(/\D/g, '').slice(-4) !== t.slice(-4)) throw new Error('Permohonan tidak dijumpai. Semak ID dan nombor telefon.');
  const out = { ID: r.ID, Status: r.Status, Catatan: r.Catatan };
  if (r.Status === 'Diluluskan' && r.Kod) { const k = cariKod_(r.Kod); if (k) out.kod = kodAwam_(k); }
  return out;
}

/** Pelajar: sahkan kod & daftar peranti (had ikut Tetapan). */
function semakKod(kod, peranti) {
  const k = cariKod_(kod);
  if (!k) throw new Error('Kod tidak dijumpai. Semak semula ejaan kod.');
  if (k.Status !== 'Aktif') throw new Error('Kod ini telah dibatalkan. Hubungi guru.');
  if (String(k.Tamat).slice(0, 10) < today_()) throw new Error('Kod ini telah tamat pada ' + String(k.Tamat).slice(0, 10) + '. Sila beli pakej baharu.');
  const dev = String(peranti || '').replace(/[^\w-]/g, '').slice(0, 40);
  if (dev) {
    return withLock_(function () {
      const fresh = cariKod_(kod);
      const list = String(fresh.Peranti || '').split(',').map(function (x) { return x.trim(); }).filter(String);
      if (list.indexOf(dev) < 0) {
        const had = Number(readSettings_().Had_Peranti) || 2;
        if (list.length >= had) throw new Error('Kod ini sudah digunakan pada ' + had + ' peranti. Hubungi guru untuk set semula.');
        list.push(dev); setCell_(SH.KOD, fresh.ID, 'Peranti', list.join(','));
      }
      return kodAwam_(fresh);
    });
  }
  return kodAwam_(k);
}

/** Guru: luluskan permohonan → jana kod. */
function luluskanPermohonan(id, pin) {
  checkPin_(pin);
  return withLock_(function () {
    const r = readTable_(SH.MOHON).filter(function (x) { return String(x.ID) === String(id); })[0];
    if (!r) throw new Error('Permohonan tidak dijumpai.');
    if (r.Status === 'Diluluskan') throw new Error('Permohonan ini sudah diluluskan (' + r.Kod + ').');
    const k = buatKod_({ Nama: r.Nama, Kelas_Sekolah: r.Kelas_Sekolah, Telefon: r.Telefon, Bulan: r.Pakej_Bulan, Permohonan_ID: r.ID });
    setCell_(SH.MOHON, r.ID, 'Status', 'Diluluskan'); setCell_(SH.MOHON, r.ID, 'Kod', k.Kod); setCell_(SH.MOHON, r.ID, 'Diproses', now_());
    if (r.Emel) { try { MailApp.sendEmail(r.Emel, "Kod akses Taqti' Pro anda", 'Assalamualaikum ' + r.Nama + ',\n\nBayaran anda telah disahkan.\nKod akses: ' + k.Kod + '\nSah sehingga: ' + k.Tamat + '\n\nMasukkan kod ini di tab Latihan.'); } catch (e) {} }
    return { mohon: Object.assign({}, r, { Status: 'Diluluskan', Kod: k.Kod, Diproses: now_() }), kod: k };
  });
}
function tolakPermohonan(id, sebab, pin) {
  checkPin_(pin);
  return withLock_(function () {
    setCell_(SH.MOHON, id, 'Status', 'Ditolak'); setCell_(SH.MOHON, id, 'Catatan', clean_(sebab).slice(0, 200)); setCell_(SH.MOHON, id, 'Diproses', now_());
    return true;
  });
}
/** Guru: jana kod terus (promosi / bayaran tunai). */
function janaKodManual(o, pin) {
  checkPin_(pin);
  if (clean_(o.Nama).length < 2) throw new Error('Tulis nama penerima kod.');
  if (!(Number(o.Bulan) >= 1 && Number(o.Bulan) <= 24)) throw new Error('Tempoh 1–24 bulan.');
  return withLock_(function () { return buatKod_(o); });
}
/** Guru: batal / aktifkan semula / set semula peranti / lanjut tempoh. */
function urusKod(id, tindakan, nilai, pin) {
  checkPin_(pin);
  return withLock_(function () {
    const k = readTable_(SH.KOD).filter(function (x) { return String(x.ID) === String(id); })[0];
    if (!k) throw new Error('Kod tidak dijumpai.');
    if (tindakan === 'batal') setCell_(SH.KOD, id, 'Status', 'Batal');
    else if (tindakan === 'aktif') setCell_(SH.KOD, id, 'Status', 'Aktif');
    else if (tindakan === 'peranti') setCell_(SH.KOD, id, 'Peranti', '');
    else if (tindakan === 'lanjut') {
      const n = Number(nilai); if (!(n >= 1 && n <= 24)) throw new Error('Tempoh 1–24 bulan.');
      const base = String(k.Tamat).slice(0, 10) > today_() ? String(k.Tamat).slice(0, 10) : today_();
      setCell_(SH.KOD, id, 'Tamat', tambahBulan_(base, n)); setCell_(SH.KOD, id, 'Status', 'Aktif');
    } else throw new Error('Tindakan tidak sah.');
    return readTable_(SH.KOD).filter(function (x) { return String(x.ID) === String(id); })[0];
  });
}

/** Guru: tukar Mod_Akses (Percuma / Berbayar) terus dari aplikasi. */
function setModAkses(mode, pin) {
  checkPin_(pin);
  const v = String(mode) === 'Berbayar' ? 'Berbayar' : 'Percuma';
  return withLock_(function () {
    const sh = SpreadsheetApp.getActive().getSheetByName(SH.TETAPAN);
    const keys = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues().map(function (r) { return String(r[0]).trim(); });
    const i = keys.indexOf('Mod_Akses');
    if (i < 0) sh.appendRow(['Mod_Akses', v, 'Percuma = Latihan terbuka kepada semua. Berbayar = Latihan perlu kod akses.']);
    else sh.getRange(i + 1, 2).setValue(v);
    return aksesInfo_(readSettings_());
  });
}
