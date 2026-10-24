# Taqti' Pro

Latihan & semakan syair 'Arudh & Qafiyah. Frontend di GitHub Pages, pelayan di Google Apps Script + Google Sheet.

Asal: `GAS-projek/SEMAK BAHAR KOMERSIL` (v2.8) — kekal sebagai rujukan, tidak diubah.

| Folder | Isi |
|---|---|
| `gas/` | Kod Apps Script (v2.9 = v2.8 + `doPost` API). Salin ke editor Apps Script. |
| `docs/` | Frontend GitHub Pages: `index.html` (dijana) + `api.js` (URL pelayan). |
| `alat/bina.py` | Jana semula `docs/index.html` daripada fail HTML dalam `gas/`. |

## Pemasangan

1. **Apps Script** — salin `gas/Code.js` ke `Code.gs`. *Deploy → Manage deployments* → edit → *New version*.
   *Execute as: Me*, *Who has access: Anyone*. Salin URL `/exec`.
2. **`docs/api.js`** — gantikan `TAMPAL_URL_WEB_APP_DI_SINI` dengan URL `/exec`.
3. **GitHub Pages** — *Settings → Pages* → *Deploy from a branch* → `main`, folder **`/docs`**.

## Kemas kini

- Ubah `gas/Code.js` → salin ke Apps Script → deploy *New version*.
- Ubah `gas/*.html` → `python alat/bina.py` → commit `docs/index.html`.

## Nota

- PIN guru, kod akses dan had peranti disemak di pelayan.
- Pelajar yang guna URL GAS lama akan dikira peranti baharu di GitHub Pages. Set semula peranti di tab Akses jika tersekat.
