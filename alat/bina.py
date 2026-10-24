"""
Bina docs/index.html untuk GitHub Pages daripada fail dalam gas/.

Cara guna (dari akar repo):
    python alat/bina.py

Ia membaca gas/Index.html, menggantikan setiap
<?!= include('X'); ?> dengan kandungan X.html, dan memuatkan api.js
sebelum skrip aplikasi. Jalankan semula setiap kali fail GAS dikemas kini.
"""
import re
from pathlib import Path

SINI = Path(__file__).resolve().parent.parent / "docs"
SUMBER = SINI.parent / "gas"

html = (SUMBER / "Index.html").read_text(encoding="utf-8")

def ganti_include(m):
    nama = m.group(1)
    return (SUMBER / f"{nama}.html").read_text(encoding="utf-8")

html, n = re.subn(r"<\?!=\s*include\('([^']+)'\);?\s*\?>", ganti_include, html)

if "<?" in html:
    raise SystemExit("Masih ada scriptlet GAS (<? ... ?>) yang belum ditukar.")

# <base target="_top"> hanya perlu dalam iframe GAS
html = html.replace('<base target="_top">\n', "", 1)

# Tajuk tab (dalam GAS ia diset oleh doGet)
if "<title>" not in html:
    html = html.replace(
        '<meta charset="utf-8">',
        "<meta charset=\"utf-8\">\n<title>Taqti' Pro — Latihan 'Arudh &amp; Qafiyah</title>",
        1,
    )

# api.js mesti dimuat sebelum App/Latihan/Sets/Akses
import time
html = html.replace("</head>", f'<script src="api.js?v={int(time.time())}"></script>\n</head>', 1)

(SINI / "index.html").write_text(html, encoding="utf-8")
print(f"index.html siap ({n} fail digabungkan, {len(html)//1024} KB).")
