# BPS Report Generator

**Laporan Penilaian Biopsikososial (PKSP)**
Kementerian Kesihatan Malaysia

App Flask untuk jana laporan BPS dalam format `.docx` dan `.pdf` terus dari borang web.
Optional: guna AI (OpenRouter / DeepSeek) untuk bantu draf syor — perlu internet.

---

## 🚀 Cara PALING SENANG (untuk pengguna)

1. Dapatkan fail **`BPS-Portable.zip`** dari pengedar (orang yang bagi app ni).
2. Unzip ke mana-mana folder.
3. **Double-click `BPS.exe`** → browser akan terbuka sendiri ke `http://localhost:5000`.
4. Isi borang, tekan **Jana Laporan**.

✅ Tak perlu install Python. ✅ Tak perlu pip. ✅ Tak perlu internet (kecuali nak guna AI).

```
BPS-Portable/
├── BPS.exe            ← double-click ni
├── static/            ← letak gambar letterhead hospital
├── output/            ← laporan disimpan sini
└── config.txt         ← ubah OUTPUT_PATH kalau nak tukar folder simpan
```

---

## 🛠️ Cara BUILD SENDIRI (untuk developer)

Kalau kau clone/download source code dari GitHub:

```bat
# 1. Pastikan ada Python 3.11+ (atau guna uv.exe yang dah ada sekali)
# 2. Buat environment & install dependensi
setup.bat

# 3. Build versi portable (hasilkan dist/BPS-Portable/BPS.exe)
build_portable.bat
```

Atau pasang manual:

```bat
python -m venv .venv
.venv\Scripts\pip install flask python-docx fpdf2 requests pyinstaller
python -m PyInstaller --noconsole --onefile --name BPS ^
    --add-data "templates;templates" --add-data "static;static" ^
    --hidden-import weasyprint --icon BPS.ico app.py
```

---

## 📌 Letterhead Hospital

Letak gambar letterhead hospital korang di `static/`:

```
static/letterhead-header.png
static/letterhead-footer.png
```

Kalau takde, app tetap jalan — header/footer kosong.

---

## 💾 Di mana laporan disimpan?

1. `OUTPUT_PATH` dalam `config.txt` (kalau diisi & wujud)
2. Folder `output/` di sebelah `BPS.exe` (lalai)

---

## 🤖 Guna AI (Online)

- Buka menu syor → pilih **Generate Syor (AI)**.
- Perlu internet + API key OpenRouter (letak di config atau `.env`).
- Tanpa internet → guna **Generate Syor (Offline)** (template tetap).

---

## ⚠️ Nota Penting

- **Auto-update TIDAK aktif** dalam versi EXE. Kalau app dikemaskini,
  dapatkan `BPS-Portable.zip` terkini dari pengedar.
- Data pesakit (`output/`) dan `config.txt` **tidak dimasukkan ke GitHub**
  (gitignored) — jaga kerahsiaan pesakit.

---

## 📞 Sokongan

Hubungi pengedar app (Pegawai Kerja Sosial Perubatan) untuk sebarang isu.
