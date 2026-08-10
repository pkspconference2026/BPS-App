# 📋 PANDUAN UPDATE BPS APP
## Cara keluarkan versi baru untuk kawan-kawan

---

### ⚙️ Langkah-langkah

**Langkah 1 — Ubah fail**
Edit `app.py` / `templates/` / `static/` macam biasa.

**Langkah 2 — Naikkan versi**
Buka `app.py`, cari:
```python
APP_VERSION = '1.0.1'
```
Tukar nombor ikut pembaharuan:
- Pembetulan kecil → `1.0.1` → `1.0.2`
- Ciri baru → `1.1.0`
- Perubahan besar → `2.0.0`

**Langkah 3 — Update manifest**
Buka `update_manifest.json`, ubah `"version"` dan `"change_log"`:
```json
{
  "version": "1.0.2",
  "change_log": "Terangkan perubahan di sini",
  "files": [
    { "path": "app.py", "url": "https://raw.githubusercontent.com/pkspconference2026/BPS-App/main/app.py" },
    { "path": "templates/index.html", "url": "https://raw.githubusercontent.com/pkspconference2026/BPS-App/main/templates/index.html" },
    { "path": "update_manifest.json", "url": "https://raw.githubusercontent.com/pkspconference2026/BPS-App/main/update_manifest.json" }
  ]
}
```
> Senarai `"files"` hanya fail yang berubah. `update_manifest.json` mesti sentiasa ada.

**Langkah 4 — Commit & push**
```bash
cd C:\Users\User\Desktop\BPS-App
git add -A
git commit -m "v1.0.2: ringkasan perubahan"
git push
```

**Langkah 5 — Info kawan**
Kawan buka app → tekan **🔄 Semak Update** → OK → restart app. Mereka dah dapat versi baru! 🎉

---

### ✅ Checklist Setiap Update
- [ ] APP_VERSION dalam app.py naik
- [ ] Version dalam update_manifest.json sama
- [ ] change_log ditulis ringkas dan jelas
- [ ] files hanya senarai fail yang berubah
- [ ] git push berjaya (takde error rejected)

---

### ⚠️ Nota Penting
- Kalau `git push` rejected, guna `git pull` dulu pastu `git push` semula
- `config.txt` setiap PC berbeza — **jangan ubah** untuk PC orang lain
- Kalau PC kawan takde internet, update kena guna USB (cara lama)

### 🔗 Pautan Repo
https://github.com/pkspconference2026/BPS-App