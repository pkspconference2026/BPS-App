# Rumusan BPS Report Generator

**Lokasi projek:** `C:\Users\User\Desktop\BPS-App`  
**Repo GitHub:** https://github.com/pkspconference2026/BPS-App  
**Versi semasa:** `v1.2.1`  
**Branch:** `main`

## 1. Konfigurasi Hospital dan Pegawai

Maklumat berikut boleh ditukar dalam `config.txt`:

```ini
NAMA_HOSPITAL = Hospital Hulu Terengganu
NAMA_PEGAWAI =
NAMA_SISTEM = Kementerian Kesihatan Malaysia
```

- Jawatan kekal: **PEGAWAI KERJA SOSIAL PERUBATAN**.
- `NAMA_PEGAWAI` sengaja dikosongkan supaya pengguna di PC lain mengisinya sendiri.
- Input `nama_pk` dalam borang masih diutamakan berbanding nilai konfigurasi.
- Letterhead Surat Iringan boleh ditukar dengan menggantikan:
  - `static/letterhead-header.png`
  - `static/letterhead-footer.png`

## 2. Header Sistem

Header sistem dipaparkan secara umum sebagai:

> Pegawai Kerja Sosial Perubatan – Kementerian Kesihatan Malaysia

## 3. Rujukan Surat Iringan TDI

Placeholder bagi **No. Rujukan Fail** telah ditukar kepada:

```text
MSW37/26
```

Apabila pengguna memasukkan `MSW37/26`, Surat Iringan TDI akan memaparkan:

```text
Ruj. Kami : UKSP/HHT/MSW37/26
```

Tahun penuh `/2026` tidak ditambah kerana `/26` sudah menunjukkan tahun 2026.

### Prioriti rujukan

1. Rujukan TDI yang dimasukkan secara manual.
2. `UKSP/HHT/{No. Rujukan Fail}`.
3. Fallback lama: `HHT/UKSP/MSW15/{tahun}`.

## 4. Pembetulan Lain

- Ejaan **Terengganu** telah dibetulkan.
- Favicon BPS telah ditambah.
- Folder nested lama `BPS-App/BPS-App/` telah dibuang.
- Pembacaan `config.txt` menggunakan padanan terakhir supaya nilai override berfungsi.
- Fungsi Syor Offline dan Syor AI DeepSeek telah diuji dan berfungsi.

## 5. Git Terakhir

- `c711154` — Logic Rujukan Kami TDI dan placeholder `MSW37/26`.
- `74c03fd` — Manifest `v1.2.1`.
- Semua perubahan telah di-commit dan push ke GitHub.

## 6. Portable dan GitHub Release

Halaman release:

https://github.com/pkspconference2026/BPS-App/releases/latest

> **Perhatian:** ZIP portable yang pernah dibina mungkin masih mengandungi build `v1.2.0`. Kod GitHub dan auto-update sudah berada pada `v1.2.1`. Untuk ZIP yang terus mengandungi perubahan rujukan terkini, bina semula EXE/ZIP `v1.2.1` dan kemas kini GitHub Release.

## 7. Desktop Shortcut

- `BPS Report.lnk` — membuka sistem BPS.
- `Stop BPS.lnk` — menghentikan server BPS.

## 8. Perangkap Penting: Dua Server pada Port 5000

Kadangkala `pythonw.exe` daripada shortcut dan `python.exe` daripada ujian developer berjalan serentak pada port 5000. Keadaan ini menyebabkan perubahan kod kelihatan seperti tidak berfungsi kerana permintaan dijawab oleh server lama.

Semak listener dengan:

```bash
netstat -ano | grep ":5000.*LISTEN"
```

Pastikan hanya **satu listener** sebelum menguji perubahan.

## 9. Keadaan Akhir

- Fungsi `MSW37/26` → `UKSP/HHT/MSW37/26` telah diuji secara langsung dan berjaya.
- Folder ujian dalam Dropbox telah dibersihkan.
- Tiada tugasan tambahan yang belum selesai selain pilihan untuk rebuild portable `v1.2.1`.
- Komunikasi dengan pengguna hendaklah menggunakan **Bahasa Melayu kasual sepenuhnya**.
