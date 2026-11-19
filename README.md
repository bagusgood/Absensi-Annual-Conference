# Absensi Digital — Annual Conference 2026

Aplikasi check-in peserta berbasis Flask + SQLite dengan latar backdrop acara.

## Fitur
- Form check-in peserta: nama, instansi, dan jenis kehadiran (Luring/Daring) wajib; jabatan, WhatsApp, email opsional.
- Mencegah entri ganda berdasarkan kombinasi nama + instansi.
- Dashboard admin dengan pencarian dan pembaruan live otomatis setiap 5 detik.
- QR Code check-in otomatis berdasarkan alamat URL aplikasi, dapat diunduh dari dashboard admin.
- Ekspor daftar kehadiran ke Excel (.xlsx).
- Database SQLite otomatis dibuat sebagai `attendance.db`.

## Menjalankan di Windows
1. Ekstrak ZIP ini.
2. Buka Command Prompt di folder hasil ekstrak.
3. Buat virtual environment (opsional tetapi disarankan):
   `py -m venv venv`
4. Aktifkan:
   `venv\Scripts\activate`
5. Instal dependensi:
   `pip install -r requirements.txt`
6. Jalankan:
   `python app.py`
7. Buka `http://127.0.0.1:5000`.

Untuk diakses perangkat lain pada jaringan yang sama, buka `http://IP-KOMPUTER-SERVER:5000`.

## Login admin
- URL: `/admin/login`
- Username default: `admin`
- Password default: `GantiPasswordAdmin2026!`

**WAJIB sebelum digunakan:** ubah password dan SECRET_KEY melalui environment variable.

Contoh PowerShell:
```powershell
$env:ADMIN_USERNAME="admin"
$env:ADMIN_PASSWORD="PasswordKuatAnda"
$env:SECRET_KEY="string-acak-panjang-dan-rahasia"
python app.py
```

## Pengaturan hari acara
Secara default, check-in dapat dicoba kapan saja untuk pengujian.
Pada hari acara, aktifkan pembatasan tanggal dengan:
```powershell
$env:ENFORCE_EVENT_DATE="1"
python app.py
```
Dengan pengaturan tersebut, check-in hanya diterima pada 24 November 2026 (mengikuti tanggal lokal komputer server).

## QR Code
Untuk membuat QR absensi, buat QR yang mengarah ke alamat aplikasi, misalnya:
`http://IP-KOMPUTER-SERVER:5000/`
Jika aplikasi dipublikasikan melalui domain/HTTPS, gunakan URL HTTPS tersebut.

## Catatan operasional
- Jalankan aplikasi pada satu komputer/server yang stabil dan pastikan perangkat peserta berada pada jaringan yang dapat mengakses server.
- Untuk penggunaan di luar jaringan lokal, gunakan hosting/server dengan HTTPS.
- Data tersimpan di `attendance.db`. Lakukan backup file ini setelah acara.
- Jangan membagikan kredensial admin kepada peserta.


## QR Code dan Live Dashboard
1. Jalankan aplikasi dan buka `http://127.0.0.1:5000/admin/login`.
2. Login admin, lalu QR Code tampil pada dashboard.
3. Klik **Unduh QR Code** untuk menyimpan gambar QR.
4. Cetak QR atau tampilkan di layar registrasi. Peserta memindai QR dengan kamera HP, lalu mengisi form.
5. Dashboard admin memperbarui jumlah dan daftar kehadiran otomatis setiap 5 detik.

**Penting untuk acara:** QR Code menggunakan alamat host/domain yang sedang dipakai saat admin membuka dashboard. Jika peserta memakai HP, jangan gunakan `127.0.0.1` atau `localhost` pada URL. Gunakan IP lokal server (contoh `http://192.168.1.10:5000/`) atau domain publik HTTPS. Pastikan HP peserta dan server terhubung ke jaringan yang sama jika memakai IP lokal.

Jika QR sudah terlanjur diunduh dari alamat localhost, buka dashboard menggunakan alamat IP server atau domain yang dapat diakses peserta, lalu unduh QR lagi.


Database lama otomatis dimigrasikan dengan nilai default Luring untuk data absensi yang sudah ada. Jenis kehadiran baru tampil di dashboard live dan file Excel.


## Gambar halaman check-in berhasil
Dua foto sambutan ditampilkan berdampingan pada halaman konfirmasi setelah peserta berhasil mengisi absensi. File gambar berada di `static/greeting_man.png` dan `static/greeting_woman.png`.
