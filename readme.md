# Sistem Penjadwalan Otomatis Sidang Tugas Akhir

Sistem ini merupakan aplikasi berbasis web untuk menjadwalkan sidang tugas akhir secara otomatis dengan mempertimbangkan berbagai kendala dan prioritas. Aplikasi dirancang untuk membantu program studi dalam mengelola jadwal sidang TA dengan efisien, meminimalkan konflik jadwal, dan memaksimalkan penggunaan waktu.

## Fitur Utama

1. **Penjadwalan Otomatis** - Mengalokasikan jadwal sidang TA secara otomatis dengan mempertimbangkan berbagai kendala.
2. **Upload Data via Excel** - Mengimpor data jadwal mengajar dan tim sidang dari file Excel.
3. **Request Jadwal** - Memungkinkan mahasiswa meminta jadwal spesifik melalui file Excel.
4. **Manajemen Ketidaktersediaan Dosen** - Dosen dapat menandai slot waktu di mana mereka tidak tersedia.
5. **Visualisasi Kalender** - Menampilkan jadwal dalam bentuk kalender mingguan.
6. **Ekspor Hasil** - Mengekspor hasil penjadwalan ke format Excel.
7. **Analisis Slot Tersedia** - Menampilkan slot waktu yang masih tersedia beserta dosen yang tersedia.
8. **Filter dan Pencarian** - Memfilter jadwal berdasarkan hari, sesi, pekan, dan dosen.
9. **Dukungan Hari Libur** - Mengecualikan tanggal libur dari penjadwalan.

## Algoritma dan Persyaratan Prioritas

Sistem menggunakan algoritma penjadwalan berbasis dosen dengan prioritas sebagai berikut:

1. **Request Jadwal = Mutlak**: Request jadwal dari Excel diutamakan, bahkan jika bentrok dengan jadwal mengajar dan ketidaktersediaan dosen.
2. **Konflik Dosen = Dilarang**: Dosen tidak boleh terjadwal di dua sidang bersamaan.
3. **Ketidaktersediaan Dosen = Dipatuhi**: Dosen tidak akan dijadwalkan pada slot di mana mereka menyatakan tidak tersedia.
4. **Sidang Paralel = Diperbolehkan**: Boleh ada beberapa sidang pada slot yang sama jika dosennya berbeda.
5. **Jadwal Mengajar < Request Sidang**: Jadwal mengajar bisa "digantikan" oleh request sidang.
6. **Maksimalkan Penjadwalan**: Usahakan semua mahasiswa mendapat jadwal sidang.

## Cara Kerja

### 1. Algoritma Penjadwalan

Aplikasi menggunakan algoritma dua tahap:

#### Tahap 1: Penjadwalan Request (Prioritas Mutlak)
- Jadwalkan semua request dari Excel terlebih dahulu.
- Request akan diutamakan bahkan jika bentrok dengan jadwal mengajar.
- Jika terjadi bentrok antar request (dosen yang sama pada slot yang sama), request yang lebih dulu akan diprioritaskan.

#### Tahap 2: Penjadwalan Mahasiswa Tanpa Request
- Jadwalkan mahasiswa yang tidak memiliki request pada slot-slot yang tersisa.
- Mencari slot yang tidak bentrok dengan jadwal sidang lain dan memperhatikan ketidaktersediaan dosen.
- Jika tidak ada slot tanpa bentrok, jadwalkan pada slot dengan bentrok jadwal mengajar (jadwal mengajar akan digantikan).

### 2. Struktur Data

Sistem menggunakan struktur data berikut:
- `jadwalHariTanggal`: Menyimpan informasi slot waktu untuk setiap hari dan tanggal.
- `dosenSidangTerpakai`: Melacak dosen yang sudah dijadwalkan sidang per slot.
- `ketidaktersediaanDosen`: Menyimpan data slot di mana dosen tidak tersedia.
- `jadwalMengajar`: Menyimpan jadwal mengajar dosen.
- `timSidang`: Menyimpan data tim penguji dan pembimbing untuk setiap mahasiswa.

## Cara Penggunaan

### 1. Setup Awal
1. Download template Excel.
2. Isi data jadwal mengajar pada sheet "JadwalMengajar".
3. Isi data tim sidang dan request jadwal pada sheet "TimSidang".

### 2. Penjadwalan
1. Pilih tanggal mulai sidang.
2. Tentukan jumlah minggu penjadwalan.
3. Masukkan tanggal libur jika ada.
4. Atur ketidaktersediaan dosen jika diperlukan.
5. Upload file Excel yang sudah diisi.
6. Sistem akan melakukan penjadwalan otomatis.

### 3. Melihat Hasil
1. Tab "Hasil Penjadwalan" menampilkan hasil jadwal sidang.
2. Tab "Slot Tersedia" menampilkan slot waktu yang masih tersedia.
3. Tab "Kalender" menampilkan visualisasi jadwal dalam bentuk kalender.

### 4. Export Data
- Klik "Export Hasil ke Excel" untuk mengunduh hasil penjadwalan.
- Klik "Export Slot Tersedia ke Excel" untuk mengunduh data slot tersedia.

## Format File Excel

### Sheet JadwalMengajar
- `Hari`: Hari jadwal mengajar (Senin-Jumat)
- `Sesi`: Sesi jadwal (1-4)
- `Jam`: Rentang waktu sesi
- `Ruangan`: Ruangan untuk mengajar
- `Mata Kuliah`: Nama mata kuliah
- `Kode Dosen`: Kode dosen pengajar

### Sheet TimSidang
- `Nama Mahasiswa / NIM`: Nama dan NIM mahasiswa
- `Pembimbing 1`: Nama dosen pembimbing 1
- `Pembimbing 2`: Nama dosen pembimbing 2
- `Penguji 1`: Nama dosen penguji 1
- `Penguji 2`: Nama dosen penguji 2
- `Request Tanggal`: Tanggal request sidang (format: DD/MM/YYYY)
- `Request Sesi`: Sesi request sidang (1-4)

## Jadwal Jam dan Sesi

Sistem menggunakan 4 sesi jadwal per hari:
- Sesi 1: 08:00-09:30
- Sesi 2: 10:30-12:00
- Sesi 3: 13:00-14:30
- Sesi 4: 16:00-17:30

## Pengaturan Pekan

Pekan (minggu) dihitung berdasarkan tanggal mulai sidang:
- **Pekan 1**: 7 hari pertama dari tanggal mulai
- **Pekan 2**: 7 hari berikutnya
- dst.

Contoh: Jika sidang dimulai Kamis, 8 Mei 2025:
- **Pekan 1**: Kamis, 8 Mei - Rabu, 14 Mei 2025
- **Pekan 2**: Kamis, 15 Mei - Rabu, 21 Mei 2025

## Teknologi yang Digunakan

- HTML5 & CSS3
- JavaScript (ES6+)
- Bootstrap 5 untuk UI
- SheetJS (xlsx) untuk manipulasi file Excel

## Persyaratan Sistem

- Browser web modern dengan dukungan JavaScript ES6+
- Koneksi internet untuk memuat dependensi (Bootstrap, SheetJS)

## Instalasi dan Menjalankan

1. Clone repositori:
   ```bash
   git clone https://github.com/username/sistem-penjadwalan-sidang-ta.git
   ```

2. Buka folder proyek:
   ```bash
   cd sistem-penjadwalan-sidang-ta
   ```

3. Buka file `index.html` menggunakan browser web.

## Keterbatasan dan Pengembangan Masa Depan

- Penambahan fitur notifikasi email untuk dosen dan mahasiswa
- Integrasi dengan sistem akademik
- Penambahan fitur penjadwalan multi-ruangan
- Optimasi algoritma untuk dataset yang sangat besar
- Dukungan untuk sidang daring/hybrid

## Lisensi

© 2025. Hak Cipta Dilindungi.

## Kredit dan Kontributor

Dikembangkan oleh [Nama Anda/Tim]