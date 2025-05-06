// Daftar kode dosen dan nama lengkap (hardcoded)
const daftarDosen = {
    'DNA': 'Dwi Nur Amalia, S.Kom., M.Kom',
    'YTW': 'Yuyun Tri Wiranti, S.Kom., M.MT',
    'PDA': 'Ir. I Putu Deny Arthawan Sugih Prabowo, M.Eng',
    'VKA': 'Vika Fitratunnany Insanittaqwa, S.Kom., M.Kom',
    'HLH': 'Henokh Lugo Hariyanto, S.Si., M.Sc',
    'MIA': 'M. Ihsan Alfani Putera, S.Tr.,Kom, M.Kom',
    'DAP': 'Dwi Arief Prambudi, S.Kom., M.Kom',
    'NNA': 'Nursanti Novi Arisa, S.Pd., M.Kom',
    'ADL': 'Aidil Saputra Kirsan, S.ST., M.Tr.Kom',
    'HIS': 'Hendy Indrawan Sunardi, S.Kom, M.Eng',
    'AWS': 'Arif Wicaksono, S.Kom., M.Kom',
    'SRN': 'Sri Rahayu Natasia, S.Komp., M.Si., M.Sc'
};

// Mapping kode dosen ke nama lengkap
const kodeToDosen = {};
for (const [kode, nama] of Object.entries(daftarDosen)) {
    kodeToDosen[nama] = kode;
}

// Jadwal jam berdasarkan sesi
const jadwalJam = {
    1: '07:30-09:10',
    2: '09:20-11:00',
    3: '13:50-15:30',
    4: '15:50-17:30'
};

// Mapping hari ke indeks
const hariToIndex = {
    'Senin': 1,
    'Selasa': 2,
    'Rabu': 3,
    'Kamis': 4,
    'Jumat': 5
};

// Mapping indeks ke hari
const indexToHari = {
    1: 'Senin',
    2: 'Selasa',
    3: 'Rabu',
    4: 'Kamis',
    5: 'Jumat'
};

// Menyimpan jadwal
let jadwalMengajar = [];
let timSidang = [];
let hasilJadwal = [];
let slotTersedia = [];
let jadwalHariTanggal = {};
let tanggalMulaiSidang = null;
let jumlahMinggu = 0;

// Helper function untuk mendapatkan nomor minggu dari tanggal
function getWeekNumber(date) {
    const firstDayOfYear = new Date(date.getFullYear(), 0, 1);
    const pastDaysOfYear = (date - firstDayOfYear) / 86400000;
    return Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
}

// Perbaikan fungsi parseDate
function parseDate(dateStr) {
    // Jika dateStr adalah angka (serial date Excel)
    if (typeof dateStr === 'number') {
        // Konversi serial date Excel ke Date JavaScript
        // Excel: 1 = 1 Jan 1900, JavaScript: 1 = 1 Jan 1970
        // Selisih: 25569 hari
        const excelEpoch = new Date(1900, 0, 1);
        const jsDate = new Date(excelEpoch);
        jsDate.setDate(excelEpoch.getDate() + dateStr - 2); // -2 karena koreksi leap year Excel
        return jsDate;
    }
    
    // Jika dateStr adalah string (format DD/MM/YYYY)
    if (typeof dateStr === 'string' && dateStr.includes('/')) {
        try {
            const [day, month, year] = dateStr.split('/').map(Number);
            return new Date(year, month - 1, day);
        } catch (error) {
            console.error('Error saat parsing tanggal string:', error);
            return new Date(); // Fallback ke tanggal saat ini
        }
    }
    
    // Fallback: kembalikan tanggal saat ini
    console.error('dateStr bukan format yang didukung:', dateStr);
    return new Date();
}

// Helper function untuk format tanggal ke string (DD/MM/YYYY)
function formatDate(date) {
    return `${date.getDate().toString().padStart(2, '0')}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getFullYear()}`;
}

// Inisialisasi tanggal untuk penjadwalan
function inisialisasiTanggal(tanggalMulai, jumlahMahasiswa) {
    // Konversi string tanggal ke objek Date
    tanggalMulaiSidang = new Date(tanggalMulai);
    
    // Hitung perkiraan jumlah minggu yang dibutuhkan 
    // PERBAIKAN: Menggunakan formula yang lebih realistis berdasarkan jumlah dosen dan mahasiswa
    const totalDosenUnik = Object.keys(daftarDosen).length;
    const rataSlotPerHari = (4 * 5) / (totalDosenUnik / 4); // 4 sesi, 5 hari, rata-rata 4 dosen per tim
    
    // Asumsi setiap dosen bisa hadir dalam 2 sidang per hari maksimal
    const kapasitasHarian = Math.floor(totalDosenUnik * 2 / 4); // tiap sidang butuh 4 dosen
    const jumlahHariMinimal = Math.ceil(jumlahMahasiswa / kapasitasHarian);
    jumlahMinggu = Math.ceil(jumlahHariMinimal / 5) + 1; // +1 untuk buffer
    
    console.log(`Perkiraan jumlah minggu: ${jumlahMinggu} berdasarkan ${jumlahMahasiswa} mahasiswa dan ${totalDosenUnik} dosen`);
    
    // Buat mapping hari ke tanggal untuk jumlahMinggu ke depan
    const jadwalHariTanggal = {};
    
    // Simpan tanggal saat ini untuk iterasi
    let currentDate = new Date(tanggalMulaiSidang);
    
    // Dapatkan hari dari tanggal mulai (0 = Minggu, 1 = Senin, dst)
    const hariMulai = currentDate.getDay();
    
    // Jika hari mulai adalah hari Sabtu (6) atau Minggu (0), sesuaikan ke Senin berikutnya
    if (hariMulai === 0) { // Minggu
        currentDate.setDate(currentDate.getDate() + 1); // Tambah 1 hari ke Senin
    } else if (hariMulai === 6) { // Sabtu
        currentDate.setDate(currentDate.getDate() + 2); // Tambah 2 hari ke Senin
    }
    
    // Total hari kerja yang perlu dijadwalkan
    const totalHariKerja = jumlahMinggu * 5; // 5 hari kerja per minggu
    
    // Iterasi untuk setiap hari kerja
    for (let i = 0; i < totalHariKerja; i++) {
        // Dapatkan nama hari (1 = Senin, ..., 5 = Jumat)
        const dayOfWeek = currentDate.getDay();
        
        // Lewati hari Sabtu dan Minggu
        if (dayOfWeek === 0 || dayOfWeek === 6) {
            currentDate.setDate(currentDate.getDate() + 1);
            continue;
        }
        
        // Konversi angka hari ke nama hari
        const namaHari = indexToHari[dayOfWeek];
        
        // Format tanggal: DD/MM/YYYY
        const tanggalStr = formatDate(currentDate);
        
        // Buat key unik untuk setiap hari+tanggal
        const key = `${namaHari}-${tanggalStr}`;
        
        // Simpan dalam mapping
        jadwalHariTanggal[key] = {
            hari: namaHari,
            tanggal: tanggalStr,
            date: new Date(currentDate), // Simpan objek Date untuk operasi perbandingan
            slots: {} // Akan diisi dengan slot yang tersedia
        };
        
        // Inisialisasi slots untuk setiap sesi
        for (let sesi = 1; sesi <= 4; sesi++) {
            jadwalHariTanggal[key].slots[sesi] = {
                tersedia: true,
                jam: jadwalJam[sesi],
                jadwalMengajar: [],
                jadwalSidang: null
            };
        }
        
        // Lanjut ke hari berikutnya
        currentDate.setDate(currentDate.getDate() + 1);
    }
    
    return jadwalHariTanggal;
}

// Fungsi untuk membuat template Excel
function createTemplate() {
    // Buat workbook baru
    const wb = XLSX.utils.book_new();
    
    // Template jadwal mengajar
    const jadwalData = [
        ['Hari', 'Sesi', 'Jam', 'Ruangan', 'Mata Kuliah', 'Kode Dosen'],
        ['Senin', 1, '07:30-09:10', 'A308', 'Statistika', 'SRN'],
        ['Senin', 1, '07:30-09:10', 'E101', 'Manajemen dan Organisasi B', 'YTW'],
        ['Senin', 1, '07:30-09:10', 'E102', 'Rekayasa Perangkat Lunak B', 'SRN'],
        ['Senin', 2, '10:20-12:00', 'A308', 'Desain Interaksi Antarmuka dan Pengalaman Pengguna B', 'MIA'],
        ['Senin', 2, '10:20-12:00', 'E101', 'Perencanaan Manajemen Proyek Teknologi Informasi A', 'YTW']
    ];
    const jadwalSheet = XLSX.utils.aoa_to_sheet(jadwalData);
    XLSX.utils.book_append_sheet(wb, jadwalSheet, 'JadwalMengajar');
    
    // Template tim sidang dengan kolom request jadwal
    const timData = [
        ['Nama Mahasiswa / NIM', 'Pembimbing 1', 'Pembimbing 2', 'Penguji 1', 'Penguji 2', 'Request Tanggal', 'Request Sesi'],
        ['Athifah Shyla Maritza / 10211021', 'Dwi Nur Amalia, S.Kom., M.Kom', 'Yuyun Tri Wiranti, S.Kom., M.MT', 'Ir. I Putu Deny Arthawan Sugih Prabowo, M.Eng', 'Vika Fitratunnany Insanittaqwa, S.Kom., M.Kom', '09/05/2025', 3],
        ['Annisa Rosmawati / 10211019', 'Yuyun Tri Wiranti, S.Kom., M.MT', 'Henokh Lugo Hariyanto, S.Si., M.Sc', 'M. Ihsan Alfani Putera, S.Tr.,Kom, M.Kom', 'Dwi Arief Prambudi, S.Kom., M.Kom', '', '']
    ];
    const timSheet = XLSX.utils.aoa_to_sheet(timData);
    XLSX.utils.book_append_sheet(wb, timSheet, 'TimSidang');
    
    // Download file
    XLSX.writeFile(wb, 'template_sidang_ta.xlsx');
}

// PERBAIKAN: Fungsi untuk menghitung beban dosen
function hitungBebanDosen(timSidang) {
    const bebanDosen = {};
    
    // Inisialisasi beban untuk semua dosen
    Object.values(daftarDosen).forEach(nama => {
        bebanDosen[nama] = 0;
    });
    
    // Hitung berapa kali setiap dosen muncul dalam tim sidang
    timSidang.forEach(tim => {
        const pembimbing1 = tim['Pembimbing 1'];
        const pembimbing2 = tim['Pembimbing 2'];
        const penguji1 = tim['Penguji 1'];
        const penguji2 = tim['Penguji 2'];
        
        bebanDosen[pembimbing1] = (bebanDosen[pembimbing1] || 0) + 1;
        bebanDosen[pembimbing2] = (bebanDosen[pembimbing2] || 0) + 1;
        bebanDosen[penguji1] = (bebanDosen[penguji1] || 0) + 1;
        bebanDosen[penguji2] = (bebanDosen[penguji2] || 0) + 1;
    });
    
    return bebanDosen;
}

// PERBAIKAN: Fungsi untuk menghitung skor slot berdasarkan distribusi beban dosen
function calculateSlotScore(slot, jadwalHariTanggal, jadwalDosenTerpakai, dosenTim, bebanDosen) {
    const hariTanggalKey = slot.hariTanggalKey;
    const sesi = slot.sesi;
    const hari = jadwalHariTanggal[hariTanggalKey].hari;
    
    let skor = 0;
    
    // PERBAIKAN: Prioritaskan dosen dengan beban tinggi untuk diagendakan lebih awal
    let bebanTimSaatIni = 0;
    dosenTim.forEach(dosen => {
        bebanTimSaatIni += bebanDosen[dosen] || 0;
        
        // Bonus untuk dosen dengan beban tinggi
        if (bebanDosen[dosen] > 5) {
            skor += 3;
        }
    });
    
    // Bonus untuk tim dengan beban total tinggi
    skor += Math.min(10, bebanTimSaatIni / 4);
    
    // Prioritaskan slot yang lebih awal dalam jadwal (minggu pertama)
    const tanggal = new Date(jadwalHariTanggal[hariTanggalKey].date);
    const selisihHari = Math.floor((tanggal - tanggalMulaiSidang) / (24 * 60 * 60 * 1000));
    skor -= selisihHari * 0.2; // Penalti kecil untuk slot yang lebih jauh
    
    // Prioritaskan hari yang sudah memiliki sidang lain (konsolidasi)
    let hariSudahAdaSidang = false;
    let jumlahSidangHariIni = 0;
    
    Object.keys(jadwalHariTanggal).forEach(key => {
        if (jadwalHariTanggal[key].hari === hari) {
            for (let s = 1; s <= 4; s++) {
                if (jadwalHariTanggal[key].slots[s].jadwalSidang) {
                    hariSudahAdaSidang = true;
                    jumlahSidangHariIni++;
                }
            }
        }
    });
    
    if (hariSudahAdaSidang) {
        skor += 2; // Bonus untuk konsolidasi sidang pada hari yang sama
        
        // PERBAIKAN: Namun hindari hari yang sudah terlalu banyak sidang
        if (jumlahSidangHariIni >= 3) {
            skor -= jumlahSidangHariIni; // Penalti untuk menghindari overloading hari tertentu
        }
    }
    
    // PERBAIKAN: Prioritaskan distribusi beban yang lebih merata antar hari
    // Hitung beban jadwal terpakai saat ini untuk setiap dosen
    const dosenLoad = {};
    Object.values(daftarDosen).forEach(nama => {
        dosenLoad[nama] = 0;
    });
    
    // Tambahkan beban saat ini
    Object.keys(jadwalDosenTerpakai).forEach(dosen => {
        dosenLoad[dosen] = jadwalDosenTerpakai[dosen].length;
    });
    
    // Distribusi score: prioritaskan tim dengan dosen yang belum banyak terjadwal
    const bobotDistribusi = 1.5; // Bobot untuk faktor distribusi
    let totalBeban = 0;
    dosenTim.forEach(dosen => {
        totalBeban += dosenLoad[dosen] || 0;
    });
    
    // Nilai rata-rata beban
    const bebanRata = totalBeban / dosenTim.length;
    
    // Penalti untuk beban yang tidak merata
    dosenTim.forEach(dosen => {
        const beban = dosenLoad[dosen] || 0;
        const selisih = Math.abs(beban - bebanRata);
        skor -= selisih * bobotDistribusi;
    });
    
    // PERBAIKAN: Prioritaskan sesi tengah (sesi 2 dan 3) daripada sesi pagi pertama atau sore terakhir
    if (sesi === 2 || sesi === 3) {
        skor += 1;
    }
    
    return skor;
}

// PERBAIKAN: Implementasi algoritma genetik untuk optimasi jadwal
class GenetikPenjadwalan {
    constructor(timSidang, jadwalHariTanggal, jadwalMengajar) {
        this.timSidang = timSidang;
        this.jadwalHariTanggal = jadwalHariTanggal;
        this.jadwalMengajar = jadwalMengajar;
        this.ukuranPopulasi = 20;
        this.tingkatMutasi = 0.2;
        this.generasiMaksimum = 25;
        
        // Inisialisasi variabel-variabel
        this.populasi = [];
        this.bebanDosen = hitungBebanDosen(timSidang);
        
        // Cache untuk jadwal mengajar dosen
        this.dosenToJadwal = {};
        
        // Inisialisasi cache jadwal mengajar
        this.inisialisasiJadwalMengajar();
    }
    
    // Inisialisasi jadwal mengajar
    inisialisasiJadwalMengajar() {
        // Buat pemetaan nama lengkap dosen ke jadwal mengajar mereka
        this.jadwalMengajar.forEach(jadwal => {
            const kodeDosen = jadwal['Kode Dosen'];
            const namaLengkap = daftarDosen[kodeDosen];
            
            if (!this.dosenToJadwal[namaLengkap]) {
                this.dosenToJadwal[namaLengkap] = [];
            }
            
            // Simpan jadwal mengajar dengan hari-tanggal
            Object.keys(this.jadwalHariTanggal).forEach(key => {
                if (this.jadwalHariTanggal[key].hari === jadwal.Hari) {
                    const sesi = parseInt(jadwal.Sesi);
                    this.dosenToJadwal[namaLengkap].push({
                        hariTanggalKey: key,
                        sesi: sesi
                    });
                    
                    // Tandai slot ini sebagai tidak tersedia untuk jadwal sidang
                    this.jadwalHariTanggal[key].slots[sesi].tersedia = false;
                    this.jadwalHariTanggal[key].slots[sesi].jadwalMengajar.push({
                        dosen: namaLengkap,
                        mataKuliah: jadwal['Mata Kuliah'],
                        ruangan: jadwal['Ruangan']
                    });
                }
            });
        });
    }
    
    // Menghasilkan solusi awal secara acak
    generasiAwal() {
        this.populasi = [];
        
        // Buat beberapa solusi awal
        for (let i = 0; i < this.ukuranPopulasi; i++) {
            const solusi = this.buatSolusiAcak();
            this.populasi.push(solusi);
        }
    }
    
    // Membuat satu solusi acak
    buatSolusiAcak() {
        const solusi = [];
        const jadwalDosenTerpakai = {};
        
        // Tandai semua dosen mana yang sudah terjadwal pada slot mana
        const slotTerpakai = JSON.parse(JSON.stringify(this.jadwalHariTanggal));
        
        // Urutkan tim sidang: prioritaskan yang memiliki request jadwal
        const sortedTimSidang = [...this.timSidang].sort((a, b) => {
            const aHasRequest = a['Request Tanggal'] && a['Request Sesi'];
            const bHasRequest = b['Request Tanggal'] && b['Request Sesi'];
            
            if (aHasRequest && !bHasRequest) return -1;
            if (!aHasRequest && bHasRequest) return 1;
            return 0;
        });
        
        // Untuk setiap tim sidang
        for (let i = 0; i < sortedTimSidang.length; i++) {
            const tim = sortedTimSidang[i];
            const namaMahasiswa = tim['Nama Mahasiswa / NIM'];
            const pembimbing1 = tim['Pembimbing 1'];
            const pembimbing2 = tim['Pembimbing 2'];
            const penguji1 = tim['Penguji 1'];
            const penguji2 = tim['Penguji 2'];
            const requestTanggal = tim['Request Tanggal'];
            const requestSesi = tim['Request Sesi'] ? parseInt(tim['Request Sesi']) : null;
            
            const dosenTim = [pembimbing1, pembimbing2, penguji1, penguji2];
            
            // Coba penuhi request jika ada
            let slotDitemukan = false;
            
            if (requestTanggal && requestSesi) {
                // Cari key untuk hari-tanggal yang sesuai dengan request
                const requestDate = parseDate(requestTanggal);
                // getDay() mengembalikan 0-6, sesuaikan untuk indexToHari yang menggunakan 1-5
                const dayIndex = requestDate.getDay();
                const requestDay = indexToHari[dayIndex === 0 ? 1 : dayIndex]; // Fallback ke Senin jika Minggu
                // Format tanggal ke string DD/MM/YYYY untuk key
                const formattedRequestDate = formatDate(requestDate);
                const requestKey = `${requestDay}-${formattedRequestDate}`;
                
                // Cek apakah key tersebut ada dalam jadwalHariTanggal dan tersedia
                if (slotTerpakai[requestKey] && 
                    slotTerpakai[requestKey].slots[requestSesi] && 
                    slotTerpakai[requestKey].slots[requestSesi].tersedia) {
                    
                    // Cek apakah ada dosen tim yang mengajar pada slot ini
                    let adaYangMengajar = false;
                    for (const dosen of dosenTim) {
                        if (this.dosenToJadwal[dosen]) {
                            for (const jadwal of this.dosenToJadwal[dosen]) {
                                if (jadwal.hariTanggalKey === requestKey && jadwal.sesi === requestSesi) {
                                    adaYangMengajar = true;
                                    break;
                                }
                            }
                        }
                        if (adaYangMengajar) break;
                    }
                    
                    // Cek apakah ada dosen yang sudah terjadwal sidang pada slot ini
                    let adaYangSidang = false;
                    for (const dosen of dosenTim) {
                        if (jadwalDosenTerpakai[dosen]) {
                            for (const jadwal of jadwalDosenTerpakai[dosen]) {
                                if (jadwal.hariTanggalKey === requestKey && jadwal.sesi === requestSesi) {
                                    adaYangSidang = true;
                                    break;
                                }
                            }
                        }
                        if (adaYangSidang) break;
                    }
                    
                    // Jika tidak ada konflik, gunakan slot request
                    if (!adaYangMengajar && !adaYangSidang) {
                        // Tandai slot ini sebagai digunakan
                        slotTerpakai[requestKey].slots[requestSesi].tersedia = false;
                        
                        // Update jadwal dosen terpakai
                        dosenTim.forEach(dosen => {
                            if (!jadwalDosenTerpakai[dosen]) {
                                jadwalDosenTerpakai[dosen] = [];
                            }
                            jadwalDosenTerpakai[dosen].push({
                                hariTanggalKey: requestKey,
                                sesi: requestSesi
                            });
                        });
                        
                        // Tambahkan ke solusi
                        solusi.push({
                            mahasiswaIndex: i,
                            hariTanggalKey: requestKey,
                            sesi: requestSesi,
                            isRequest: true
                        });
                        
                        slotDitemukan = true;
                    }
                }
            }
            
            // Jika tidak bisa memenuhi request atau tidak ada request, cari slot lain secara acak
            if (!slotDitemukan) {
                // Daftar semua slot yang masih tersedia
                const slotTersedia = [];
                
                Object.keys(slotTerpakai).forEach(hariTanggalKey => {
                    for (let sesi = 1; sesi <= 4; sesi++) {
                        if (slotTerpakai[hariTanggalKey].slots[sesi].tersedia) {
                            // Cek apakah ada dosen yang mengajar pada slot ini
                            let adaYangMengajar = false;
                            for (const dosen of dosenTim) {
                                if (this.dosenToJadwal[dosen]) {
                                    for (const jadwal of this.dosenToJadwal[dosen]) {
                                        if (jadwal.hariTanggalKey === hariTanggalKey && jadwal.sesi === sesi) {
                                            adaYangMengajar = true;
                                            break;
                                        }
                                    }
                                }
                                if (adaYangMengajar) break;
                            }
                            
                            // Cek apakah ada dosen yang sudah terjadwal sidang pada slot ini
                            let adaYangSidang = false;
                            for (const dosen of dosenTim) {
                                if (jadwalDosenTerpakai[dosen]) {
                                    for (const jadwal of jadwalDosenTerpakai[dosen]) {
                                        if (jadwal.hariTanggalKey === hariTanggalKey && jadwal.sesi === sesi) {
                                            adaYangSidang = true;
                                            break;
                                        }
                                    }
                                }
                                if (adaYangSidang) break;
                            }
                            
                            // Jika tidak ada konflik, tambahkan ke slot tersedia
                            if (!adaYangMengajar && !adaYangSidang) {
                                slotTersedia.push({
                                    hariTanggalKey: hariTanggalKey,
                                    sesi: sesi
                                });
                            }
                        }
                    }
                });
                
                // Jika ada slot tersedia, pilih secara acak
                if (slotTersedia.length > 0) {
                    // Hitung skor untuk setiap slot tersedia
                    slotTersedia.forEach(slot => {
                        slot.score = calculateSlotScore(slot, this.jadwalHariTanggal, jadwalDosenTerpakai, dosenTim, this.bebanDosen);
                    });
                    
                    // Urutkan berdasarkan skor tertinggi
                    slotTersedia.sort((a, b) => b.score - a.score);
                    
                    // Pilih salah satu dari yang terbaik (sedikit keacakan)
                    const topN = Math.min(3, slotTersedia.length);
                    const randomIndex = Math.floor(Math.random() * topN);
                    const slotTerpilih = slotTersedia[randomIndex];
                    
                    const hariTanggalKey = slotTerpilih.hariTanggalKey;
                    const sesi = slotTerpilih.sesi;
                    
                    // Tandai slot ini sebagai digunakan
                    slotTerpakai[hariTanggalKey].slots[sesi].tersedia = false;
                    
                    // Update jadwal dosen terpakai
                    dosenTim.forEach(dosen => {
                        if (!jadwalDosenTerpakai[dosen]) {
                            jadwalDosenTerpakai[dosen] = [];
                        }
                        jadwalDosenTerpakai[dosen].push({
                            hariTanggalKey: hariTanggalKey,
                            sesi: sesi
                        });
                    });
                    
                    // Tambahkan ke solusi
                    solusi.push({
                        mahasiswaIndex: i,
                        hariTanggalKey: hariTanggalKey,
                        sesi: sesi,
                        isRequest: false
                    });
                } else {
                    // Jika tidak ada slot tersedia, tandai sebagai tidak terjadwal
                    solusi.push({
                        mahasiswaIndex: i,
                        hariTanggalKey: null,
                        sesi: null,
                        isRequest: false
                    });
                }
            }
        }
        
        return solusi;
    }
    
    // Evaluasi fitness solusi
    evaluasiSolusi(solusi) {
        let nilaiKualitas = 0;
        
        // Variabel untuk tracking jadwal per dosen dan per hari
        const jadwalDosenTerpakai = {};
        const jadwalPerHari = {};
        const slotTerpakai = {};
        
        // Inisialisasi tracking
        Object.values(daftarDosen).forEach(nama => {
            jadwalDosenTerpakai[nama] = [];
        });
        
        Object.keys(this.jadwalHariTanggal).forEach(key => {
            jadwalPerHari[key] = 0;
            slotTerpakai[key] = {};
            for (let sesi = 1; sesi <= 4; sesi++) {
                slotTerpakai[key][sesi] = false;
            }
        });
        
        // Untuk setiap mahasiswa dalam solusi
        for (let i = 0; i < solusi.length; i++) {
            const penugasan = solusi[i];
            
            // Jika ada jadwal, evaluasi
            if (penugasan.hariTanggalKey && penugasan.sesi) {
                // Tambah nilai untuk mahasiswa yang terjadwal
                nilaiKualitas += 100;
                
                // Tambah nilai ekstra jika memenuhi request
                if (penugasan.isRequest) {
                    nilaiKualitas += 50;
                }
                
                // Kurangi nilai jika terjadi bentrok dengan jadwal mengajar
                const tim = this.timSidang[penugasan.mahasiswaIndex];
                const dosenTim = [tim['Pembimbing 1'], tim['Pembimbing 2'], tim['Penguji 1'], tim['Penguji 2']];
                
                // Cek bentrok dengan jadwal mengajar
                for (const dosen of dosenTim) {
                    if (this.dosenToJadwal[dosen]) {
                        for (const jadwal of this.dosenToJadwal[dosen]) {
                            if (jadwal.hariTanggalKey === penugasan.hariTanggalKey && jadwal.sesi === penugasan.sesi) {
                                nilaiKualitas -= 500; // Penalti besar untuk bentrok dengan jadwal mengajar
                            }
                        }
                    }
                }
                
                // Cek bentrok dengan jadwal sidang lain
                if (slotTerpakai[penugasan.hariTanggalKey][penugasan.sesi]) {
                    nilaiKualitas -= 500; // Penalti besar untuk bentrok dengan sidang lain
                }
                
                // Tandai slot ini sebagai terpakai
                slotTerpakai[penugasan.hariTanggalKey][penugasan.sesi] = true;
                
                // Update jadwal per dosen
                for (const dosen of dosenTim) {
                    jadwalDosenTerpakai[dosen].push({
                        hariTanggalKey: penugasan.hariTanggalKey,
                        sesi: penugasan.sesi
                    });
                }
                
                // Update jadwal per hari
                jadwalPerHari[penugasan.hariTanggalKey]++;
                
                // Cek beban dosen (penalti untuk beban yang tidak merata)
                for (const dosen of Object.keys(jadwalDosenTerpakai)) {
                    const jumlahJadwal = jadwalDosenTerpakai[dosen].length;
                    // Reward untuk distribusi yang merata
                    if (jumlahJadwal > 0) {
                        nilaiKualitas -= Math.abs(jumlahJadwal - (this.bebanDosen[dosen] || 0) / 2) * 2;
                    }
                }
                
                // Reward untuk konsolidasi jadwal (beberapa sidang dalam sehari)
                for (const hari of Object.keys(jadwalPerHari)) {
                    if (jadwalPerHari[hari] > 1) {
                        nilaiKualitas += Math.min(jadwalPerHari[hari], 3) * 2; // Max 3 sidang per hari
                    }
                    // Penalti jika terlalu banyak sidang dalam sehari
                    if (jadwalPerHari[hari] > 3) {
                        nilaiKualitas -= (jadwalPerHari[hari] - 3) * 5;
                    }
                }
                
                // Tambahan: preferensi untuk sesi tengah (2 dan 3)
                if (penugasan.sesi === 2 || penugasan.sesi === 3) {
                    nilaiKualitas += 1;
                }
            } else {
                // Penalti untuk mahasiswa yang tidak terjadwal
                nilaiKualitas -= 200;
            }
        }
        
        // Jika nilai kualitas terlalu negatif, set minimal 1
        return Math.max(1, nilaiKualitas);
    }
    
    // Seleksi orang tua menggunakan metode roulette wheel
    seleksiOrangTua() {
        // Hitung total fitness
        let totalFitness = 0;
        const fitness = [];
        
        for (let i = 0; i < this.populasi.length; i++) {
            const nilai = this.evaluasiSolusi(this.populasi[i]);
            fitness.push(nilai);
            totalFitness += nilai;
        }
        
        // Pilih orang tua berdasarkan probabilitas fitness
        const randomValue = Math.random() * totalFitness;
        let sum = 0;
        
        for (let i = 0; i < this.populasi.length; i++) {
            sum += fitness[i];
            if (sum >= randomValue) {
                return this.populasi[i];
            }
        }
        
        // Jika ada kesalahan, kembalikan individu terakhir
        return this.populasi[this.populasi.length - 1];
    }
    
    // Crossover dua parent menjadi dua child
    crossover(parent1, parent2) {
        // Pilih titik crossover secara acak
        const crossoverPoint = Math.floor(Math.random() * parent1.length);
        
        // Buat dua anak
        const child1 = [...parent1.slice(0, crossoverPoint), ...parent2.slice(crossoverPoint)];
        const child2 = [...parent2.slice(0, crossoverPoint), ...parent1.slice(crossoverPoint)];
        
        // Perbaiki konflik jika ada
        this.perbaikiKonflik(child1);
        this.perbaikiKonflik(child2);
        
        return [child1, child2];
    }
    
    // Perbaiki konflik dalam solusi (misalnya jika ada slot yang digunakan lebih dari sekali)
    perbaikiKonflik(solusi) {
        const slotTerpakai = {};
        const konflik = [];
        
        // Tandai slot yang terpakai
        for (let i = 0; i < solusi.length; i++) {
            const penugasan = solusi[i];
            
            // Lewati yang tidak terjadwal
            if (!penugasan.hariTanggalKey || !penugasan.sesi) {
                continue;
            }
            
            const key = `${penugasan.hariTanggalKey}-${penugasan.sesi}`;
            
            if (!slotTerpakai[key]) {
                slotTerpakai[key] = [i];
            } else {
                slotTerpakai[key].push(i);
                konflik.push(key);
            }
        }
        
        // Perbaiki konflik dengan mencari slot baru
        for (const key of konflik) {
            const indeks = slotTerpakai[key];
            
            // Pertahankan yang pertama, cari slot baru untuk yang lain
            for (let i = 1; i < indeks.length; i++) {
                const mahasiswaIndex = indeks[i];
                
                // Reset jadwal mahasiswa ini
                solusi[mahasiswaIndex].hariTanggalKey = null;
                solusi[mahasiswaIndex].sesi = null;
                solusi[mahasiswaIndex].isRequest = false;
                
                // Coba cari slot baru secara acak
                // Fungsi ini sederhana; untuk implementasi lengkap, perlu algoritma pencarian slot yang lebih kompleks
                this.cariSlotBaru(solusi, mahasiswaIndex);
            }
        }
    }
    
    // Mencari slot baru untuk mahasiswa tertentu
    cariSlotBaru(solusi, mahasiswaIndex) {
        const tim = this.timSidang[solusi[mahasiswaIndex].mahasiswaIndex];
        const dosenTim = [tim['Pembimbing 1'], tim['Pembimbing 2'], tim['Penguji 1'], tim['Penguji 2']];
        
        // Daftar slot yang sudah terpakai
        const slotTerpakai = {};
        
        // Tandai slot yang sudah digunakan dalam solusi saat ini
        for (const penugasan of solusi) {
            if (penugasan.hariTanggalKey && penugasan.sesi) {
                const key = `${penugasan.hariTanggalKey}-${penugasan.sesi}`;
                slotTerpakai[key] = true;
            }
        }
        
        // Cari semua slot yang masih tersedia
        const slotTersedia = [];
        
        Object.keys(this.jadwalHariTanggal).forEach(hariTanggalKey => {
            for (let sesi = 1; sesi <= 4; sesi++) {
                const key = `${hariTanggalKey}-${sesi}`;
                
                // Lewati jika slot sudah terpakai
                if (slotTerpakai[key] || !this.jadwalHariTanggal[hariTanggalKey].slots[sesi].tersedia) {
                    continue;
                }
                
                // Cek apakah ada dosen yang mengajar pada slot ini
                let adaYangMengajar = false;
                for (const dosen of dosenTim) {
                    if (this.dosenToJadwal[dosen]) {
                        for (const jadwal of this.dosenToJadwal[dosen]) {
                            if (jadwal.hariTanggalKey === hariTanggalKey && jadwal.sesi === sesi) {
                                adaYangMengajar = true;
                                break;
                            }
                        }
                    }
                    if (adaYangMengajar) break;
                }
                
                // Jika tidak ada konflik dengan jadwal mengajar, tambahkan
                if (!adaYangMengajar) {
                    slotTersedia.push({
                        hariTanggalKey: hariTanggalKey,
                        sesi: sesi
                    });
                }
            }
        });
        
        // Jika ada slot tersedia, pilih secara acak
        if (slotTersedia.length > 0) {
            const randomIndex = Math.floor(Math.random() * slotTersedia.length);
            const slotTerpilih = slotTersedia[randomIndex];
            
            solusi[mahasiswaIndex].hariTanggalKey = slotTerpilih.hariTanggalKey;
            solusi[mahasiswaIndex].sesi = slotTerpilih.sesi;
            solusi[mahasiswaIndex].isRequest = false;
        }
        // Jika tidak ada slot tersedia, biarkan null (tidak terjadwal)
    }
    
    // Mutasi solusi
    mutasi(solusi) {
        // Untuk beberapa mahasiswa secara acak
        for (let i = 0; i < solusi.length; i++) {
            // Lakukan mutasi dengan probabilitas tertentu
            if (Math.random() < this.tingkatMutasi) {
                // Jika sebelumnya tidak terjadwal, coba jadwalkan
                if (!solusi[i].hariTanggalKey || !solusi[i].sesi) {
                    this.cariSlotBaru(solusi, i);
                } else {
                    // Jika sudah terjadwal, cari slot baru dengan probabilitas tertentu
                    // Jika request, pertahankan dengan probabilitas lebih tinggi
                    if (solusi[i].isRequest) {
                        if (Math.random() < 0.2) { // Lebih kecil probabilitas mengubah request
                            solusi[i].hariTanggalKey = null;
                            solusi[i].sesi = null;
                            solusi[i].isRequest = false;
                            this.cariSlotBaru(solusi, i);
                        }
                    } else {
                        solusi[i].hariTanggalKey = null;
                        solusi[i].sesi = null;
                        this.cariSlotBaru(solusi, i);
                    }
                }
            }
        }
    }
    
    // Evolusi populasi untuk generasi berikutnya
    evolusi() {
        // Evaluasi populasi saat ini
        const evaluasi = this.populasi.map(solusi => ({
            solusi,
            fitness: this.evaluasiSolusi(solusi)
        }));
        
        // Urutkan berdasarkan fitness
        evaluasi.sort((a, b) => b.fitness - a.fitness);
        
        // Pertahankan beberapa solusi terbaik (elitism)
        const newPopulasi = evaluasi.slice(0, 2).map(e => e.solusi);
        
        // Buat populasi baru melalui crossover dan mutasi
        while (newPopulasi.length < this.ukuranPopulasi) {
            // Pilih dua orangtua
            const parent1 = this.seleksiOrangTua();
            const parent2 = this.seleksiOrangTua();
            
            // Lakukan crossover
            const [child1, child2] = this.crossover(parent1, parent2);
            
            // Lakukan mutasi
            this.mutasi(child1);
            this.mutasi(child2);
            
            // Tambahkan ke populasi baru
            newPopulasi.push(child1);
            if (newPopulasi.length < this.ukuranPopulasi) {
                newPopulasi.push(child2);
            }
        }
        
        // Update populasi
        this.populasi = newPopulasi;
    }
    
    // Mendapatkan solusi terbaik saat ini
    getSolusiTerbaik() {
        let terbaik = null;
        let fitnessTerbaik = -Infinity;
        
        for (const solusi of this.populasi) {
            const fitness = this.evaluasiSolusi(solusi);
            if (fitness > fitnessTerbaik) {
                fitnessTerbaik = fitness;
                terbaik = solusi;
            }
        }
        
        return terbaik;
    }
    
    // Mencari solusi dengan algoritma genetik
    cariSolusi() {
        // Inisialisasi populasi awal
        this.generasiAwal();
        
        // Evolusi untuk beberapa generasi
        for (let i = 0; i < this.generasiMaksimum; i++) {
            this.evolusi();
            
            // Tampilkan pesan debug
            if (i % 5 === 0) {
                const terbaik = this.getSolusiTerbaik();
                const fitness = this.evaluasiSolusi(terbaik);
                const terjadwal = terbaik.filter(p => p.hariTanggalKey !== null).length;
                
                console.log(`Generasi ${i}: Fitness terbaik = ${fitness}, Terjadwal = ${terjadwal}/${terbaik.length}`);
            }
        }
        
        // Dapatkan solusi terbaik
        return this.getSolusiTerbaik();
    }
    
    // Konversi solusi ke format hasil jadwal
    konversiKeHasilJadwal(solusi) {
        const hasil = [];
        
        for (let i = 0; i < solusi.length; i++) {
            const penugasan = solusi[i];
            const tim = this.timSidang[penugasan.mahasiswaIndex];
            
            if (penugasan.hariTanggalKey && penugasan.sesi) {
                // Mahasiswa terjadwal
                hasil.push({
                    'Nama Mahasiswa': tim['Nama Mahasiswa / NIM'],
                    'Pembimbing 1': tim['Pembimbing 1'],
                    'Pembimbing 2': tim['Pembimbing 2'],
                    'Penguji 1': tim['Penguji 1'],
                    'Penguji 2': tim['Penguji 2'],
                    'Hari': this.jadwalHariTanggal[penugasan.hariTanggalKey].hari,
                    'Tanggal': this.jadwalHariTanggal[penugasan.hariTanggalKey].tanggal,
                    'Sesi': penugasan.sesi,
                    'Jam': jadwalJam[penugasan.sesi],
                    'IsRequest': penugasan.isRequest
                });
                
                // Tandai slot sebagai terpakai
                this.jadwalHariTanggal[penugasan.hariTanggalKey].slots[penugasan.sesi].tersedia = false;
                this.jadwalHariTanggal[penugasan.hariTanggalKey].slots[penugasan.sesi].jadwalSidang = {
                    mahasiswa: tim['Nama Mahasiswa / NIM'],
                    dosen: [tim['Pembimbing 1'], tim['Pembimbing 2'], tim['Penguji 1'], tim['Penguji 2']],
                    isRequest: penugasan.isRequest
                };
            } else {
                // Mahasiswa tidak terjadwal
                hasil.push({
                    'Nama Mahasiswa': tim['Nama Mahasiswa / NIM'],
                    'Pembimbing 1': tim['Pembimbing 1'],
                    'Pembimbing 2': tim['Pembimbing 2'],
                    'Penguji 1': tim['Penguji 1'],
                    'Penguji 2': tim['Penguji 2'],
                    'Hari': 'Tidak ada slot tersedia',
                    'Tanggal': '-',
                    'Sesi': '-',
                    'Jam': '-',
                    'IsRequest': false
                });
            }
        }
        
        return hasil;
    }
}

// PERBAIKAN: Algoritma Penjadwalan yang Lebih Baik
function scheduleTA(jadwalMengajar, timSidang, jadwalHariTanggal) {
    console.log(`Mulai penjadwalan dengan ${timSidang.length} mahasiswa dan ${Object.keys(daftarDosen).length} dosen`);
    
    // PERBAIKAN: Menyimpan beban awal tiap dosen dalam tim sidang
    const bebanAwalDosen = {};
    Object.values(daftarDosen).forEach(nama => {
        bebanAwalDosen[nama] = 0;
    });
    
    timSidang.forEach(tim => {
        const pembimbing1 = tim['Pembimbing 1'];
        const pembimbing2 = tim['Pembimbing 2'];
        const penguji1 = tim['Penguji 1'];
        const penguji2 = tim['Penguji 2'];
        
        bebanAwalDosen[pembimbing1] = (bebanAwalDosen[pembimbing1] || 0) + 1;
        bebanAwalDosen[pembimbing2] = (bebanAwalDosen[pembimbing2] || 0) + 1;
        bebanAwalDosen[penguji1] = (bebanAwalDosen[penguji1] || 0) + 1;
        bebanAwalDosen[penguji2] = (bebanAwalDosen[penguji2] || 0) + 1;
    });
    
    // Tampilkan beban awal dosen
    console.log("Beban awal dosen (jumlah keterlibatan dalam tim):");
    Object.entries(bebanAwalDosen)
        .sort((a, b) => b[1] - a[1])
        .forEach(([dosen, beban]) => {
            console.log(`- ${dosen}: ${beban} tim`);
        });
    
    // PERBAIKAN: Menyiapkan mapping jadwal mengajar setiap dosen
    const dosenToJadwal = {};
    Object.values(daftarDosen).forEach(nama => {
        dosenToJadwal[nama] = [];
    });
    
    // Tandai jadwal mengajar dosen
    jadwalMengajar.forEach(jadwal => {
        const kodeDosen = jadwal['Kode Dosen'];
        const namaLengkap = daftarDosen[kodeDosen];
        
        if (!dosenToJadwal[namaLengkap]) {
            dosenToJadwal[namaLengkap] = [];
        }
        
        Object.keys(jadwalHariTanggal).forEach(key => {
            if (jadwalHariTanggal[key].hari === jadwal.Hari) {
                const sesi = parseInt(jadwal.Sesi);
                dosenToJadwal[namaLengkap].push({
                    hariTanggalKey: key,
                    sesi: sesi
                });
                
                // Tandai slot ini sebagai tidak tersedia untuk jadwal sidang
                jadwalHariTanggal[key].slots[sesi].tersedia = false;
                jadwalHariTanggal[key].slots[sesi].jadwalMengajar.push({
                    dosen: namaLengkap,
                    mataKuliah: jadwal['Mata Kuliah'],
                    ruangan: jadwal['Ruangan']
                });
            }
        });
    });
    
    // PERBAIKAN: Lebih banyak variabel debugging
    let totalRequestDipenuhi = 0;
    let totalRequest = 0;

    // PERBAIKAN: Urutkan mahasiswa berdasarkan constraint terbanyak
    // (1. Mahasiswa dengan request, 2. Mahasiswa dengan dosen berbeban tinggi)
    const sortedTimSidang = timSidang.map((tim, index) => {
        const hasRequest = tim['Request Tanggal'] && tim['Request Sesi'];
        if (hasRequest) totalRequest++;
        
        // Hitung jumlah beban total dosen dalam tim
        const dosenTim = [tim['Pembimbing 1'], tim['Pembimbing 2'], tim['Penguji 1'], tim['Penguji 2']];
        const totalBebanDosen = dosenTim.reduce((sum, dosen) => sum + (bebanAwalDosen[dosen] || 0), 0);
        
        return {
            tim: tim,
            index: index,
            hasRequest: hasRequest,
            totalBebanDosen: totalBebanDosen
        };
    }).sort((a, b) => {
        // Prioritaskan request terlebih dahulu
        if (a.hasRequest && !b.hasRequest) return -1;
        if (!a.hasRequest && b.hasRequest) return 1;
        
        // Jika keduanya request atau keduanya tidak request, prioritaskan tim dengan dosen berbeban tinggi
        return b.totalBebanDosen - a.totalBebanDosen;
    });
    
    console.log(`Total mahasiswa dengan request: ${totalRequest} dari ${timSidang.length}`);
    
    // Hasil jadwal
    const hasilJadwal = new Array(timSidang.length);
    
    // Jadwal dosen terpakai untuk sidang
    const jadwalDosenTerpakai = {};
    Object.values(daftarDosen).forEach(nama => {
        jadwalDosenTerpakai[nama] = [];
    });
    
    // PERBAIKAN: Tambahkan variable untuk membantu backtracking
    let assignedCount = 0;
    const maxRetries = 3; // Jumlah maksimum percobaan re-scheduling per mahasiswa
    
    // PERBAIKAN: Implementasi CSP dengan backtracking
    // Fungsi rekursif untuk assign tim sidang ke slot
    function assignTimToSlot(timIndex) {
        // Basis kasus: semua tim telah dijadwalkan
        if (timIndex >= sortedTimSidang.length) {
            return true;
        }
        
        const timData = sortedTimSidang[timIndex];
        const tim = timData.tim;
        const originalIndex = timData.index;
        
        const namaMahasiswa = tim['Nama Mahasiswa / NIM'];
        const pembimbing1 = tim['Pembimbing 1'];
        const pembimbing2 = tim['Pembimbing 2'];
        const penguji1 = tim['Penguji 1'];
        const penguji2 = tim['Penguji 2'];
        const requestTanggal = tim['Request Tanggal'];
        const requestSesi = tim['Request Sesi'] ? parseInt(tim['Request Sesi']) : null;
        
        const dosenTim = [pembimbing1, pembimbing2, penguji1, penguji2];
        
        // Cek apakah ada request jadwal
        let slotDitemukan = false;
        let isRequest = false;
        
        // PERBAIKAN: Array untuk menyimpan semua slot yang bisa digunakan
        let candidateSlots = [];
        
        if (requestTanggal && requestSesi) {
            // Cari key untuk hari-tanggal yang sesuai dengan request
            const requestDate = parseDate(requestTanggal);
            const dayIndex = requestDate.getDay();
            const requestDay = indexToHari[dayIndex === 0 ? 1 : (dayIndex === 6 ? 5 : dayIndex)]; // Fallback sesuai
            const formattedRequestDate = formatDate(requestDate);
            const requestKey = `${requestDay}-${formattedRequestDate}`;
            
            // Cek apakah slot request tersedia
            if (jadwalHariTanggal[requestKey] && 
                jadwalHariTanggal[requestKey].slots[requestSesi] && 
                jadwalHariTanggal[requestKey].slots[requestSesi].tersedia) {
                
                // Cek apakah ada dosen yang mengajar pada slot ini
                let adaYangMengajar = false;
                for (const dosen of dosenTim) {
                    if (dosenToJadwal[dosen]) {
                        for (const jadwal of dosenToJadwal[dosen]) {
                            if (jadwal.hariTanggalKey === requestKey && jadwal.sesi === requestSesi) {
                                adaYangMengajar = true;
                                break;
                            }
                        }
                    }
                    if (adaYangMengajar) break;
                }
                
                // Cek apakah ada dosen yang sudah terjadwal sidang pada slot ini
                let adaYangSidang = false;
                for (const dosen of dosenTim) {
                    if (jadwalDosenTerpakai[dosen]) {
                        for (const jadwal of jadwalDosenTerpakai[dosen]) {
                            if (jadwal.hariTanggalKey === requestKey && jadwal.sesi === requestSesi) {
                                adaYangSidang = true;
                                break;
                            }
                        }
                    }
                    if (adaYangSidang) break;
                }
                
                // Jika tidak ada konflik, gunakan slot request
                if (!adaYangMengajar && !adaYangSidang) {
                    // Tambahkan ke kandidat dengan prioritas tertinggi
                    candidateSlots.push({
                        hariTanggalKey: requestKey,
                        sesi: requestSesi,
                        isRequest: true,
                        score: 1000 // Skor tinggi untuk request
                    });
                    
                    // PERBAIKAN: Pesan debug untuk request yang dipenuhi
                    console.log(`Request ${namaMahasiswa} untuk tanggal ${requestTanggal} sesi ${requestSesi} mungkin bisa dipenuhi`);
                }
            } else {
                // PERBAIKAN: Pesan debug untuk request yang tidak tersedia
                console.log(`Request ${namaMahasiswa} untuk tanggal ${requestTanggal} sesi ${requestSesi} TIDAK TERSEDIA: ${!jadwalHariTanggal[requestKey] ? 'Tanggal tidak ada' : 'Slot tidak tersedia'}`);
            }
        }
        
        // Jika tidak ada request atau request tidak bisa dipenuhi, cari slot tersedia
        // PERBAIKAN: Temukan semua slot tersedia
        Object.keys(jadwalHariTanggal).forEach(hariTanggalKey => {
            for (let sesi = 1; sesi <= 4; sesi++) {
                if (jadwalHariTanggal[hariTanggalKey].slots[sesi].tersedia) {
                    // Cek konflik dengan jadwal mengajar
                    let adaYangMengajar = false;
                    for (const dosen of dosenTim) {
                        if (dosenToJadwal[dosen]) {
                            for (const jadwal of dosenToJadwal[dosen]) {
                                if (jadwal.hariTanggalKey === hariTanggalKey && jadwal.sesi === sesi) {
                                    adaYangMengajar = true;
                                    break;
                                }
                            }
                        }
                        if (adaYangMengajar) break;
                    }
                    
                    // Cek konflik dengan jadwal sidang lain
                    let adaYangSidang = false;
                    for (const dosen of dosenTim) {
                        if (jadwalDosenTerpakai[dosen]) {
                            for (const jadwal of jadwalDosenTerpakai[dosen]) {
                                if (jadwal.hariTanggalKey === hariTanggalKey && jadwal.sesi === sesi) {
                                    adaYangSidang = true;
                                    break;
                                }
                            }
                        }
                        if (adaYangSidang) break;
                    }
                    
                    // Jika tidak ada konflik, tambahkan ke kandidat
                    if (!adaYangMengajar && !adaYangSidang) {
                        // Hitung skor untuk slot ini
                        let score = 0;
                        
                        // PERBAIKAN: Lebih detail dalam menghitung skor
                        // Preferensi untuk hari kerja awal (Senin-Selasa)
                        const hari = jadwalHariTanggal[hariTanggalKey].hari;
                        if (hari === 'Senin' || hari === 'Selasa') {
                            score += 5;
                        } else if (hari === 'Rabu') {
                            score += 3; 
                        }
                        
                        // Preferensi untuk sesi tengah (2 dan 3)
                        if (sesi === 2 || sesi === 3) {
                            score += 3;
                        }
                        
                        // Preferensi untuk tanggal yang lebih awal
                        const tanggal = new Date(jadwalHariTanggal[hariTanggalKey].date);
                        const selisihHari = Math.floor((tanggal - new Date(jadwalHariTanggal[Object.keys(jadwalHariTanggal)[0]].date)) / (24 * 60 * 60 * 1000));
                        score -= selisihHari;
                        
                        // Tambahkan ke kandidat
                        candidateSlots.push({
                            hariTanggalKey: hariTanggalKey,
                            sesi: sesi,
                            isRequest: false,
                            score: score
                        });
                    }
                }
            }
        });
        
        // Urutkan kandidat berdasarkan skor
        candidateSlots.sort((a, b) => b.score - a.score);
        
        // PERBAIKAN: Debug jumlah slot kandidat
        console.log(`${namaMahasiswa}: ${candidateSlots.length} slot kandidat ditemukan`);
        
        // PERBAIKAN: Implementasi backtracking
        // Coba setiap kandidat slot secara berurutan
        for (let retryCount = 0; retryCount < Math.min(maxRetries, candidateSlots.length); retryCount++) {
            const slot = candidateSlots[retryCount];
            
            if (!slot) continue; // Skip jika tidak ada slot kandidat lagi
            
            const hariTanggalKey = slot.hariTanggalKey;
            const sesi = slot.sesi;
            const isRequest = slot.isRequest;
            
            // Tandai slot ini sebagai digunakan
            jadwalHariTanggal[hariTanggalKey].slots[sesi].tersedia = false;
            jadwalHariTanggal[hariTanggalKey].slots[sesi].jadwalSidang = {
                mahasiswa: namaMahasiswa,
                dosen: dosenTim,
                isRequest: isRequest
            };
            
            // Update jadwal dosen terpakai
            dosenTim.forEach(dosen => {
                jadwalDosenTerpakai[dosen].push({
                    hariTanggalKey: hariTanggalKey,
                    sesi: sesi
                });
            });
            
            // Tambahkan ke hasil
            hasilJadwal[originalIndex] = {
                'Nama Mahasiswa': namaMahasiswa,
                'Pembimbing 1': pembimbing1,
                'Pembimbing 2': pembimbing2,
                'Penguji 1': penguji1,
                'Penguji 2': penguji2,
                'Hari': jadwalHariTanggal[hariTanggalKey].hari,
                'Tanggal': jadwalHariTanggal[hariTanggalKey].tanggal,
                'Sesi': sesi,
                'Jam': jadwalJam[sesi],
                'IsRequest': isRequest
            };
            
            if (isRequest) {
                totalRequestDipenuhi++;
            }
            
            assignedCount++;
            
            // Mencoba melanjutkan ke tim berikutnya
            if (assignTimToSlot(timIndex + 1)) {
                return true; // Solusi ditemukan!
            }
            
            // Jika tidak berhasil, backtrack: batalkan penugasan saat ini
            // dan coba slot berikutnya
            assignedCount--;
            if (isRequest) {
                totalRequestDipenuhi--;
            }
            
            // Hapus dari jadwal dosen terpakai
            dosenTim.forEach(dosen => {
                const idx = jadwalDosenTerpakai[dosen].findIndex(j => 
                    j.hariTanggalKey === hariTanggalKey && j.sesi === sesi);
                if (idx !== -1) {
                    jadwalDosenTerpakai[dosen].splice(idx, 1);
                }
            });
            
            // Tandai slot sebagai tersedia lagi
            jadwalHariTanggal[hariTanggalKey].slots[sesi].tersedia = true;
            jadwalHariTanggal[hariTanggalKey].slots[sesi].jadwalSidang = null;
            
            // PERBAIKAN: Pesan debug backtracking
            console.log(`Backtracking dari ${namaMahasiswa} (coba slot alternatif #${retryCount + 1})`);
        }
        
        // Jika semua slot gagal, buat hasil "Tidak ada slot tersedia"
        hasilJadwal[originalIndex] = {
            'Nama Mahasiswa': namaMahasiswa,
            'Pembimbing 1': pembimbing1,
            'Pembimbing 2': pembimbing2,
            'Penguji 1': penguji1,
            'Penguji 2': penguji2,
            'Hari': 'Tidak ada slot tersedia',
            'Tanggal': '-',
            'Sesi': '-',
            'Jam': '-',
            'IsRequest': false
        };
        
        // Lanjutkan ke tim berikutnya meskipun yang ini gagal
        return assignTimToSlot(timIndex + 1);
    }
    
    // PERBAIKAN: Coba pendekatan kombinatorial dengan relaksasi bertahap
    console.time('Waktu Penjadwalan');
    
    // Pertama: Coba menjadwalkan semua tim
    console.log("Tahap 1: Mencoba menjadwalkan semua mahasiswa");
    let result = assignTimToSlot(0);
    
    // PERBAIKAN: Jika belum optimal (< 80% terjadwal), coba relaksasi dan pendekatan lain
    const threshold = Math.floor(timSidang.length * 0.8); // 80% mahasiswa terjadwal
    
    if (assignedCount < threshold) {
        console.log(`Penjadwalan kurang optimal: ${assignedCount}/${timSidang.length}. Mencoba pendekatan kedua...`);
        
        // Reset semua jadwal sidang
        Object.keys(jadwalHariTanggal).forEach(key => {
            for (let sesi = 1; sesi <= 4; sesi++) {
                if (jadwalHariTanggal[key].slots[sesi].jadwalSidang) {
                    jadwalHariTanggal[key].slots[sesi].tersedia = true;
                    jadwalHariTanggal[key].slots[sesi].jadwalSidang = null;
                }
            }
        });
        
        // Reset jadwal dosen terpakai
        Object.keys(jadwalDosenTerpakai).forEach(dosen => {
            jadwalDosenTerpakai[dosen] = [];
        });
        
        assignedCount = 0;
        totalRequestDipenuhi = 0;
        
        // PENDEKATAN ALTERNATIF: Greedy + Constraint Relaxation
        console.log("Tahap 2: Pendekatan greedy dengan relaksasi constraint");
        
        // Implementasi pendekatan greedy sederhana (tanpa backtracking)
        for (const timData of sortedTimSidang) {
            const tim = timData.tim;
            const originalIndex = timData.index;
            
            const namaMahasiswa = tim['Nama Mahasiswa / NIM'];
            const pembimbing1 = tim['Pembimbing 1'];
            const pembimbing2 = tim['Pembimbing 2'];
            const penguji1 = tim['Penguji 1'];
            const penguji2 = tim['Penguji 2'];
            const requestTanggal = tim['Request Tanggal'];
            const requestSesi = tim['Request Sesi'] ? parseInt(tim['Request Sesi']) : null;
            
            const dosenTim = [pembimbing1, pembimbing2, penguji1, penguji2];
            
            let slotDitemukan = false;
            let slotTerpilih = null;
            let isRequest = false;
            
            // Coba penuhi request terlebih dahulu
            if (requestTanggal && requestSesi) {
                const requestDate = parseDate(requestTanggal);
                const dayIndex = requestDate.getDay();
                const requestDay = indexToHari[dayIndex === 0 ? 1 : (dayIndex === 6 ? 5 : dayIndex)];
                const formattedRequestDate = formatDate(requestDate);
                const requestKey = `${requestDay}-${formattedRequestDate}`;
                
                if (jadwalHariTanggal[requestKey] && 
                    jadwalHariTanggal[requestKey].slots[requestSesi] && 
                    jadwalHariTanggal[requestKey].slots[requestSesi].tersedia) {
                    
                    // Buat versi relaksasi: hanya memeriksa bentrok jadwal mengajar
                    let adaYangMengajar = false;
                    for (const dosen of dosenTim) {
                        if (dosenToJadwal[dosen]) {
                            for (const jadwal of dosenToJadwal[dosen]) {
                                if (jadwal.hariTanggalKey === requestKey && jadwal.sesi === requestSesi) {
                                    adaYangMengajar = true;
                                    break;
                                }
                            }
                        }
                        if (adaYangMengajar) break;
                    }
                    
                    // Versi relaksasi: izinkan hingga 1 dosen yang sidang bentrok
                    let jumlahDosenSidangBentrok = 0;
                    let dosenKonflik = [];
                    
                    for (const dosen of dosenTim) {
                        if (jadwalDosenTerpakai[dosen]) {
                            for (const jadwal of jadwalDosenTerpakai[dosen]) {
                                if (jadwal.hariTanggalKey === requestKey && jadwal.sesi === requestSesi) {
                                    jumlahDosenSidangBentrok++;
                                    dosenKonflik.push(dosen);
                                    break;
                                }
                            }
                        }
                    }
                    
                    // Relaksasi: izinkan maksimal 1 dosen konflik untuk request
                    const MAX_KONFLIK_REQUEST = 1;
                    
                    if (!adaYangMengajar && jumlahDosenSidangBentrok <= MAX_KONFLIK_REQUEST) {
                        slotTerpilih = {
                            hariTanggalKey: requestKey,
                            sesi: requestSesi
                        };
                        slotDitemukan = true;
                        isRequest = true;
                        
                        if (jumlahDosenSidangBentrok > 0) {
                            console.log(`Request ${namaMahasiswa} dipenuhi dengan relaksasi (${jumlahDosenSidangBentrok} dosen konflik: ${dosenKonflik.join(", ")})`);
                        } else {
                            console.log(`Request ${namaMahasiswa} dipenuhi tanpa konflik`);
                        }
                    }
                }
            }
            
            // Jika request tidak bisa dipenuhi, cari slot lain
            if (!slotDitemukan) {
                const candidateSlots = [];
                
                Object.keys(jadwalHariTanggal).forEach(hariTanggalKey => {
                    for (let sesi = 1; sesi <= 4; sesi++) {
                        if (jadwalHariTanggal[hariTanggalKey].slots[sesi].tersedia) {
                            // Versi relaksasi: cek apakah ada konflik jadwal mengajar
                            let adaYangMengajar = false;
                            for (const dosen of dosenTim) {
                                if (dosenToJadwal[dosen]) {
                                    for (const jadwal of dosenToJadwal[dosen]) {
                                        if (jadwal.hariTanggalKey === hariTanggalKey && jadwal.sesi === sesi) {
                                            adaYangMengajar = true;
                                            break;
                                        }
                                    }
                                }
                                if (adaYangMengajar) break;
                            }
                            
                            // Relaksasi: hitung dosen yang sudah sidang pada slot ini
                            let jumlahDosenSidangBentrok = 0;
                            let dosenKonflik = [];
                            
                            for (const dosen of dosenTim) {
                                if (jadwalDosenTerpakai[dosen]) {
                                    for (const jadwal of jadwalDosenTerpakai[dosen]) {
                                        if (jadwal.hariTanggalKey === hariTanggalKey && jadwal.sesi === sesi) {
                                            jumlahDosenSidangBentrok++;
                                            dosenKonflik.push(dosen);
                                            break;
                                        }
                                    }
                                }
                            }
                            
                            // Relaksasi: izinkan maksimal 2 dosen konflik untuk non-request
                            const MAX_KONFLIK_NORMAL = 2;
                            
                            if (!adaYangMengajar && jumlahDosenSidangBentrok <= MAX_KONFLIK_NORMAL) {
                                // Hitung skor untuk slot ini
                                let score = 100 - (jumlahDosenSidangBentrok * 40); // Penalti besar untuk konflik
                                
                                // Preferensi hari dan sesi
                                const hari = jadwalHariTanggal[hariTanggalKey].hari;
                                if (hari === 'Senin' || hari === 'Selasa') score += 5;
                                if (sesi === 2 || sesi === 3) score += 3;
                                
                                // Tambahkan ke kandidat
                                candidateSlots.push({
                                    hariTanggalKey: hariTanggalKey,
                                    sesi: sesi,
                                    score: score,
                                    konflik: jumlahDosenSidangBentrok,
                                    dosenKonflik: dosenKonflik
                                });
                            }
                        }
                    }
                });
                
                // Urutkan berdasarkan skor (tertinggi dulu)
                candidateSlots.sort((a, b) => b.score - a.score);
                
                // Ambil slot terbaik jika ada
                if (candidateSlots.length > 0) {
                    slotTerpilih = candidateSlots[0];
                    slotDitemukan = true;
                    
                    if (slotTerpilih.konflik > 0) {
                        console.log(`${namaMahasiswa} dijadwalkan dengan ${slotTerpilih.konflik} dosen konflik: ${slotTerpilih.dosenKonflik.join(", ")}`);
                    }
                }
            }
            
            // Jika ditemukan slot, jadwalkan
            if (slotDitemukan) {
                const hariTanggalKey = slotTerpilih.hariTanggalKey;
                const sesi = slotTerpilih.sesi;
                
                // Tandai slot ini sebagai digunakan
                jadwalHariTanggal[hariTanggalKey].slots[sesi].tersedia = false;
                jadwalHariTanggal[hariTanggalKey].slots[sesi].jadwalSidang = {
                    mahasiswa: namaMahasiswa,
                    dosen: dosenTim,
                    isRequest: isRequest
                };
                
                // Update jadwal dosen terpakai (meskipun mungkin ada konflik)
                dosenTim.forEach(dosen => {
                    jadwalDosenTerpakai[dosen].push({
                        hariTanggalKey: hariTanggalKey,
                        sesi: sesi
                    });
                });
                
                // Tambahkan ke hasil
                hasilJadwal[originalIndex] = {
                    'Nama Mahasiswa': namaMahasiswa,
                    'Pembimbing 1': pembimbing1,
                    'Pembimbing 2': pembimbing2,
                    'Penguji 1': penguji1,
                    'Penguji 2': penguji2,
                    'Hari': jadwalHariTanggal[hariTanggalKey].hari,
                    'Tanggal': jadwalHariTanggal[hariTanggalKey].tanggal,
                    'Sesi': sesi,
                    'Jam': jadwalJam[sesi],
                    'IsRequest': isRequest
                };
                
                if (isRequest) {
                    totalRequestDipenuhi++;
                }
                
                assignedCount++;
            } else {
                // Tidak menemukan slot
                hasilJadwal[originalIndex] = {
                    'Nama Mahasiswa': namaMahasiswa,
                    'Pembimbing 1': pembimbing1,
                    'Pembimbing 2': pembimbing2,
                    'Penguji 1': penguji1,
                    'Penguji 2': penguji2,
                    'Hari': 'Tidak ada slot tersedia',
                    'Tanggal': '-',
                    'Sesi': '-',
                    'Jam': '-',
                    'IsRequest': false
                };
            }
        }
    }
    
    console.timeEnd('Waktu Penjadwalan');
    console.log(`Penjadwalan selesai: ${assignedCount}/${timSidang.length} mahasiswa berhasil dijadwalkan (${Math.round(assignedCount/timSidang.length*100)}%)`);
    console.log(`Request terpenuhi: ${totalRequestDipenuhi}/${totalRequest} (${Math.round(totalRequestDipenuhi/totalRequest*100)}%)`);
    
    // Hitung beban akhir dosen
    const bebanAkhirDosen = {};
    Object.values(daftarDosen).forEach(nama => {
        bebanAkhirDosen[nama] = 0;
    });
    
    hasilJadwal.forEach(h => {
        if (h['Hari'] !== 'Tidak ada slot tersedia') {
            bebanAkhirDosen[h['Pembimbing 1']] = (bebanAkhirDosen[h['Pembimbing 1']] || 0) + 1;
            bebanAkhirDosen[h['Pembimbing 2']] = (bebanAkhirDosen[h['Pembimbing 2']] || 0) + 1;
            bebanAkhirDosen[h['Penguji 1']] = (bebanAkhirDosen[h['Penguji 1']] || 0) + 1;
            bebanAkhirDosen[h['Penguji 2']] = (bebanAkhirDosen[h['Penguji 2']] || 0) + 1;
        }
    });
    
    console.log('Beban akhir dosen:');
    Object.entries(bebanAkhirDosen).sort((a, b) => b[1] - a[1]).forEach(([dosen, beban]) => {
        console.log(`- ${dosen}: ${beban} sidang`);
    });
    
    return hasilJadwal;
}

// Tambahan: Cek apakah data request valid
function cekRequestValid(timSidang, jadwalHariTanggal) {
    const requestTidakValid = [];
    
    timSidang.forEach(tim => {
        const namaMahasiswa = tim['Nama Mahasiswa / NIM'];
        const requestTanggal = tim['Request Tanggal'];
        const requestSesi = tim['Request Sesi'] ? parseInt(tim['Request Sesi']) : null;
        
        if (requestTanggal && requestSesi) {
            const requestDate = parseDate(requestTanggal);
            const dayIndex = requestDate.getDay();
            const requestDay = indexToHari[dayIndex === 0 ? 1 : (dayIndex === 6 ? 5 : dayIndex)]; // Fallback sesuai
            const formattedRequestDate = formatDate(requestDate);
            const requestKey = `${requestDay}-${formattedRequestDate}`;
            
            if (!jadwalHariTanggal[requestKey]) {
                requestTidakValid.push({
                    mahasiswa: namaMahasiswa,
                    tanggal: requestTanggal,
                    sesi: requestSesi,
                    alasan: 'Tanggal tidak ada dalam rentang penjadwalan'
                });
            }
        }
    });
    
    if (requestTidakValid.length > 0) {
        console.log('PERHATIAN: Ditemukan request yang tidak valid:');
        requestTidakValid.forEach(req => {
            console.log(`- ${req.mahasiswa}: Tanggal ${req.tanggal} Sesi ${req.sesi} (${req.alasan})`);
        });
    }
    
    return requestTidakValid;
}

// PERBAIKAN: Fungsi untuk membuat daftar slot tersedia setelah penjadwalan
function generateAvailableSlots(jadwalMengajar, hasilJadwal, jadwalHariTanggal) {
    const slotTersedia = [];
    
    // Iterasi semua slot dari awal hingga akhir periode
    Object.keys(jadwalHariTanggal).forEach(hariTanggalKey => {
        const hariTanggal = jadwalHariTanggal[hariTanggalKey];
        
        // Iterasi semua sesi dalam hari ini
        for (let sesi = 1; sesi <= 4; sesi++) {
            if (hariTanggal.slots[sesi].tersedia) {
                // Cari dosen yang tersedia pada slot ini
                const dosenTersedia = [];
                
                Object.keys(daftarDosen).forEach(kode => {
                    const namaLengkap = daftarDosen[kode];
                    let dosenSibuk = false;
                    
                    // Cek apakah dosen mengajar pada slot ini
                    for (const jadwal of jadwalMengajar) {
                        if (jadwal['Kode Dosen'] === kode && 
                            hariTanggal.hari === jadwal['Hari'] && 
                            sesi === parseInt(jadwal['Sesi'])) {
                            dosenSibuk = true;
                            break;
                        }
                    }
                    
                    // Cek apakah dosen sidang pada slot ini
                    if (!dosenSibuk) {
                        for (const jadwal of hasilJadwal) {
                            if (jadwal['Hari'] !== 'Tidak ada slot tersedia' &&
                                jadwal['Hari'] === hariTanggal.hari &&
                                jadwal['Tanggal'] === hariTanggal.tanggal &&
                                jadwal['Sesi'] === sesi &&
                                (jadwal['Pembimbing 1'] === namaLengkap ||
                                 jadwal['Pembimbing 2'] === namaLengkap ||
                                 jadwal['Penguji 1'] === namaLengkap ||
                                 jadwal['Penguji 2'] === namaLengkap)) {
                                dosenSibuk = true;
                                break;
                            }
                        }
                    }
                    
                    if (!dosenSibuk) {
                        dosenTersedia.push(kode);
                    }
                });
                
                // Hanya tambahkan slot jika ada dosen yang tersedia
                if (dosenTersedia.length > 0) {
                    slotTersedia.push({
                        'Hari': hariTanggal.hari,
                        'Tanggal': hariTanggal.tanggal,
                        'Sesi': sesi,
                        'Jam': jadwalJam[sesi],
                        'Ketersediaan Dosen': dosenTersedia.join(', ')
                    });
                }
            }
        }
    });
    
    return slotTersedia;
}

// Parsing data Excel
function processExcelFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Baca jadwal mengajar
        const jadwalSheet = workbook.Sheets['JadwalMengajar'];
        const jadwalData = XLSX.utils.sheet_to_json(jadwalSheet, { header: 1 });
        
        // Konversi data ke format yang dibutuhkan
        const headers = jadwalData[0];
        jadwalMengajar = [];
        
        for (let i = 1; i < jadwalData.length; i++) {
            const row = jadwalData[i];
            if (row.length > 0) {
                const item = {};
                for (let j = 0; j < headers.length; j++) {
                    item[headers[j]] = row[j];
                }
                jadwalMengajar.push(item);
            }
        }
        
        // Baca tim sidang
        const timSheet = workbook.Sheets['TimSidang'];
        const timData = XLSX.utils.sheet_to_json(timSheet, { header: 1 });
        
        // Konversi data ke format yang dibutuhkan
        const timHeaders = timData[0];
        timSidang = [];
        
        for (let i = 1; i < timData.length; i++) {
            const row = timData[i];
            if (row.length > 0) {
                const item = {};
                for (let j = 0; j < timHeaders.length; j++) {
                    item[timHeaders[j]] = row[j] || ''; // Pastikan nilai tidak undefined
                }
                timSidang.push(item);
            }
        }
        
        // Dapatkan tanggal mulai sidang dari input
        const tanggalMulai = document.getElementById('tanggal-mulai').value;
        if (!tanggalMulai) {
            alert('Silakan masukkan tanggal mulai sidang terlebih dahulu!');
            return;
        }

        // Konversi format input date (YYYY-MM-DD) ke format DD/MM/YYYY
        const dateObj = new Date(tanggalMulai);
        const formattedDate = formatDate(dateObj);
        
        // Tampilkan info dasar
        console.log(`Total jadwal mengajar: ${jadwalMengajar.length}`);
        console.log(`Total tim sidang: ${timSidang.length}`);
        console.log(`Tanggal mulai sidang: ${formattedDate}`);
        
        // Inisialisasi jadwal hari-tanggal
        jadwalHariTanggal = inisialisasiTanggal(tanggalMulai, timSidang.length);
        
        // Cek apakah ada request yang tidak valid
        const requestTidakValid = cekRequestValid(timSidang, jadwalHariTanggal);
        
        // Jika ada request tidak valid, tampilkan peringatan
        if (requestTidakValid.length > 0) {
            const pesanPesan = [
                'Perhatian: Beberapa request jadwal tidak valid:',
                ...requestTidakValid.map(req => `- ${req.mahasiswa}: Tanggal ${req.tanggal} Sesi ${req.sesi}`)
            ].join('\n');
            
            alert(pesanPesan + '\n\nPenjadwalan tetap akan dilanjutkan.');
        }
        
        // Lakukan penjadwalan
        console.time('Waktu Penjadwalan');
        hasilJadwal = scheduleTA(jadwalMengajar, timSidang, jadwalHariTanggal);
        console.timeEnd('Waktu Penjadwalan');
        
        // Generate slot tersedia
        slotTersedia = generateAvailableSlots(jadwalMengajar, hasilJadwal, jadwalHariTanggal);
        
        // Tampilkan hasil
        displayResults(hasilJadwal);
        
        // Tampilkan slot tersedia
        displayAvailableSlots(slotTersedia);
        
        // Update kalender
        updateCalendar(jadwalHariTanggal);
        
        // Aktifkan tab hasil
        document.getElementById('result-tab').click();
        
        // Isi dropdown filter dosen
        populateDosenDropdowns();
        
        // Isi dropdown minggu untuk kalender
        populateWeekDropdown();
        
        // Tampilkan statistik penjadwalan
        const terjadwal = hasilJadwal.filter(h => h.Hari !== 'Tidak ada slot tersedia').length;
        console.log(`Ringkasan: ${terjadwal} dari ${timSidang.length} mahasiswa berhasil dijadwalkan (${Math.round(terjadwal/timSidang.length*100)}%)`);
    };
    
    reader.readAsArrayBuffer(file);
}

// Menampilkan hasil
function displayResults(results) {
    const tableBody = document.querySelector('#result-table tbody');
    
    // Kosongkan tabel
    tableBody.innerHTML = '';
    
    // Isi tabel dengan hasil
    results.forEach(result => {
        const row = document.createElement('tr');
        const mahasiswa = result['Nama Mahasiswa'].split(' / ')[0]; // Ambil nama mahasiswa saja tanpa NIM
        
        // Tandai baris yang merupakan request dengan warna berbeda
        if (result['IsRequest']) {
            row.classList.add('request-row');
        }
        
        row.innerHTML = `
            <td>${mahasiswa}</td>
            <td>${result['Pembimbing 1']}</td>
            <td>${result['Pembimbing 2']}</td>
            <td>${result['Penguji 1']}</td>
            <td>${result['Penguji 2']}</td>
            <td>${result.Hari === 'Tidak ada slot tersedia' ? 'Tidak ada slot tersedia' : `${result.Hari}, ${result.Tanggal}`}</td>
            <td>${result.Sesi}</td>
            <td>${result.Jam}</td>
        `;
        
        tableBody.appendChild(row);
    });
}

// Menampilkan slot tersedia
function displayAvailableSlots(slots) {
    const tableBody = document.querySelector('#available-table tbody');
    
    // Kosongkan tabel
    tableBody.innerHTML = '';
    
    // Isi tabel dengan slots
    slots.forEach(slot => {
        const row = document.createElement('tr');
        
        row.innerHTML = `
            <td>${slot.Hari}</td>
            <td>${slot.Tanggal}</td>
            <td>${slot.Sesi}</td>
            <td>${slot.Jam}</td>
            <td>${slot['Ketersediaan Dosen']}</td>
        `;
        
        tableBody.appendChild(row);
    });
}

// Fungsi untuk mengisi dropdown dosen
function populateDosenDropdowns() {
    const filterLecturer = document.getElementById('filter-lecturer');
    const calendarLecturer = document.getElementById('calendar-lecturer');
    
    // Kosongkan opsi kecuali "Semua Dosen"
    while (filterLecturer.options.length > 1) {
        filterLecturer.remove(1);
    }
    
    while (calendarLecturer.options.length > 1) {
        calendarLecturer.remove(1);
    }
    
    // Tambahkan opsi dosen
    Object.keys(daftarDosen).forEach(kode => {
        const nama = daftarDosen[kode];
        
        const option1 = document.createElement('option');
        option1.value = nama;
        option1.textContent = `${nama} (${kode})`;
        filterLecturer.appendChild(option1);
        
        const option2 = document.createElement('option');
        option2.value = nama;
        option2.textContent = `${nama} (${kode})`;
        calendarLecturer.appendChild(option2);
    });
}

// Fungsi untuk mengisi dropdown minggu untuk kalender
function populateWeekDropdown() {
    const calendarWeek = document.getElementById('calendar-week');
    
    // Kosongkan opsi
    calendarWeek.innerHTML = '';
    
    // Kelompokkan berdasarkan minggu
    const weeks = {};
    
    Object.keys(jadwalHariTanggal).forEach(key => {
        const data = jadwalHariTanggal[key];
        const date = new Date(data.date);
        const weekNum = getWeekNumber(date);
        
        if (!weeks[weekNum]) {
            weeks[weekNum] = [];
        }
        
        weeks[weekNum].push({
            key: key,
            data: data
        });
    });
    
    // Tambahkan opsi untuk setiap minggu
    Object.keys(weeks).sort().forEach((weekNum, index) => {
        const weekData = weeks[weekNum];
        const startDate = weekData[0].data.tanggal;
        const endDate = weekData[weekData.length - 1].data.tanggal;
        
        const option = document.createElement('option');
        option.value = weekNum;
        option.textContent = `Minggu ${index + 1} (${startDate} - ${endDate})`;
        calendarWeek.appendChild(option);
    });
    
    // Default to first week
    if (calendarWeek.options.length > 0) {
        calendarWeek.selectedIndex = 0;
    }
    
    // Trigger calendar update
    calendarWeek.dispatchEvent(new Event('change'));
}

// Update tampilan kalender
function updateCalendar() {
    // Reset tampilan kalender
    const calendarContainer = document.querySelector('.calendar-container');
    calendarContainer.innerHTML = '';
    
    // Pilih dosen dan minggu
    const selectedDosen = document.getElementById('calendar-lecturer').value;
    const selectedWeek = document.getElementById('calendar-week').value;
    
    // Kelompokkan berdasarkan minggu
    const weeks = {};
    
    Object.keys(jadwalHariTanggal).forEach(key => {
        const data = jadwalHariTanggal[key];
        const date = new Date(data.date);
        const weekNum = getWeekNumber(date);
        
        if (!weeks[weekNum]) {
            weeks[weekNum] = [];
        }
        
        weeks[weekNum].push({
            key: key,
            data: data
        });
    });
    
    // Hanya tampilkan minggu yang dipilih
    if (selectedWeek && weeks[selectedWeek]) {
        const weekData = weeks[selectedWeek];
        
        // Buat header untuk minggu ini
        const weekHeader = document.createElement('div');
        weekHeader.className = 'week-header';
        weekHeader.textContent = `Minggu ${Object.keys(weeks).indexOf(selectedWeek) + 1}`;
        calendarContainer.appendChild(weekHeader);
        
        // Buat grid kalender untuk minggu ini
        const calendarGrid = document.createElement('div');
        calendarGrid.className = 'calendar-grid';
        
        // Buat header hari
        const calendarHeader = document.createElement('div');
        calendarHeader.className = 'calendar-header';
        
        const emptyCell = document.createElement('div');
        emptyCell.className = 'empty-cell';
        calendarHeader.appendChild(emptyCell);
        
        weekData.forEach(day => {
            const dayHeader = document.createElement('div');
            dayHeader.className = 'day-header';
            dayHeader.textContent = `${day.data.hari}, ${day.data.tanggal}`;
            calendarHeader.appendChild(dayHeader);
        });
        
        calendarGrid.appendChild(calendarHeader);
        
        // Buat body kalender
        const calendarBody = document.createElement('div');
        calendarBody.className = 'calendar-body';
        
        // Untuk setiap sesi
        for (let sesi = 1; sesi <= 4; sesi++) {
            // Tambahkan cell waktu
            const timeCell = document.createElement('div');
            timeCell.className = 'time-cell';
            timeCell.innerHTML = `Sesi ${sesi}<br>${jadwalJam[sesi]}`;
            calendarBody.appendChild(timeCell);
            
            // Untuk setiap hari dalam minggu
            weekData.forEach(day => {
                const cellId = `cell-${day.data.hari}-${day.data.tanggal}-${sesi}`;
                const cell = document.createElement('div');
                cell.className = 'calendar-cell status-available';
                cell.id = cellId;
                
                // Cek slot ini untuk jadwal mengajar
                if (day.data.slots[sesi].jadwalMengajar.length > 0) {
                    let showTeaching = false;
                    
                    if (selectedDosen === 'all') {
                        showTeaching = true;
                    } else {
                        // Cek apakah dosen yang dipilih mengajar pada slot ini
                        for (const jadwal of day.data.slots[sesi].jadwalMengajar) {
                            if (jadwal.dosen === selectedDosen) {
                                showTeaching = true;
                                break;
                            }
                        }
                    }
                    
                    if (showTeaching) {
                        cell.classList.remove('status-available');
                        cell.classList.add('status-teaching');
                        
                        // Tambahkan info jadwal mengajar
                        day.data.slots[sesi].jadwalMengajar.forEach(jadwal => {
                            if (selectedDosen === 'all' || jadwal.dosen === selectedDosen) {
                                const eventDiv = document.createElement('div');
                                eventDiv.className = 'calendar-event event-teaching';
                                eventDiv.textContent = `${jadwal.mataKuliah} (${kodeToDosen[jadwal.dosen]}) - ${jadwal.ruangan}`;
                                cell.appendChild(eventDiv);
                            }
                        });
                    }
                }
                
                // Cek slot ini untuk jadwal sidang
                if (day.data.slots[sesi].jadwalSidang) {
                    const sidang = day.data.slots[sesi].jadwalSidang;
                    let showSidang = false;
                    
                    if (selectedDosen === 'all') {
                        showSidang = true;
                    } else {
                        // Cek apakah dosen yang dipilih terlibat dalam sidang ini
                        for (const dosen of sidang.dosen) {
                            if (dosen === selectedDosen) {
                                showSidang = true;
                                break;
                            }
                        }
                    }
                    
                    if (showSidang) {
                        cell.classList.remove('status-available');
                        cell.classList.remove('status-teaching');
                        cell.classList.add(sidang.isRequest ? 'status-request' : 'status-defense');
                        
                        // Tambahkan info jadwal sidang
                        const eventDiv = document.createElement('div');
                        eventDiv.className = sidang.isRequest ? 'calendar-event event-request' : 'calendar-event event-defense';
                        
                        // Ambil nama mahasiswa saja tanpa NIM
                        const mahasiswa = sidang.mahasiswa.split(' / ')[0];
                        
                        eventDiv.textContent = `Sidang TA: ${mahasiswa}`;
                        
                        if (sidang.isRequest) {
                            eventDiv.textContent += ' (Request)';
                        }
                        
                        cell.appendChild(eventDiv);
                    }
                }
                
                calendarBody.appendChild(cell);
            });
        }
        
        calendarGrid.appendChild(calendarBody);
        calendarContainer.appendChild(calendarGrid);
    }
}

// Filter slot tersedia
function filterAvailableSlots() {
    const dayFilter = document.getElementById('filter-day').value;
    const sessionFilter = document.getElementById('filter-session').value;
    const lecturerFilter = document.getElementById('filter-lecturer').value;
    
    let filteredSlots = [...slotTersedia];
    
    // Filter berdasarkan hari
    if (dayFilter !== 'all') {
        filteredSlots = filteredSlots.filter(slot => slot.Hari === dayFilter);
    }
    
    // Filter berdasarkan sesi
    if (sessionFilter !== 'all') {
        filteredSlots = filteredSlots.filter(slot => slot.Sesi === parseInt(sessionFilter));
    }
    
    // Filter berdasarkan dosen
    if (lecturerFilter !== 'all') {
        filteredSlots = filteredSlots.filter(slot => {
            const dosenList = slot['Ketersediaan Dosen'].split(', ');
            for (const kodeDosen of dosenList) {
                if (daftarDosen[kodeDosen] === lecturerFilter) {
                    return true;
                }
            }
            return false;
        });
    }
    
    // Tampilkan hasil filter
    displayAvailableSlots(filteredSlots);
}

// Export hasil ke Excel
function exportResults() {
    // Buat workbook baru
    const wb = XLSX.utils.book_new();
    
    // Data untuk sheet hasil
    const hasilData = [
        ['Nama Mahasiswa', 'Pembimbing 1', 'Pembimbing 2', 'Penguji 1', 'Penguji 2', 'Hari', 'Tanggal', 'Sesi', 'Jam', 'Request']
    ];
    
    hasilJadwal.forEach(jadwal => {
        const mahasiswa = jadwal['Nama Mahasiswa'].split(' / ')[0];
        
        hasilData.push([
            mahasiswa,
            jadwal['Pembimbing 1'],
            jadwal['Pembimbing 2'],
            jadwal['Penguji 1'],
            jadwal['Penguji 2'],
            jadwal.Hari,
            jadwal.Tanggal,
            jadwal.Sesi,
            jadwal.Jam,
            jadwal.IsRequest ? 'Ya' : 'Tidak'
        ]);
    });
    
    const hasilSheet = XLSX.utils.aoa_to_sheet(hasilData);
    XLSX.utils.book_append_sheet(wb, hasilSheet, 'Hasil Penjadwalan');
    
    // Tambahkan sheet statistik
    const statsData = [['Nama Dosen', 'Jumlah Sidang']];
    
    const bebanPerDosen = {};
    hasilJadwal.forEach(h => {
        if (h.Hari !== 'Tidak ada slot tersedia') {
            bebanPerDosen[h['Pembimbing 1']] = (bebanPerDosen[h['Pembimbing 1']] || 0) + 1;
            bebanPerDosen[h['Pembimbing 2']] = (bebanPerDosen[h['Pembimbing 2']] || 0) + 1;
            bebanPerDosen[h['Penguji 1']] = (bebanPerDosen[h['Penguji 1']] || 0) + 1;
            bebanPerDosen[h['Penguji 2']] = (bebanPerDosen[h['Penguji 2']] || 0) + 1;
        }
    });
    
    Object.entries(bebanPerDosen).sort((a, b) => b[1] - a[1]).forEach(([dosen, beban]) => {
        statsData.push([dosen, beban]);
    });
    
    const statsSheet = XLSX.utils.aoa_to_sheet(statsData);
    XLSX.utils.book_append_sheet(wb, statsSheet, 'Statistik Beban Dosen');
    
    // Download file
    XLSX.writeFile(wb, 'hasil_sidang_ta.xlsx');
}

// Export slot tersedia ke Excel
function exportAvailableSlots() {
    // Buat workbook baru
    const wb = XLSX.utils.book_new();
    
    // Data untuk sheet slot tersedia
    const slotData = [
        ['Hari', 'Tanggal', 'Sesi', 'Jam', 'Ketersediaan Dosen']
    ];
    
    slotTersedia.forEach(slot => {
        slotData.push([
            slot.Hari,
            slot.Tanggal,
            slot.Sesi,
            slot.Jam,
            slot['Ketersediaan Dosen']
        ]);
    });
    
    const slotSheet = XLSX.utils.aoa_to_sheet(slotData);
    XLSX.utils.book_append_sheet(wb, slotSheet, 'Slot Tersedia');
    
    // Download file
    XLSX.writeFile(wb, 'slot_tersedia.xlsx');
}

// Event listeners
document.addEventListener('DOMContentLoaded', function() {
    // Download template
    document.getElementById('download-template').addEventListener('click', createTemplate);
    
    // File upload
    document.getElementById('file-input').addEventListener('change', function(e) {
        if (e.target.files.length) {
            processExcelFile(e.target.files[0]);
        }
    });
    
    // Drag & drop
    const dropArea = document.getElementById('drop-area');
    
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });
    
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, unhighlight, false);
    });
    
    function highlight() {
        dropArea.classList.add('highlight');
    }
    
    function unhighlight() {
        dropArea.classList.remove('highlight');
    }
    
    dropArea.addEventListener('drop', function(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length && files[0].name.endsWith('.xlsx')) {
            processExcelFile(files[0]);
        }
    });
    
    // Export hasil
    document.getElementById('export-result').addEventListener('click', exportResults);
    
    // Export slot tersedia
    document.getElementById('export-available').addEventListener('click', exportAvailableSlots);
    
    // Filter slot tersedia
    document.getElementById('filter-day').addEventListener('change', filterAvailableSlots);
    document.getElementById('filter-session').addEventListener('change', filterAvailableSlots);
    document.getElementById('filter-lecturer').addEventListener('change', filterAvailableSlots);
    
    // Update kalender saat pilihan dosen atau minggu berubah
    document.getElementById('calendar-lecturer').addEventListener('change', updateCalendar);
    document.getElementById('calendar-week').addEventListener('change', updateCalendar);
});