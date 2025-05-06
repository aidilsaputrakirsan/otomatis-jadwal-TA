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
    
    // Hitung perkiraan jumlah minggu yang dibutuhkan (1 minggu dapat menampung sekitar 20 mahasiswa)
    jumlahMinggu = Math.ceil(jumlahMahasiswa / 10) + 2; // Tambah 1 untuk buffer
    
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

// Algoritma penjadwalan
function scheduleTA(jadwalMengajar, timSidang, jadwalHariTanggal) {
    const hasilJadwal = [];
    
    // Jadwal yang sudah digunakan oleh dosen
    const jadwalDosenTerpakai = {};
    
    // Buat pemetaan nama lengkap dosen ke jadwal mengajar mereka
    const dosenToJadwal = {};
    
    // Memasukkan jadwal mengajar ke pemetaan
    jadwalMengajar.forEach(jadwal => {
        const kodeDosen = jadwal['Kode Dosen'];
        const namaLengkap = daftarDosen[kodeDosen];
        
        if (!dosenToJadwal[namaLengkap]) {
            dosenToJadwal[namaLengkap] = [];
        }
        
        // Simpan jadwal mengajar dengan hari-tanggal
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
    
    // Urutkan tim sidang: prioritaskan yang memiliki request jadwal
    const sortedTimSidang = [...timSidang].sort((a, b) => {
        const aHasRequest = a['Request Tanggal'] && a['Request Sesi'];
        const bHasRequest = b['Request Tanggal'] && b['Request Sesi'];
        
        if (aHasRequest && !bHasRequest) return -1;
        if (!aHasRequest && bHasRequest) return 1;
        return 0;
    });
    
    // Untuk debugging
    console.log("Total hari yang tersedia:", Object.keys(jadwalHariTanggal).length);
    console.log("Total mahasiswa:", sortedTimSidang.length);
    
    // Untuk setiap tim sidang
    sortedTimSidang.forEach((tim, index) => {
        const namaMahasiswa = tim['Nama Mahasiswa / NIM'];
        const pembimbing1 = tim['Pembimbing 1'];
        const pembimbing2 = tim['Pembimbing 2'];
        const penguji1 = tim['Penguji 1'];
        const penguji2 = tim['Penguji 2'];
        const requestTanggal = tim['Request Tanggal'];
        const requestSesi = tim['Request Sesi'] ? parseInt(tim['Request Sesi']) : null;
        
        const dosenTim = [pembimbing1, pembimbing2, penguji1, penguji2];
        
        // Cek apakah ada request jadwal
        if (requestTanggal && requestSesi) {
            // Cari key untuk hari-tanggal yang sesuai dengan request
            const requestDate = parseDate(requestTanggal);
            // getDay() mengembalikan 0-6, sesuaikan untuk indexToHari yang menggunakan 1-5
            const dayIndex = requestDate.getDay();
            const requestDay = indexToHari[dayIndex === 0 ? 1 : dayIndex]; // Fallback ke Senin jika Minggu
            // Format tanggal ke string DD/MM/YYYY untuk key
            const formattedRequestDate = formatDate(requestDate);
            const requestKey = `${requestDay}-${formattedRequestDate}`;
            
            // Cek apakah key tersebut ada dalam jadwalHariTanggal
            if (jadwalHariTanggal[requestKey] && 
                jadwalHariTanggal[requestKey].slots[requestSesi] && 
                jadwalHariTanggal[requestKey].slots[requestSesi].tersedia) {
                
                // Cek apakah ada dosen tim yang mengajar pada slot ini
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
                    // Tandai slot ini sebagai digunakan
                    jadwalHariTanggal[requestKey].slots[requestSesi].tersedia = false;
                    jadwalHariTanggal[requestKey].slots[requestSesi].jadwalSidang = {
                        mahasiswa: namaMahasiswa,
                        dosen: dosenTim,
                        isRequest: true
                    };
                    
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
                    
                    // Tambahkan ke hasil
                    hasilJadwal.push({
                        'Nama Mahasiswa': namaMahasiswa,
                        'Pembimbing 1': pembimbing1,
                        'Pembimbing 2': pembimbing2,
                        'Penguji 1': penguji1,
                        'Penguji 2': penguji2,
                        'Hari': jadwalHariTanggal[requestKey].hari,
                        'Tanggal': jadwalHariTanggal[requestKey].tanggal,
                        'Sesi': requestSesi,
                        'Jam': jadwalJam[requestSesi],
                        'IsRequest': true
                    });
                    
                    return; // Lanjut ke mahasiswa berikutnya
                }
            }
        }
        
        // Jika tidak ada request atau request tidak bisa dipenuhi, cari slot tersedia
        const slotTersedia = [];
        
        // Iterasi semua slot dari awal hingga akhir periode
        Object.keys(jadwalHariTanggal).forEach(hariTanggalKey => {
            const hariTanggal = jadwalHariTanggal[hariTanggalKey];
            
            // Iterasi semua sesi dalam hari ini
            for (let sesi = 1; sesi <= 4; sesi++) {
                if (hariTanggal.slots[sesi].tersedia) {
                    // Cek apakah ada dosen yang mengajar pada slot ini
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

        console.log(`Mahasiswa #${index+1}: ${namaMahasiswa} - Slot tersedia: ${slotTersedia.length}`);
        
        // Di dalam loop mahasiswa pada scheduleTA
        if (slotTersedia.length === 0) {
            console.log(`TIDAK ADA SLOT untuk ${namaMahasiswa}. Dosen tim:`);
            dosenTim.forEach(dosen => {
                console.log(`- ${dosen}: ${jadwalDosenTerpakai[dosen] ? jadwalDosenTerpakai[dosen].length : 0} jadwal terpakai`);
            });
            console.log("Cek konflik jadwal dosen:");
            
            // Periksa setiap slot yang ada
            let totalSlot = 0;
            Object.keys(jadwalHariTanggal).forEach(key => {
                for (let sesi = 1; sesi <= 4; sesi++) {
                    if (jadwalHariTanggal[key].slots[sesi].tersedia) {
                        totalSlot++;
                        
                        // Cek mengapa slot ini tidak bisa digunakan
                        let alasanTolak = [];
                        
                        for (const dosen of dosenTim) {
                            // Cek jadwal mengajar
                            if (dosenToJadwal[dosen]) {
                                for (const jadwal of dosenToJadwal[dosen]) {
                                    if (jadwal.hariTanggalKey === key && jadwal.sesi === sesi) {
                                        alasanTolak.push(`${dosen} mengajar`);
                                        break;
                                    }
                                }
                            }
                            
                            // Cek jadwal sidang
                            if (jadwalDosenTerpakai[dosen]) {
                                for (const jadwal of jadwalDosenTerpakai[dosen]) {
                                    if (jadwal.hariTanggalKey === key && jadwal.sesi === sesi) {
                                        alasanTolak.push(`${dosen} sidang lain`);
                                        break;
                                    }
                                }
                            }
                        }
                        
                        if (alasanTolak.length > 0) {
                            console.log(`  Slot ${jadwalHariTanggal[key].hari}, ${jadwalHariTanggal[key].tanggal}, Sesi ${sesi}: ${alasanTolak.join(', ')}`);
                        }
                    }
                }
            });
            console.log(`Total slot tersedia secara umum: ${totalSlot}`);
        }

        // Pilih slot pertama yang tersedia
        if (slotTersedia.length > 0) {
            // Hitung skor untuk setiap slot tersedia
            slotTersedia.forEach(slot => {
                slot.score = calculateSlotScore(slot, jadwalHariTanggal, jadwalDosenTerpakai, dosenTim);
            });
            
            // Urutkan berdasarkan skor tertinggi
            slotTersedia.sort((a, b) => b.score - a.score);
            const slotTerpilih = slotTersedia[0];
            const hariTanggalKey = slotTerpilih.hariTanggalKey;
            const sesi = slotTerpilih.sesi;
            
            // Tandai slot ini sebagai digunakan
            jadwalHariTanggal[hariTanggalKey].slots[sesi].tersedia = false;
            jadwalHariTanggal[hariTanggalKey].slots[sesi].jadwalSidang = {
                mahasiswa: namaMahasiswa,
                dosen: dosenTim,
                isRequest: false
            };
            
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
            
            // Tambahkan ke hasil
            hasilJadwal.push({
                'Nama Mahasiswa': namaMahasiswa,
                'Pembimbing 1': pembimbing1,
                'Pembimbing 2': pembimbing2,
                'Penguji 1': penguji1,
                'Penguji 2': penguji2,
                'Hari': jadwalHariTanggal[hariTanggalKey].hari,
                'Tanggal': jadwalHariTanggal[hariTanggalKey].tanggal,
                'Sesi': sesi,
                'Jam': jadwalJam[sesi],
                'IsRequest': false
            });
        } else {
            // Tidak ada slot tersedia
            hasilJadwal.push({
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
            });
        }
    });
    
    return hasilJadwal;
}

// Fungsi untuk menghasilkan slot tersedia
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
        
        // Inisialisasi jadwal hari-tanggal
        jadwalHariTanggal = inisialisasiTanggal(tanggalMulai, timSidang.length);
        
        // Lakukan penjadwalan
        hasilJadwal = scheduleTA(jadwalMengajar, timSidang, jadwalHariTanggal);
        
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

// Fungsi untuk menghitung skor slot berdasarkan distribusi beban dosen
function calculateSlotScore(slot, jadwalHariTanggal, jadwalDosenTerpakai, dosenTim) {
    const hariTanggalKey = slot.hariTanggalKey;
    const sesi = slot.sesi;
    const hari = jadwalHariTanggal[hariTanggalKey].hari;
    
    let skor = 0;
    
    // Prioritaskan hari yang sudah memiliki sidang lain (konsolidasi)
    let hariSudahAdaSidang = false;
    Object.keys(jadwalHariTanggal).forEach(key => {
        if (jadwalHariTanggal[key].hari === hari) {
            for (let s = 1; s <= 4; s++) {
                if (jadwalHariTanggal[key].slots[s].jadwalSidang) {
                    hariSudahAdaSidang = true;
                    break;
                }
            }
        }
    });
    
    if (hariSudahAdaSidang) {
        skor += 5; // Bonus untuk konsolidasi sidang pada hari yang sama
    }
    
    // Prioritaskan slot yang mendistribusikan beban dosen lebih merata
    const dosenLoad = {};
    Object.keys(daftarDosen).forEach(kode => {
        dosenLoad[daftarDosen[kode]] = 0;
    });
    
    // Hitung beban saat ini untuk setiap dosen
    Object.keys(jadwalDosenTerpakai).forEach(dosen => {
        dosenLoad[dosen] = jadwalDosenTerpakai[dosen].length;
    });
    
    // Slot dengan dosen yang memiliki beban lebih rendah mendapat skor lebih tinggi
    let totalBeban = 0;
    dosenTim.forEach(dosen => {
        totalBeban += dosenLoad[dosen] || 0;
    });
    
    skor -= totalBeban; // Kurangi skor berbanding lurus dengan beban dosen
    
    return skor;
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