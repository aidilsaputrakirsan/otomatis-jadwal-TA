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

// Helper function untuk mendapatkan pekan dari tanggal mulai
function getPekanDariTanggalMulai(tanggal, tanggalMulai) {
    const selisihHari = Math.floor((tanggal - tanggalMulai) / (24 * 60 * 60 * 1000));
    return Math.floor(selisihHari / 7) + 1;
}

// Function untuk parsing tanggal
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
function inisialisasiTanggal(tanggalMulai, jumlahMingguParam) {
    // Konversi string tanggal ke objek Date
    tanggalMulaiSidang = new Date(tanggalMulai);
    
    // Gunakan parameter atau default 6 minggu
    jumlahMinggu = jumlahMingguParam || 6;
    
    console.log(`Inisialisasi jadwal untuk ${jumlahMinggu} minggu (${jumlahMinggu * 5} hari kerja)`);
    
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
        
        // Hitung pekan
        const pekan = getPekanDariTanggalMulai(currentDate, tanggalMulaiSidang);
        
        // Simpan dalam mapping
        jadwalHariTanggal[key] = {
            hari: namaHari,
            tanggal: tanggalStr,
            date: new Date(currentDate),
            pekan: pekan,
            slots: {}
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

// Penjadwalan TA dengan prioritas request
function scheduleTA(jadwalMengajar, timSidang, jadwalHariTanggal) {
    console.log(`Mulai penjadwalan dengan ${timSidang.length} mahasiswa`);
    
    // Hasil jadwal
    const hasilJadwal = new Array(timSidang.length);
    
    // Buat pemetaan nama lengkap dosen ke jadwal mengajar mereka
    const dosenToJadwal = {};
    Object.values(daftarDosen).forEach(nama => {
        dosenToJadwal[nama] = [];
    });
    
    // Memasukkan jadwal mengajar ke pemetaan
    jadwalMengajar.forEach(jadwal => {
        const kodeDosen = jadwal['Kode Dosen'];
        const namaLengkap = daftarDosen[kodeDosen];
        
        if (!namaLengkap) {
            console.warn(`Kode dosen tidak dikenal: ${kodeDosen}`);
            return;
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
    
    // Ubah struktur data untuk melacak dosen yang sudah ada jadwal sidang
    const dosenSidangTerpakai = {}; 
    
    // Inisialisasi untuk semua dosen
    Object.values(daftarDosen).forEach(nama => {
        dosenSidangTerpakai[nama] = {}; // Untuk setiap dosen, simpan slot waktu yang terpakai
    });
    
    // PERBAIKAN: Gunakan satu array untuk menyimpan request yang terpenuhi dan tidak
    const requestTerpenuhi = [];
    const requestTidakTerpenuhi = [];
    
    // Langkah 1: Jadwalkan semua request terlebih dahulu (MUTLAK) - TIDAK PEDULI BENTROK JADWAL MENGAJAR
    console.log("=== Langkah 1: Jadwalkan semua request (prioritas mutlak) ===");
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
        
        // Jika ada request, jadwalkan MUTLAK tanpa mempertimbangkan bentrok jadwal mengajar
        if (requestTanggal && requestSesi) {
            // Log untuk debugging
            console.log(`Processing request for ${namaMahasiswa}:`);
            console.log(` - Original date value: ${requestTanggal} (${typeof requestTanggal})`);
            
            const requestDate = parseDate(requestTanggal);
            console.log(` - Parsed date: ${requestDate}`);
            
            const dayIndex = requestDate.getDay();
            console.log(` - Day of week: ${dayIndex}`);
            
            // Jika weekend, sesuaikan ke hari kerja terdekat
            let adjustedDayIndex = dayIndex;
            if (dayIndex === 0) adjustedDayIndex = 1; // Minggu -> Senin
            if (dayIndex === 6) adjustedDayIndex = 5; // Sabtu -> Jumat
            
            const requestDay = indexToHari[adjustedDayIndex];
            const formattedRequestDate = formatDate(requestDate);
            const requestKey = `${requestDay}-${formattedRequestDate}`;
            
            console.log(` - Adjusted day: ${requestDay}`);
            console.log(` - Formatted date: ${formattedRequestDate}`);
            console.log(` - Request key: ${requestKey}`);
            
            // Cek apakah key tersebut ada dalam jadwalHariTanggal
            if (jadwalHariTanggal[requestKey]) {
                // Cek apakah ada dosen dalam tim yang sudah punya jadwal sidang pada slot ini
                let dosenBentrok = [];
                for (const dosen of dosenTim) {
                    const slotKey = `${requestKey}-${requestSesi}`;
                    if (dosenSidangTerpakai[dosen] && dosenSidangTerpakai[dosen][slotKey]) {
                        dosenBentrok.push(dosen);
                    }
                }
                
                if (dosenBentrok.length > 0) {
                    console.log(`${namaMahasiswa}: Request bentrok dengan sidang lain - Dosen ${dosenBentrok.join(', ')} sudah ada jadwal pada slot ini`);
                    requestTidakTerpenuhi.push({
                        mahasiswa: namaMahasiswa,
                        tanggal: formattedRequestDate,
                        hari: requestDay,
                        sesi: requestSesi,
                        alasan: `Dosen bentrok: ${dosenBentrok.join(', ')}`
                    });
                    continue;
                }
                
                // PERBAIKAN: Tandai jadwal mengajar sebagai diganti untuk sidang
                let dosenMengajar = [];
                for (const dosen of dosenTim) {
                    if (dosenToJadwal[dosen]) {
                        for (const jadwal of dosenToJadwal[dosen]) {
                            if (jadwal.hariTanggalKey === requestKey && jadwal.sesi === requestSesi) {
                                dosenMengajar.push(dosen);
                            }
                        }
                    }
                }
                
                if (dosenMengajar.length > 0) {
                    console.log(`${namaMahasiswa}: Request OK, TETAPI MENGGANTI jadwal mengajar ${dosenMengajar.join(', ')}`);
                } else {
                    console.log(`${namaMahasiswa}: Request diterima tanpa bentrok`);
                }
                
                // Tandai dosen sebagai terpakai pada slot ini
                for (const dosen of dosenTim) {
                    if (!dosenSidangTerpakai[dosen]) {
                        dosenSidangTerpakai[dosen] = {};
                    }
                    const slotKey = `${requestKey}-${requestSesi}`;
                    dosenSidangTerpakai[dosen][slotKey] = true;
                }
                
                // Hapus jadwal mengajar pada slot ini (jika ada)
                jadwalHariTanggal[requestKey].slots[requestSesi].tersedia = false;
                jadwalHariTanggal[requestKey].slots[requestSesi].jadwalMengajar = []; // Hapus jadwal mengajar
                jadwalHariTanggal[requestKey].slots[requestSesi].jadwalSidang = {
                    mahasiswa: namaMahasiswa,
                    dosen: dosenTim,
                    isRequest: true
                };
                
                // Tambahkan ke hasil
                hasilJadwal[i] = {
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
                };
                
                // Catat request yang terpenuhi
                requestTerpenuhi.push({
                    mahasiswa: namaMahasiswa,
                    tanggal: formattedRequestDate,
                    hari: requestDay,
                    sesi: requestSesi
                });
            } else {
                console.log(`${namaMahasiswa}: Request tanggal tidak dalam rentang penjadwalan (${requestDay}, ${formattedRequestDate})`);
                requestTidakTerpenuhi.push({
                    mahasiswa: namaMahasiswa,
                    tanggal: formattedRequestDate,
                    hari: requestDay,
                    sesi: requestSesi,
                    alasan: 'Tanggal di luar rentang'
                });
                
                // Lewati dan akan dijadwalkan pada tahap berikutnya
                continue;
            }
        } else {
            // Lewati yang tidak punya request (akan dijadwalkan pada langkah berikutnya)
            continue;
        }
    }
    
    // Tampilkan informasi request
    console.log(`\nREQUEST TERPENUHI: ${requestTerpenuhi.length}`);
    requestTerpenuhi.forEach(req => {
        console.log(`- ${req.mahasiswa}: ${req.hari}, ${req.tanggal}, Sesi ${req.sesi}`);
    });
    
    console.log(`\nREQUEST TIDAK TERPENUHI: ${requestTidakTerpenuhi.length}`);
    requestTidakTerpenuhi.forEach(req => {
        console.log(`- ${req.mahasiswa}: ${req.hari}, ${req.tanggal}, Sesi ${req.sesi} (${req.alasan})`);
    });
    
    // Langkah 2: Jadwalkan mahasiswa yang belum terjadwal
    console.log("\n=== Langkah 2: Jadwalkan mahasiswa yang belum terjadwal ===");
    for (let i = 0; i < sortedTimSidang.length; i++) {
        const tim = sortedTimSidang[i];
        const namaMahasiswa = tim['Nama Mahasiswa / NIM'];
        const pembimbing1 = tim['Pembimbing 1'];
        const pembimbing2 = tim['Pembimbing 2'];
        const penguji1 = tim['Penguji 1'];
        const penguji2 = tim['Penguji 2'];
        
        const dosenTim = [pembimbing1, pembimbing2, penguji1, penguji2];
        
        // Lewati yang sudah terjadwal pada langkah 1
        if (hasilJadwal[i]) {
            continue;
        }
        
        // Cari slot tersedia
        let slotDitemukan = false;
        let hariTanggalKey = null;
        let sesi = null;
        
        // Cari dari hari pertama hingga akhir
        Object.keys(jadwalHariTanggal)
            .sort((a, b) => {
                // Sort berdasarkan tanggal (dari awal ke akhir)
                return jadwalHariTanggal[a].date - jadwalHariTanggal[b].date;
            })
            .some(key => {
                // Cek setiap sesi dalam hari ini
                for (let s = 1; s <= 4; s++) {
                    const slotKey = `${key}-${s}`;
                    
                    // Cek apakah ada dosen tim yang sudah terpakai pada slot ini
                    let dosenBentrok = [];
                    for (const dosen of dosenTim) {
                        if (dosenSidangTerpakai[dosen] && dosenSidangTerpakai[dosen][slotKey]) {
                            dosenBentrok.push(dosen);
                        }
                    }
                    
                    // Jika tidak ada bentrok dengan jadwal sidang dosen, gunakan slot ini
                    if (dosenBentrok.length === 0 && jadwalHariTanggal[key].slots[s].tersedia) {
                        // Tandai dosen sebagai terpakai pada slot ini
                        for (const dosen of dosenTim) {
                            if (!dosenSidangTerpakai[dosen]) {
                                dosenSidangTerpakai[dosen] = {};
                            }
                            dosenSidangTerpakai[dosen][slotKey] = true;
                        }
                        
                        // Simpan key dan sesi
                        hariTanggalKey = key;
                        sesi = s;
                        slotDitemukan = true;
                        
                        // Keluar dari loop
                        return true;
                    }
                }
                // Lanjut ke hari berikutnya
                return false;
            });
        
        // Jika tidak ada slot tanpa bentrok, cari slot dengan bentrok jadwal mengajar
        if (!slotDitemukan) {
            Object.keys(jadwalHariTanggal)
                .sort((a, b) => {
                    // Sort berdasarkan tanggal (dari awal ke akhir)
                    return jadwalHariTanggal[a].date - jadwalHariTanggal[b].date;
                })
                .some(key => {
                    // Cek setiap sesi dalam hari ini
                    for (let s = 1; s <= 4; s++) {
                        const slotKey = `${key}-${s}`;
                        
                        // Cek apakah ada dosen tim yang sudah terpakai pada slot ini
                        let dosenBentrok = [];
                        for (const dosen of dosenTim) {
                            if (dosenSidangTerpakai[dosen] && dosenSidangTerpakai[dosen][slotKey]) {
                                dosenBentrok.push(dosen);
                            }
                        }
                        
                        // Jika tidak ada bentrok dengan jadwal sidang dosen (tapi mungkin bentrok dengan jadwal mengajar), gunakan slot ini
                        if (dosenBentrok.length === 0) {
                            // Tandai dosen sebagai terpakai pada slot ini
                            for (const dosen of dosenTim) {
                                if (!dosenSidangTerpakai[dosen]) {
                                    dosenSidangTerpakai[dosen] = {};
                                }
                                dosenSidangTerpakai[dosen][slotKey] = true;
                            }
                            
                            // Simpan key dan sesi
                            hariTanggalKey = key;
                            sesi = s;
                            slotDitemukan = true;
                            
                            // Keluar dari loop
                            return true;
                        }
                    }
                    // Lanjut ke hari berikutnya
                    return false;
                });
        }
        
        if (slotDitemukan) {
            // Tandai slot dalam jadwalHariTanggal
            jadwalHariTanggal[hariTanggalKey].slots[sesi].tersedia = false;
            jadwalHariTanggal[hariTanggalKey].slots[sesi].jadwalMengajar = []; // Hapus jadwal mengajar jika ada
            jadwalHariTanggal[hariTanggalKey].slots[sesi].jadwalSidang = {
                mahasiswa: namaMahasiswa,
                dosen: dosenTim,
                isRequest: false
            };
            
            // Tambahkan ke hasil
            hasilJadwal[i] = {
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
            };
            
            console.log(`${namaMahasiswa}: Dijadwalkan pada ${jadwalHariTanggal[hariTanggalKey].hari}, ${jadwalHariTanggal[hariTanggalKey].tanggal}, sesi ${sesi}`);
        } else {
            // Tidak ada slot tersedia
            hasilJadwal[i] = {
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
            console.log(`${namaMahasiswa}: TIDAK ADA SLOT TERSEDIA`);
        }
    }
    
    // Hitung statistik
    const mahasiswaTerjadwal = hasilJadwal.filter(h => h && h.Hari !== 'Tidak ada slot tersedia').length;
    const totalMahasiswa = timSidang.length;
    console.log(`\nPenjadwalan selesai: ${mahasiswaTerjadwal}/${totalMahasiswa} mahasiswa berhasil dijadwalkan (${Math.round(mahasiswaTerjadwal/totalMahasiswa*100)}%)`);
    
    return hasilJadwal;
}

// Fungsi untuk menghasilkan slot tersedia
function generateAvailableSlots(jadwalMengajar, hasilJadwal, jadwalHariTanggal) {
    const slotTersedia = [];
    
    // Set untuk menyimpan dosen yang sudah dijadwalkan per slot
    const dosenPerSlot = {};
    
    // Tandai dosen yang sudah dijadwalkan per slot
    hasilJadwal.forEach(jadwal => {
        if (jadwal.Hari !== 'Tidak ada slot tersedia') {
            const key = `${jadwal.Hari}-${jadwal.Tanggal}-${jadwal.Sesi}`;
            
            if (!dosenPerSlot[key]) {
                dosenPerSlot[key] = new Set();
            }
            
            // Tambahkan semua dosen dalam tim
            if (jadwal['Pembimbing 1']) dosenPerSlot[key].add(jadwal['Pembimbing 1']);
            if (jadwal['Pembimbing 2']) dosenPerSlot[key].add(jadwal['Pembimbing 2']);
            if (jadwal['Penguji 1']) dosenPerSlot[key].add(jadwal['Penguji 1']);
            if (jadwal['Penguji 2']) dosenPerSlot[key].add(jadwal['Penguji 2']);
        }
    });
    
    // Iterasi semua slot dari awal hingga akhir periode
    Object.keys(jadwalHariTanggal).forEach(hariTanggalKey => {
        const hariTanggal = jadwalHariTanggal[hariTanggalKey];
        
        // Iterasi semua sesi dalam hari ini
        for (let sesi = 1; sesi <= 4; sesi++) {
            // Cek dosen yang tersedia di slot ini
            const key = `${hariTanggal.hari}-${hariTanggal.tanggal}-${sesi}`;
            const dosenTerpakai = dosenPerSlot[key] || new Set();
            
            // Cari dosen yang tersedia (tidak terpakai di slot ini)
            const dosenTersedia = [];
            Object.keys(daftarDosen).forEach(kode => {
                const namaLengkap = daftarDosen[kode];
                if (!dosenTerpakai.has(namaLengkap)) {
                    dosenTersedia.push(kode);
                }
            });
            
            // Hanya tambahkan slot jika ada dosen yang tersedia
            if (dosenTersedia.length > 0) {
                slotTersedia.push({
                    'Hari': hariTanggal.hari,
                    'Tanggal': hariTanggal.tanggal,
                    'Pekan': hariTanggal.pekan,
                    'Sesi': sesi,
                    'Jam': jadwalJam[sesi],
                    'Ketersediaan Dosen': dosenTersedia.join(', ')
                });
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
        const workbook = XLSX.read(data, {type: 'array'});
        
        // Coba baca jadwal mengajar (jika ada)
        if (workbook.SheetNames.includes('JadwalMengajar')) {
            const jadwalSheet = workbook.Sheets['JadwalMengajar'];
            const jadwalData = XLSX.utils.sheet_to_json(jadwalSheet, {header: 1});
            
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
            
            console.log(`Berhasil membaca ${jadwalMengajar.length} jadwal mengajar dari Excel`);
        } else {
            console.log('Sheet "JadwalMengajar" tidak ditemukan dalam file Excel');
            jadwalMengajar = []; // Reset jadwal mengajar
        }
        
        // Baca tim sidang
        const timSheet = workbook.Sheets['TimSidang'];
        if (!timSheet) {
            alert('Sheet "TimSidang" tidak ditemukan dalam file Excel!');
            return;
        }
        
        const timData = XLSX.utils.sheet_to_json(timSheet, {header: 1});
        
        // Konversi data ke format yang dibutuhkan
        const timHeaders = timData[0];
        timSidang = [];
        
        for (let i = 1; i < timData.length; i++) {
            const row = timData[i];
            if (row.length > 0) {
                const item = {};
                for (let j = 0; j < timHeaders.length; j++) {
                    item[timHeaders[j]] = row[j] || '';
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
        
        // Dapatkan jumlah minggu dari input (jika ada)
        const jumlahMingguInput = parseInt(document.getElementById('jumlah-minggu').value);
        const mingguValid = !isNaN(jumlahMingguInput) && jumlahMingguInput > 0;
        
        // Inisialisasi jadwal hari-tanggal
        jadwalHariTanggal = inisialisasiTanggal(tanggalMulai, mingguValid ? jumlahMingguInput : null);
        
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
        updateCalendar();
        
        // Aktifkan tab hasil
        document.getElementById('result-tab').click();
        
        // Isi dropdown filter dosen
        populateDosenDropdowns();
        
        // Isi dropdown minggu untuk kalender
        populateWeekDropdown();
    };
    
    reader.readAsArrayBuffer(file);
}

// Fungsi untuk mengisi dropdown dosen
function populateDosenDropdowns() {
    const filterLecturer = document.getElementById('filter-lecturer');
    const calendarLecturer = document.getElementById('calendar-lecturer');
    
    if (!filterLecturer || !calendarLecturer) return;
    
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
            <td>${slot.Pekan}</td>
            <td>${slot.Sesi}</td>
            <td>${slot.Jam}</td>
            <td>${slot['Ketersediaan Dosen']}</td>
        `;
        
        tableBody.appendChild(row);
    });
}

// Fungsi untuk mengisi dropdown minggu untuk kalender
function populateWeekDropdown() {
    const calendarWeek = document.getElementById('calendar-week');
    if (!calendarWeek) return;
    
    // Kosongkan opsi
    calendarWeek.innerHTML = '';
    
    // Kelompokkan berdasarkan minggu
    const weeks = {};
    
    Object.keys(jadwalHariTanggal).forEach(key => {
        const data = jadwalHariTanggal[key];
        const pekan = data.pekan;
        
        if (!weeks[pekan]) {
            weeks[pekan] = [];
        }
        
        weeks[pekan].push({
            key: key,
            data: data
        });
    });
    
    // Tambahkan opsi untuk setiap minggu
    Object.keys(weeks).sort().forEach(pekan => {
        const weekData = weeks[pekan];
        if (weekData.length === 0) return;
        
        const startDate = weekData[0].data.tanggal;
        const endDate = weekData[weekData.length - 1].data.tanggal;
        
        const option = document.createElement('option');
        option.value = pekan;
        option.textContent = `Pekan ${pekan} (${startDate} - ${endDate})`;
        calendarWeek.appendChild(option);
    });
    
    // Default to first week
    if (calendarWeek.options.length > 0) {
        calendarWeek.selectedIndex = 0;
        calendarWeek.dispatchEvent(new Event('change'));
    }
}

// Update tampilan kalender
function updateCalendar() {
    // Reset tampilan kalender
    const calendarContainer = document.querySelector('.calendar-container');
    if (!calendarContainer) return;
    
    calendarContainer.innerHTML = '';
    
    // Pilih dosen dan minggu
    const selectedDosen = document.getElementById('calendar-lecturer');
    const selectedWeek = document.getElementById('calendar-week');
    
    if (!selectedDosen || !selectedWeek) return;
    
    const dosenValue = selectedDosen.value;
    const weekValue = selectedWeek.value;
    
    if (!jadwalHariTanggal || Object.keys(jadwalHariTanggal).length === 0) {
        calendarContainer.innerHTML = '<p>Silakan unggah file Excel dan tentukan tanggal mulai terlebih dahulu.</p>';
        return;
    }
    
    // Kelompokkan berdasarkan pekan
    const weeks = {};
    
    Object.keys(jadwalHariTanggal).forEach(key => {
        const data = jadwalHariTanggal[key];
        const pekan = data.pekan;
        
        if (!weeks[pekan]) {
            weeks[pekan] = [];
        }
        
        weeks[pekan].push({
            key: key,
            data: data
        });
    });
    
    // Sort hari dalam pekan berdasarkan indeks hari
    if (weeks[weekValue]) {
        weeks[weekValue].sort((a, b) => {
            return hariToIndex[a.data.hari] - hariToIndex[b.data.hari];
        });
    }
    
    // Hanya tampilkan pekan yang dipilih
    if (weekValue && weeks[weekValue]) {
        const weekData = weeks[weekValue];
        
        // Buat header untuk pekan ini
        const weekHeader = document.createElement('div');
        weekHeader.className = 'week-header';
        weekHeader.textContent = `Pekan ${weekValue}`;
        calendarContainer.appendChild(weekHeader);
        
        // Buat grid kalender untuk pekan ini
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
            
            // Untuk setiap hari dalam pekan
            weekData.forEach(day => {
                const cellId = `cell-${day.data.hari}-${day.data.tanggal}-${sesi}`;
                const cell = document.createElement('div');
                cell.className = 'calendar-cell status-available';
                cell.id = cellId;
                
                // Tambahkan data atribut
                cell.dataset.hari = day.data.hari;
                cell.dataset.tanggal = day.data.tanggal;
                cell.dataset.pekan = day.data.pekan;
                cell.dataset.sesi = sesi;
                
                // Cek jadwal mengajar pada slot ini
                if (day.data.slots[sesi].jadwalMengajar.length > 0) {
                    let showTeaching = false;
                    
                    if (dosenValue === 'all') {
                        showTeaching = true;
                    } else {
                        // Cek apakah dosen yang dipilih mengajar pada slot ini
                        for (const jadwal of day.data.slots[sesi].jadwalMengajar) {
                            if (jadwal.dosen === dosenValue) {
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
                            if (dosenValue === 'all' || jadwal.dosen === dosenValue) {
                                const eventDiv = document.createElement('div');
                                eventDiv.className = 'calendar-event event-teaching';
                                eventDiv.textContent = `${jadwal.mataKuliah} (${kodeToDosen[jadwal.dosen]}) - ${jadwal.ruangan}`;
                                cell.appendChild(eventDiv);
                            }
                        });
                    }
                }
                
                // Cek jadwal sidang pada slot ini
                if (day.data.slots[sesi].jadwalSidang) {
                    const sidang = day.data.slots[sesi].jadwalSidang;
                    let showSidang = false;
                    
                    if (dosenValue === 'all') {
                        showSidang = true;
                    } else {
                        // Cek apakah dosen yang dipilih terlibat dalam sidang ini
                        for (const dosen of sidang.dosen) {
                            if (dosen === dosenValue) {
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
    const pekanFilter = document.getElementById('filter-pekan').value;
    
    let filteredSlots = [...slotTersedia];
    
    // Filter berdasarkan hari
    if (dayFilter !== 'all') {
        filteredSlots = filteredSlots.filter(slot => slot.Hari === dayFilter);
    }
    
    // Filter berdasarkan sesi
    if (sessionFilter !== 'all') {
        filteredSlots = filteredSlots.filter(slot => slot.Sesi === parseInt(sessionFilter));
    }
    
    // Filter berdasarkan pekan
    if (pekanFilter !== 'all') {
        filteredSlots = filteredSlots.filter(slot => slot.Pekan === parseInt(pekanFilter));
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
        ['Hari', 'Tanggal', 'Pekan', 'Sesi', 'Jam', 'Ketersediaan Dosen']
    ];
    
    slotTersedia.forEach(slot => {
        slotData.push([
            slot.Hari,
            slot.Tanggal,
            slot.Pekan,
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
    // Isi dropdown dosen
    populateDosenDropdowns();
    
    // Download template
    const downloadTemplate = document.getElementById('download-template');
    if (downloadTemplate) {
        downloadTemplate.addEventListener('click', createTemplate);
    }
    
    // File upload
    const fileInput = document.getElementById('file-input');
    if (fileInput) {
        fileInput.addEventListener('change', function(e) {
            if (e.target.files.length) {
                processExcelFile(e.target.files[0]);
            }
        });
    }
    
    // Drag & drop
    const dropArea = document.getElementById('drop-area');
    if (dropArea) {
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
    }
    
    // Export hasil
    const exportResult = document.getElementById('export-result');
    if (exportResult) {
        exportResult.addEventListener('click', exportResults);
    }
    
    // Export slot tersedia
    const exportAvailable = document.getElementById('export-available');
    if (exportAvailable) {
        exportAvailable.addEventListener('click', exportAvailableSlots);
    }
    
    // Filter slot tersedia
    const filterDay = document.getElementById('filter-day');
    const filterSession = document.getElementById('filter-session');
    const filterPekan = document.getElementById('filter-pekan');
    
    if (filterDay && filterSession && filterPekan) {
        filterDay.addEventListener('change', filterAvailableSlots);
        filterSession.addEventListener('change', filterAvailableSlots);
        filterPekan.addEventListener('change', filterAvailableSlots);
    }
    
    // Update kalender saat pilihan dosen atau pekan berubah
    const calendarLecturer = document.getElementById('calendar-lecturer');
    const calendarWeek = document.getElementById('calendar-week');
    
    if (calendarLecturer && calendarWeek) {
        calendarLecturer.addEventListener('change', updateCalendar);
        calendarWeek.addEventListener('change', updateCalendar);
    }
});