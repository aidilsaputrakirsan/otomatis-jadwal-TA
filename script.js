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

// Tambahkan fungsi ini pada bagian awal kode (setelah deklarasi variabel)

// Fungsi untuk mengisi dropdown dosen
function populateDosenDropdowns() {
    const filterLecturer = document.getElementById('filter-lecturer');
    const calendarLecturer = document.getElementById('calendar-lecturer');
    const jadwalMengajarDosen = document.getElementById('jadwal-mengajar-dosen');
    
    if (!filterLecturer || !calendarLecturer || !jadwalMengajarDosen) return;
    
    // Kosongkan opsi kecuali "Semua Dosen"
    while (filterLecturer.options.length > 1) {
        filterLecturer.remove(1);
    }
    
    while (calendarLecturer.options.length > 1) {
        calendarLecturer.remove(1);
    }
    
    while (jadwalMengajarDosen.options.length > 0) {
        jadwalMengajarDosen.remove(0);
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
        
        const option3 = document.createElement('option');
        option3.value = nama;
        option3.textContent = `${nama} (${kode})`;
        jadwalMengajarDosen.appendChild(option3);
    });
}

// Ubah bagian event listener DOMContentLoaded di bagian bawah script

// Event listeners
document.addEventListener('DOMContentLoaded', function() {
    // Isi dropdown dosen saat halaman dimuat
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
    
    // Form jadwal mengajar
    const addJadwalMengajar = document.getElementById('add-jadwal-mengajar');
    if (addJadwalMengajar) {
        addJadwalMengajar.addEventListener('submit', function(e) {
            e.preventDefault();
            
            const dosen = document.getElementById('jadwal-mengajar-dosen').value;
            const hari = document.getElementById('jadwal-mengajar-hari').value;
            const pekan = parseInt(document.getElementById('jadwal-mengajar-pekan').value);
            const sesi = parseInt(document.getElementById('jadwal-mengajar-sesi').value);
            
            addDosenJadwalMengajar(dosen, hari, pekan, sesi);
        });
    }
});

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

// Penjadwalan TA dengan prioritas request
function scheduleTA(jadwalMengajar, timSidang, jadwalHariTanggal) {
    console.log(`Mulai penjadwalan dengan ${timSidang.length} mahasiswa`);
    
    // Hasil jadwal
    const hasilJadwal = new Array(timSidang.length);
    
    // Urutkan tim sidang: prioritaskan yang memiliki request jadwal
    const sortedTimSidang = [...timSidang].sort((a, b) => {
        const aHasRequest = a['Request Tanggal'] && a['Request Sesi'];
        const bHasRequest = b['Request Tanggal'] && b['Request Sesi'];
        
        if (aHasRequest && !bHasRequest) return -1;
        if (!aHasRequest && bHasRequest) return 1;
        return 0;
    });
    
    // Slot yang sudah digunakan untuk jadwal sidang
    const slotSidangTerpakai = {};
    
    // Langkah 1: Jadwalkan semua request terlebih dahulu (MUTLAK)
    console.log("=== Langkah 1: Jadwalkan semua request (prioritas utama) ===");
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
        
        // Jika ada request, jadwalkan tanpa mempertimbangkan bentrok jadwal mengajar
        if (requestTanggal && requestSesi) {
            const requestDate = parseDate(requestTanggal);
            const dayIndex = requestDate.getDay();
            
            // Jika weekend, sesuaikan ke hari kerja terdekat
            let adjustedDayIndex = dayIndex;
            if (dayIndex === 0) adjustedDayIndex = 1; // Minggu -> Senin
            if (dayIndex === 6) adjustedDayIndex = 5; // Sabtu -> Jumat
            
            const requestDay = indexToHari[adjustedDayIndex];
            const formattedRequestDate = formatDate(requestDate);
            const requestKey = `${requestDay}-${formattedRequestDate}`;
            
            // Cek apakah key tersebut ada dalam jadwalHariTanggal
            if (jadwalHariTanggal[requestKey]) {
                // Cek apakah slot sudah digunakan untuk sidang lain
                const slotKey = `${requestKey}-${requestSesi}`;
                if (slotSidangTerpakai[slotKey]) {
                    console.log(`${namaMahasiswa}: Request bentrok dengan sidang lain (${requestDay}, ${formattedRequestDate}, sesi ${requestSesi})`);
                    // Lewati dan akan dijadwalkan pada tahap berikutnya
                    continue;
                }
                
                // Tandai slot sebagai digunakan untuk sidang
                slotSidangTerpakai[slotKey] = true;
                
                // Tandai slot dalam jadwalHariTanggal
                jadwalHariTanggal[requestKey].slots[requestSesi].tersedia = false;
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
                
                console.log(`${namaMahasiswa}: Request dijadwalkan (${requestDay}, ${formattedRequestDate}, sesi ${requestSesi})`);
            } else {
                console.log(`${namaMahasiswa}: Request tanggal tidak dalam rentang penjadwalan (${requestDay}, ${formattedRequestDate})`);
                // Lewati dan akan dijadwalkan pada tahap berikutnya
                continue;
            }
        } else {
            // Lewati yang tidak punya request (akan dijadwalkan pada langkah berikutnya)
            continue;
        }
    }
    
    // Langkah 2: Jadwalkan mahasiswa yang belum terjadwal
    console.log("=== Langkah 2: Jadwalkan mahasiswa yang belum terjadwal ===");
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
                    // Jika slot belum digunakan untuk sidang
                    if (!slotSidangTerpakai[slotKey]) {
                        // Tandai slot ini sebagai digunakan
                        slotSidangTerpakai[slotKey] = true;
                        
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
        
        if (slotDitemukan) {
            // Tandai slot dalam jadwalHariTanggal
            jadwalHariTanggal[hariTanggalKey].slots[sesi].tersedia = false;
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
    console.log(`Penjadwalan selesai: ${mahasiswaTerjadwal}/${totalMahasiswa} mahasiswa berhasil dijadwalkan (${Math.round(mahasiswaTerjadwal/totalMahasiswa*100)}%)`);
    
    return hasilJadwal;
}

// Fungsi untuk menghasilkan slot tersedia
function generateAvailableSlots(jadwalMengajar, hasilJadwal, jadwalHariTanggal) {
    const slotTersedia = [];
    
    // Set untuk menyimpan slot yang sudah digunakan untuk sidang
    const slotSidangTerpakai = new Set();
    
    // Tandai slot yang sudah digunakan untuk sidang
    hasilJadwal.forEach(jadwal => {
        if (jadwal.Hari !== 'Tidak ada slot tersedia') {
            const key = `${jadwal.Hari}-${jadwal.Tanggal}-${jadwal.Sesi}`;
            slotSidangTerpakai.add(key);
        }
    });
    
    // Iterasi semua slot dari awal hingga akhir periode
    Object.keys(jadwalHariTanggal).forEach(hariTanggalKey => {
        const hariTanggal = jadwalHariTanggal[hariTanggalKey];
        
        // Iterasi semua sesi dalam hari ini
        for (let sesi = 1; sesi <= 4; sesi++) {
            // Lewati slot yang sudah digunakan untuk sidang
            const key = `${hariTanggal.hari}-${hariTanggal.tanggal}-${sesi}`;
            if (slotSidangTerpakai.has(key)) {
                continue;
            }
            
            // Tambahkan ke daftar slot tersedia
            slotTersedia.push({
                'Hari': hariTanggal.hari,
                'Tanggal': hariTanggal.tanggal,
                'Pekan': hariTanggal.pekan,
                'Sesi': sesi,
                'Jam': jadwalJam[sesi],
                'Ketersediaan Dosen': 'Semua Dosen' // Untuk sederhananya, anggap semua dosen tersedia
            });
        }
    });
    
    return slotTersedia;
}

// Fungsi untuk membuat template Excel
function createTemplate() {
    // Buat workbook baru
    const wb = XLSX.utils.book_new();
    
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

// Variabel untuk jadwal mengajar dosen
let dosenJadwalMengajar = {};

// Function untuk menambahkan jadwal mengajar dosen
function addDosenJadwalMengajar(dosen, hari, pekan, sesi) {
    if (!dosenJadwalMengajar[dosen]) {
        dosenJadwalMengajar[dosen] = [];
    }
    
    dosenJadwalMengajar[dosen].push({
        hari: hari,
        pekan: pekan,
        sesi: sesi
    });
    
    console.log(`Jadwal mengajar ditambahkan: ${dosen} - ${hari}, Pekan ${pekan}, Sesi ${sesi}`);
    
    // Update tampilan kalender
    updateCalendar();
}

// Function untuk menghapus jadwal mengajar dosen
function removeDosenJadwalMengajar(dosen, hari, pekan, sesi) {
    if (!dosenJadwalMengajar[dosen]) {
        return;
    }
    
    // Cari indeks jadwal yang akan dihapus
    const index = dosenJadwalMengajar[dosen].findIndex(j => 
        j.hari === hari && j.pekan === parseInt(pekan) && j.sesi === parseInt(sesi)
    );
    
    if (index !== -1) {
        dosenJadwalMengajar[dosen].splice(index, 1);
        console.log(`Jadwal mengajar dihapus: ${dosen} - ${hari}, Pekan ${pekan}, Sesi ${sesi}`);
        
        // Update tampilan kalender
        updateCalendar();
    }
}

// Parsing data Excel
function processExcelFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        
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

// Fungsi untuk mengisi dropdown dosen
function populateDosenDropdowns() {
    const filterLecturer = document.getElementById('filter-lecturer');
    const calendarLecturer = document.getElementById('calendar-lecturer');
    const jadwalMengajarDosen = document.getElementById('jadwal-mengajar-dosen');
    
    if (!filterLecturer || !calendarLecturer || !jadwalMengajarDosen) return;
    
    // Kosongkan opsi kecuali "Semua Dosen"
    while (filterLecturer.options.length > 1) {
        filterLecturer.remove(1);
    }
    
    while (calendarLecturer.options.length > 1) {
        calendarLecturer.remove(1);
    }
    
    while (jadwalMengajarDosen.options.length > 0) {
        jadwalMengajarDosen.remove(0);
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
        
        const option3 = document.createElement('option');
        option3.value = nama;
        option3.textContent = `${nama} (${kode})`;
        jadwalMengajarDosen.appendChild(option3);
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
                
                // Tambahkan event click untuk mengedit jadwal mengajar
                cell.addEventListener('click', function() {
                    handleCalendarCellClick(this);
                });
                
                // Cek jadwal mengajar dosen pada slot ini
                if (dosenValue !== 'all' && dosenJadwalMengajar[dosenValue]) {
                    const jadwalDosen = dosenJadwalMengajar[dosenValue];
                    const isTeaching = jadwalDosen.some(jadwal => 
                        jadwal.hari === day.data.hari && 
                        jadwal.pekan === parseInt(day.data.pekan) && 
                        jadwal.sesi === parseInt(sesi)
                    );
                    
                    if (isTeaching) {
                        cell.classList.remove('status-available');
                        cell.classList.add('status-teaching');
                        
                        const eventDiv = document.createElement('div');
                        eventDiv.className = 'calendar-event event-teaching';
                        eventDiv.textContent = `${dosenValue} Mengajar`;
                        cell.appendChild(eventDiv);
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

// Menangani klik pada sel kalender
function handleCalendarCellClick(cell) {
    const hari = cell.dataset.hari;
    const pekan = parseInt(cell.dataset.pekan);
    const sesi = parseInt(cell.dataset.sesi);
    const selectedDosen = document.getElementById('calendar-lecturer').value;
    
    // Jika semua dosen, tampilkan dialog untuk memilih dosen
    if (selectedDosen === 'all') {
        alert('Silakan pilih dosen terlebih dahulu untuk mengedit jadwal mengajar.');
        return;
    }
    
    // Cek apakah dosen sudah memiliki jadwal mengajar pada slot ini
    let isTeaching = false;
    if (dosenJadwalMengajar[selectedDosen]) {
        isTeaching = dosenJadwalMengajar[selectedDosen].some(jadwal => 
            jadwal.hari === hari && 
            jadwal.pekan === pekan && 
            jadwal.sesi === sesi
        );
    }
    
    // Toggle jadwal mengajar
    if (isTeaching) {
        if (confirm(`Hapus jadwal mengajar ${selectedDosen} pada ${hari}, Pekan ${pekan}, Sesi ${sesi}?`)) {
            removeDosenJadwalMengajar(selectedDosen, hari, pekan, sesi);
        }
    } else {
        if (confirm(`Tambahkan jadwal mengajar ${selectedDosen} pada ${hari}, Pekan ${pekan}, Sesi ${sesi}?`)) {
            addDosenJadwalMengajar(selectedDosen, hari, pekan, sesi);
        }
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
    
    // Form jadwal mengajar
    const addJadwalMengajar = document.getElementById('add-jadwal-mengajar');
    if (addJadwalMengajar) {
        addJadwalMengajar.addEventListener('submit', function(e) {
            e.preventDefault();
            
            const dosen = document.getElementById('jadwal-mengajar-dosen').value;
            const hari = document.getElementById('jadwal-mengajar-hari').value;
            const pekan = parseInt(document.getElementById('jadwal-mengajar-pekan').value);
            const sesi = parseInt(document.getElementById('jadwal-mengajar-sesi').value);
            
            addDosenJadwalMengajar(dosen, hari, pekan, sesi);
        });
    }
});