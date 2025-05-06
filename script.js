// Daftar kode dosen dan nama lengkap (hardcoded)
const daftarDosen = {
    'DNA': 'Dwi Nur Amalia, S.Kom., M.Kom',
    'YTW': 'Yuyun Tri Wiranti, S.Kom., M.MT',
    'PDA': 'Ir. I Putu Deny Arthawan Sugih Prabowo, M.Eng',
    'VKA': 'Vika Fitratunnany Insanittaqwa, S.Kom., M.Kom',
    'HLH': 'Henokh Lugo Hariyanto, S.Si., M.Sc.',
    'MIA': 'M. Ihsan Alfani Putera, S.T.,Kom, M.Kom',
    'DAP': 'Dwi Arief Prambudi, S.Kom., M.Kom.',
    'NNA': 'Nursanti Novi Arisa, S.Pd., M.Kom.',
    'ADL': 'Aidi Saputra Kirsan, S.ST., M.T.,Kom.',
    'HIS': 'Hendy Indrawan Sunardi, S.Kom, M.Eng.',
    'AWS': 'Arif Wicaksono, M.Kom',
    'SRN': 'Sri Rahayu Natasia, S.Komp., M.Si., M.Sc.'
};

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

// Menyimpan jadwal
let jadwalMengajar = [];
let timSidang = [];
let hasilJadwal = [];
let slotTersedia = [];

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
    
    // Template tim sidang
    const timData = [
        ['Nama Mahasiswa / NIM', 'Pembimbing 1', 'Pembimbing 2', 'Penguji 1', 'Penguji 2'],
        ['Athifah Shyla Maritza / 10211021', 'Dwi Nur Amalia, S.Kom., M.Kom', 'Yuyun Tri Wiranti, S.Kom., M.MT', 'Ir. I Putu Deny Arthawan Sugih Prabowo, M.Eng', 'Vika Fitratunnany Insanittaqwa, S.Kom., M.Kom'],
        ['Muhammad Rizky Afriza / 10211022', 'Yuyun Tri Wiranti, S.Kom., M.MT', 'Henokh Lugo Hariyanto, S.Si., M.Sc.', 'M. Ihsan Alfani Putera, S.T.,Kom, M.Kom', 'Dwi Arief Prambudi, S.Kom., M.Kom.']
    ];
    const timSheet = XLSX.utils.aoa_to_sheet(timData);
    XLSX.utils.book_append_sheet(wb, timSheet, 'TimSidang');
    
    // Download file
    XLSX.writeFile(wb, 'template_sidang_ta.xlsx');
}

// Algoritma penjadwalan
function scheduleTA(jadwalMengajar, timSidang) {
    const hasilJadwal = [];
    const hariList = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'];
    const sesiList = [1, 2, 3, 4];
    
    // Jadwal yang sudah digunakan oleh dosen
    const jadwalDosenTerpakai = {};
    
    // Buat pemetaan nama lengkap dosen ke jadwal mengajar mereka
    const dosenToJadwal = {};
    
    jadwalMengajar.forEach(jadwal => {
        const kodeDosen = jadwal['Kode Dosen'];
        const namaLengkap = daftarDosen[kodeDosen];
        
        if (!dosenToJadwal[namaLengkap]) {
            dosenToJadwal[namaLengkap] = [];
        }
        
        dosenToJadwal[namaLengkap].push({
            hari: jadwal['Hari'],
            sesi: jadwal['Sesi']
        });
    });
    
    // Untuk setiap tim sidang
    timSidang.forEach((tim, idx) => {
        const namaMahasiswa = tim['Nama Mahasiswa / NIM'];
        const pembimbing1 = tim['Pembimbing 1'];
        const pembimbing2 = tim['Pembimbing 2'];
        const penguji1 = tim['Penguji 1'];
        const penguji2 = tim['Penguji 2'];
        
        const dosenTim = [pembimbing1, pembimbing2, penguji1, penguji2];
        
        // Cari slot waktu yang tersedia
        const slotTersedia = [];
        
        hariList.forEach(hari => {
            sesiList.forEach(sesi => {
                // Cek apakah ada dosen yang mengajar pada slot ini
                let adaYangMengajar = false;
                
                for (const dosen of dosenTim) {
                    if (dosenToJadwal[dosen]) {
                        const jadwalDosen = dosenToJadwal[dosen].filter(j => 
                            j.hari === hari && j.sesi === sesi
                        );
                        
                        if (jadwalDosen.length > 0) {
                            adaYangMengajar = true;
                            break;
                        }
                    }
                }
                
                // Cek apakah ada dosen yang sudah terjadwal sidang pada slot ini
                let adaYangSidang = false;
                for (const dosen of dosenTim) {
                    if (jadwalDosenTerpakai[dosen]) {
                        for (const jadwal of jadwalDosenTerpakai[dosen]) {
                            if (jadwal[0] === hari && jadwal[1] === sesi) {
                                adaYangSidang = true;
                                break;
                            }
                        }
                    }
                    if (adaYangSidang) break;
                }
                
                if (!adaYangMengajar && !adaYangSidang) {
                    slotTersedia.push([hari, sesi]);
                }
            });
        });
        
        // Pilih slot pertama yang tersedia
        if (slotTersedia.length > 0) {
            const slotTerpilih = slotTersedia[0];
            
            // Update jadwal dosen terpakai
            dosenTim.forEach(dosen => {
                if (!jadwalDosenTerpakai[dosen]) {
                    jadwalDosenTerpakai[dosen] = [];
                }
                jadwalDosenTerpakai[dosen].push(slotTerpilih);
            });
            
            hasilJadwal.push({
                'Nama Mahasiswa': namaMahasiswa,
                'Pembimbing 1': pembimbing1,
                'Pembimbing 2': pembimbing2,
                'Penguji 1': penguji1,
                'Penguji 2': penguji2,
                'Hari': slotTerpilih[0],
                'Sesi': slotTerpilih[1],
                'Jam': jadwalJam[slotTerpilih[1]]
            });
        } else {
            hasilJadwal.push({
                'Nama Mahasiswa': namaMahasiswa,
                'Pembimbing 1': pembimbing1,
                'Pembimbing 2': pembimbing2,
                'Penguji 1': penguji1,
                'Penguji 2': penguji2,
                'Hari': 'Tidak ada slot tersedia',
                'Sesi': '-',
                'Jam': '-'
            });
        }
    });
    
    return hasilJadwal;
}

// Fungsi untuk menghasilkan slot tersedia
function generateAvailableSlots(jadwalMengajar, hasilJadwal) {
    const hariList = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'];
    const sesiList = [1, 2, 3, 4];
    
    // Menyimpan jadwal dosen (mengajar dan sidang)
    const jadwalDosen = {};
    
    // Simpan jadwal mengajar
    jadwalMengajar.forEach(jadwal => {
        const kodeDosen = jadwal['Kode Dosen'];
        const namaLengkap = daftarDosen[kodeDosen];
        
        if (!jadwalDosen[namaLengkap]) {
            jadwalDosen[namaLengkap] = [];
        }
        
        jadwalDosen[namaLengkap].push({
            hari: jadwal['Hari'],
            sesi: jadwal['Sesi'],
            tipe: 'mengajar'
        });
    });
    
    // Simpan jadwal sidang
    hasilJadwal.forEach(jadwal => {
        if (jadwal['Hari'] !== 'Tidak ada slot tersedia') {
            const dosen = [
                jadwal['Pembimbing 1'],
                jadwal['Pembimbing 2'],
                jadwal['Penguji 1'],
                jadwal['Penguji 2']
            ];
            
            dosen.forEach(d => {
                if (!jadwalDosen[d]) {
                    jadwalDosen[d] = [];
                }
                
                jadwalDosen[d].push({
                    hari: jadwal['Hari'],
                    sesi: jadwal['Sesi'],
                    tipe: 'sidang'
                });
            });
        }
    });
    
    // Generate slot tersedia - dikelompokkan berdasarkan hari dan sesi saja
    const slotTersedia = [];
    const slotDosenTersedia = {};
    
    hariList.forEach(hari => {
        sesiList.forEach(sesi => {
            // Cek apakah slot sedang digunakan untuk sidang
            let slotTerpakai = false;
            for (const jadwal of hasilJadwal) {
                if (jadwal['Hari'] === hari && jadwal['Sesi'] === sesi) {
                    slotTerpakai = true;
                    break;
                }
            }
            
            if (!slotTerpakai) {
                // Temukan dosen yang tersedia pada slot ini
                const dosenTersedia = [];
                
                Object.keys(daftarDosen).forEach(kode => {
                    const namaLengkap = daftarDosen[kode];
                    let dosenSibuk = false;
                    
                    if (jadwalDosen[namaLengkap]) {
                        for (const jadwal of jadwalDosen[namaLengkap]) {
                            if (jadwal.hari === hari && jadwal.sesi === sesi) {
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
                    const slotKey = `${hari}-${sesi}`;
                    slotDosenTersedia[slotKey] = dosenTersedia;
                    
                    slotTersedia.push({
                        'Hari': hari,
                        'Sesi': sesi,
                        'Jam': jadwalJam[sesi],
                        'Ketersediaan Dosen': dosenTersedia.join(', ')
                    });
                }
            }
        });
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
                    item[timHeaders[j]] = row[j];
                }
                timSidang.push(item);
            }
        }
        
        // Lakukan penjadwalan
        hasilJadwal = scheduleTA(jadwalMengajar, timSidang);
        
        // Generate slot tersedia
        slotTersedia = generateAvailableSlots(jadwalMengajar, hasilJadwal);
        
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
        
        row.innerHTML = `
            <td>${mahasiswa}</td>
            <td>${result['Pembimbing 1']}</td>
            <td>${result['Pembimbing 2']}</td>
            <td>${result['Penguji 1']}</td>
            <td>${result['Penguji 2']}</td>
            <td>${result.Hari}</td>
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
            <td>${slot.Sesi}</td>
            <td>${slot.Jam}</td>
            <td>${slot.Ruangan}</td>
            <td>${slot['Ketersediaan Dosen']}</td>
        `;
        
        tableBody.appendChild(row);
    });
}

// Isi dropdown filter dosen
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

// Update tampilan kalender
function updateCalendar() {
    // Reset semua cell
    const cells = document.querySelectorAll('.calendar-cell');
    cells.forEach(cell => {
        cell.className = 'calendar-cell status-available';
        cell.innerHTML = '';
    });
    
    // Pilih dosen
    const selectedDosen = document.getElementById('calendar-lecturer').value;
    
    // Isi kalender dengan jadwal mengajar
    jadwalMengajar.forEach(jadwal => {
        const hari = hariToIndex[jadwal.Hari];
        const sesi = jadwal.Sesi;
        const kodeDosen = jadwal['Kode Dosen'];
        const namaLengkap = daftarDosen[kodeDosen];
        
        if (selectedDosen === 'all' || selectedDosen === namaLengkap) {
            const cell = document.getElementById(`cell-${hari}-${sesi}`);
            
            // Tambahkan info jadwal mengajar
            const eventDiv = document.createElement('div');
            eventDiv.className = 'calendar-event event-teaching';
            eventDiv.textContent = `${jadwal['Mata Kuliah']} (${kodeDosen}) - ${jadwal.Ruangan}`;
            
            cell.classList.remove('status-available');
            cell.classList.add('status-teaching');
            cell.appendChild(eventDiv);
        }
    });
    
    // Isi kalender dengan jadwal sidang
    hasilJadwal.forEach(jadwal => {
        if (jadwal.Hari !== 'Tidak ada slot tersedia') {
            const hari = hariToIndex[jadwal.Hari];
            const sesi = jadwal.Sesi;
            const mahasiswa = jadwal['Nama Mahasiswa'].split(' / ')[0];
            
            const dosen = [
                jadwal['Pembimbing 1'],
                jadwal['Pembimbing 2'],
                jadwal['Penguji 1'],
                jadwal['Penguji 2']
            ];
            
            if (selectedDosen === 'all' || dosen.includes(selectedDosen)) {
                const cell = document.getElementById(`cell-${hari}-${sesi}`);
                
                // Tambahkan info jadwal sidang
                const eventDiv = document.createElement('div');
                eventDiv.className = 'calendar-event event-defense';
                eventDiv.textContent = `Sidang TA: ${mahasiswa}`;
                
                cell.classList.remove('status-available');
                cell.classList.remove('status-teaching');
                cell.classList.add('status-defense');
                cell.appendChild(eventDiv);
            }
        }
    });
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
        ['Nama Mahasiswa', 'Pembimbing 1', 'Pembimbing 2', 'Penguji 1', 'Penguji 2', 'Hari', 'Sesi', 'Jam']
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
            jadwal.Sesi,
            jadwal.Jam
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
        ['Hari', 'Sesi', 'Jam', 'Ruangan', 'Ketersediaan Dosen']
    ];
    
    slotTersedia.forEach(slot => {
        slotData.push([
            slot.Hari,
            slot.Sesi,
            slot.Jam,
            slot.Ruangan,
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
    
    // Update kalender saat pilihan dosen berubah
    document.getElementById('calendar-lecturer').addEventListener('change', updateCalendar);
});