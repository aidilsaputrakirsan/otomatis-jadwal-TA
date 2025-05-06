// script.js
document.addEventListener('DOMContentLoaded', function() {
    // Konfigurasi URL Google Apps Script Web App
    const scriptURL = 'https://script.google.com/macros/s/AKfycbzzyHJlllUli-SAcKbqdf7QpKFG_pipx44-jtaly87iSA77aUkMPmEDA3ThJXptZg0/exec'; // Ganti dengan URL GAS Web App Anda
    
    // Inisialisasi elemen-elemen UI
    initTabs();
    initButtons();
    
    // Fungsi untuk mengatur tab
    function initTabs() {
        const tabButtons = document.querySelectorAll('.tab-button');
        
        tabButtons.forEach(button => {
            button.addEventListener('click', function() {
                // Nonaktifkan semua tab
                tabButtons.forEach(btn => {
                    btn.classList.remove('active');
                });
                
                // Sembunyikan semua konten tab
                document.querySelectorAll('.tab-pane').forEach(pane => {
                    pane.classList.remove('active');
                });
                
                // Aktifkan tab yang dipilih
                this.classList.add('active');
                
                // Tampilkan konten tab yang dipilih
                const tabId = this.getAttribute('data-tab');
                document.getElementById(tabId).classList.add('active');
                
                // Sembunyikan hasil sebelumnya
                hideResults();
            });
        });
    }
    
    // Fungsi untuk mengatur tombol
    function initButtons() {
        // Tombol Cari Jadwal Kosong
        document.getElementById('cariJadwalBtn').addEventListener('click', function() {
            const namaDosen = document.getElementById('cariDosenInput').value.trim();
            
            if (!namaDosen) {
                showResult('cariJadwalResult', 'Silakan masukkan nama dosen terlebih dahulu.', false);
                return;
            }
            
            showLoader();
            
            // Panggil API untuk mencari jadwal kosong
            fetch(`${scriptURL}?action=cariJadwalKosongWeb&namaDosen=${encodeURIComponent(namaDosen)}`)
                .then(response => response.json())
                .then(data => {
                    hideLoader();
                    showResult('cariJadwalResult', data.message, data.success);
                })
                .catch(error => {
                    hideLoader();
                    showResult('cariJadwalResult', `Error: ${error.message}`, false);
                });
        });
        
        // Tombol Cek Jadwal Dosen
        document.getElementById('cekJadwalBtn').addEventListener('click', function() {
            const namaDosen = document.getElementById('cekDosenInput').value.trim();
            
            if (!namaDosen) {
                showResult('cekJadwalResult', 'Silakan masukkan nama dosen terlebih dahulu.', false);
                return;
            }
            
            showLoader();
            
            // Panggil API untuk cek jadwal dosen
            fetch(`${scriptURL}?action=cekJadwalDosenWeb&namaDosen=${encodeURIComponent(namaDosen)}`)
                .then(response => response.json())
                .then(data => {
                    hideLoader();
                    showResult('cekJadwalResult', data.message, data.success);
                })
                .catch(error => {
                    hideLoader();
                    showResult('cekJadwalResult', `Error: ${error.message}`, false);
                });
        });
        
        // Tombol Cek Ketersediaan Sidang
        document.getElementById('cekKetersediaanBtn').addEventListener('click', function() {
            const nomorSidang = document.getElementById('nomorSidangInput').value.trim();
            
            if (!nomorSidang) {
                showResult('jadwalkanResult', 'Silakan masukkan nomor sidang terlebih dahulu.', false);
                return;
            }
            
            showLoader();
            
            // Panggil API untuk persiapan jadwal sidang
            fetch(`${scriptURL}?action=prepareJadwalSidangWeb&nomorSidang=${encodeURIComponent(nomorSidang)}`)
                .then(response => response.json())
                .then(data => {
                    hideLoader();
                    
                    if (data.success) {
                        // Tampilkan detail sidang
                        const detailSidang = document.getElementById('detailSidang');
                        detailSidang.classList.remove('hidden');
                        
                        // Isi konten detail sidang
                        const detailHTML = `
                            <p><strong>Mahasiswa:</strong> ${data.sidang.mahasiswa}</p>
                            <p><strong>Judul:</strong> ${data.sidang.judul}</p>
                            <p><strong>Komite:</strong></p>
                            <ul>
                                <li><strong>Pembimbing 1:</strong> ${data.sidang.pembimbing1 || '-'}</li>
                                <li><strong>Pembimbing 2:</strong> ${data.sidang.pembimbing2 || '-'}</li>
                                <li><strong>Penguji 1:</strong> ${data.sidang.penguji1 || '-'}</li>
                                <li><strong>Penguji 2:</strong> ${data.sidang.penguji2 || '-'}</li>
                            </ul>
                        `;
                        
                        document.getElementById('detailSidangContent').innerHTML = detailHTML;
                        
                        // Isi dropdown slot waktu
                        const slotSelect = document.getElementById('slotSelect');
                        slotSelect.innerHTML = '';
                        
                        if (data.slotKosong.length === 0) {
                            slotSelect.innerHTML = '<option value="">Tidak ada slot kosong tersedia</option>';
                            document.getElementById('ruanganInput').disabled = true;
                            document.getElementById('jadwalkanBtn').disabled = true;
                            
                            showResult('jadwalkanResult', 'Tidak ada slot yang tersedia untuk semua dosen komite. Silakan pilih dosen pengganti atau sesuaikan jadwal dosen.', false);
                        } else {
                            data.slotKosong.forEach(function(slot, index) {
                                const option = document.createElement('option');
                                option.value = index;
                                option.textContent = `${slot.tanggal} - ${slot.waktu}`;
                                slotSelect.appendChild(option);
                            });
                            
                            document.getElementById('ruanganInput').disabled = false;
                            document.getElementById('jadwalkanBtn').disabled = false;
                            
                            showResult('jadwalkanResult', 'Berhasil memuat data sidang dan slot waktu yang tersedia.', true);
                        }
                    } else {
                        document.getElementById('detailSidang').classList.add('hidden');
                        showResult('jadwalkanResult', data.message, false);
                    }
                })
                .catch(error => {
                    hideLoader();
                    document.getElementById('detailSidang').classList.add('hidden');
                    showResult('jadwalkanResult', `Error: ${error.message}`, false);
                });
        });
        
        // Tombol Jadwalkan Sidang
        document.getElementById('jadwalkanBtn').addEventListener('click', function() {
            const nomorSidang = document.getElementById('nomorSidangInput').value.trim();
            const slotIndex = document.getElementById('slotSelect').value;
            const ruangan = document.getElementById('ruanganInput').value.trim();
            
            if (!nomorSidang || !ruangan) {
                showResult('jadwalkanResult', 'Silakan lengkapi semua data terlebih dahulu.', false);
                return;
            }
            
            showLoader();
            
            // Panggil API untuk jadwalkan sidang
            fetch(`${scriptURL}?action=jadwalkanSidangWeb&nomorSidang=${encodeURIComponent(nomorSidang)}&slotIndex=${encodeURIComponent(slotIndex)}&ruangan=${encodeURIComponent(ruangan)}`)
                .then(response => response.json())
                .then(data => {
                    hideLoader();
                    
                    if (data.success) {
                        // Reset form jadwal detail
                        document.getElementById('detailSidang').classList.add('hidden');
                        document.getElementById('nomorSidangInput').value = '';
                        document.getElementById('ruanganInput').value = '';
                    }
                    
                    showResult('jadwalkanResult', data.message, data.success);
                })
                .catch(error => {
                    hideLoader();
                    showResult('jadwalkanResult', `Error: ${error.message}`, false);
                });
        });
    }
    
    // Helper functions
    function showLoader() {
        document.getElementById('loader').classList.remove('hidden');
    }
    
    function hideLoader() {
        document.getElementById('loader').classList.add('hidden');
    }
    
    function showResult(resultId, message, isSuccess) {
        const resultElement = document.getElementById(resultId);
        const resultContent = resultElement.querySelector('.result-content');
        
        resultElement.style.display = 'block';
        resultContent.innerHTML = message;
        
        if (isSuccess) {
            resultContent.classList.add('success');
            resultContent.classList.remove('error');
        } else {
            resultContent.classList.add('error');
            resultContent.classList.remove('success');
        }
        
        // Scroll to result
        resultElement.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
    
    function hideResults() {
        document.querySelectorAll('.result').forEach(result => {
            result.style.display = 'none';
        });
    }
});