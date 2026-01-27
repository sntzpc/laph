// Inisialisasi tanggal default
document.addEventListener('DOMContentLoaded', function() {
    // Set tanggal default (hari ini dan 7 hari sebelumnya)
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 7);

    document.getElementById('startDate').valueAsDate = startDate;
    document.getElementById('endDate').valueAsDate = endDate;

    setupEventListeners();

    console.log('Aplikasi Laporan Harian dimuat');
    console.log('URL GAS:', GAS_URL);
});

// URL Google Apps Script Web App (/exec)
const GAS_URL = 'https://script.google.com/macros/s/AKfycbyt08ZffVLaFUgZLA4hveiIKOufJ1FeZbTWuGijWypdjhciOx5sMl9duHWh9-b0XhEV/exec';

// Setup semua event listeners
function setupEventListeners() {
    document.getElementById('fetchDataBtn').addEventListener('click', fetchData);
    document.getElementById('exportPusinganBtn').addEventListener('click', () => exportData('pusingan'));
    document.getElementById('exportRkhBtn').addEventListener('click', () => exportData('rkh'));
    document.getElementById('exportAllBtn').addEventListener('click', exportAllData);
}

function showLoading() {
    document.getElementById('loadingModal').classList.remove('hidden');
}
function hideLoading() {
    document.getElementById('loadingModal').classList.add('hidden');
}

function disableButtons() {
    const buttons = ['fetchDataBtn', 'exportPusinganBtn', 'exportRkhBtn', 'exportAllBtn'];
    buttons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
            btn.disabled = true;
            btn.classList.add('opacity-50', 'cursor-not-allowed');
        }
    });
}
function enableButtons() {
    const buttons = ['fetchDataBtn', 'exportPusinganBtn', 'exportRkhBtn', 'exportAllBtn'];
    buttons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
            btn.disabled = false;
            btn.classList.remove('opacity-50', 'cursor-not-allowed');
        }
    });
}

/** ==========================================================
 * JSONP helper (CORS-safe for localhost/GitHub Pages)
 * Backend wajib support param: callback=fn
 * ========================================================== */
function jsonpGet(url, { timeoutMs = 20000 } = {}) {
    return new Promise((resolve, reject) => {
        const cbName = `__cb_${Date.now()}_${Math.random().toString(16).slice(2)}`;
        const script = document.createElement('script');

        let done = false;
        const cleanup = () => {
            if (script && script.parentNode) script.parentNode.removeChild(script);
            try { delete window[cbName]; } catch (e) { window[cbName] = undefined; }
        };

        const timer = setTimeout(() => {
            if (done) return;
            done = true;
            cleanup();
            reject(new Error('JSONP timeout'));
        }, timeoutMs);

        window[cbName] = (data) => {
            if (done) return;
            done = true;
            clearTimeout(timer);
            cleanup();
            resolve(data);
        };

        script.onerror = () => {
            if (done) return;
            done = true;
            clearTimeout(timer);
            cleanup();
            reject(new Error('JSONP load error'));
        };

        const sep = url.includes('?') ? '&' : '?';
        script.src = `${url}${sep}callback=${encodeURIComponent(cbName)}`;
        document.head.appendChild(script);
    });
}

// Fungsi untuk mengambil data
async function fetchData() {
    const startDate = document.getElementById('startDate').value;
    const endDate = document.getElementById('endDate').value;

    if (!startDate || !endDate) {
        alert('Harap pilih tanggal mulai dan tanggal akhir');
        return;
    }
    if (new Date(startDate) > new Date(endDate)) {
        alert('Tanggal mulai tidak boleh lebih besar dari tanggal akhir');
        return;
    }

    showLoading();
    disableButtons();

    try {
        console.log('Mengambil data:', { startDate, endDate });

        // PUSINGAN (JSONP)
        const pusinganUrl = `${GAS_URL}?action=getPusingan&startDate=${encodeURIComponent(startDate)}&endDate=${encodeURIComponent(endDate)}`;
        const pusinganData = await jsonpGet(pusinganUrl);

        // RKH (JSONP)
        const rkhUrl = `${GAS_URL}?action=getRKH&startDate=${encodeURIComponent(startDate)}&endDate=${encodeURIComponent(endDate)}`;
        const rkhData = await jsonpGet(rkhUrl);

        // Render pusingan
        if (pusinganData && pusinganData.success) {
            displayData('pusingan', pusinganData.data, pusinganData.headers);
            document.getElementById('pusinganCount').textContent = pusinganData.count ?? (pusinganData.data ? pusinganData.data.length : 0);
            document.getElementById('exportPusinganBtn').disabled = !pusinganData.data || pusinganData.data.length === 0;
        } else {
            throw new Error('Error mengambil data pusingan: ' + (pusinganData?.error || 'Unknown'));
        }

        // Render rkh
        if (rkhData && rkhData.success) {
            displayData('rkh', rkhData.data, rkhData.headers);
            document.getElementById('rkhCount').textContent = rkhData.count ?? (rkhData.data ? rkhData.data.length : 0);
            document.getElementById('exportRkhBtn').disabled = !rkhData.data || rkhData.data.length === 0;
        } else {
            throw new Error('Error mengambil data RKH: ' + (rkhData?.error || 'Unknown'));
        }

        const total = (pusinganData?.data?.length || 0) + (rkhData?.data?.length || 0);
        document.getElementById('exportAllBtn').disabled = total === 0;

    } catch (error) {
        console.error('fetchData error:', error);
        alert('Terjadi kesalahan saat mengambil data: ' + error.message);
    } finally {
        hideLoading();
        enableButtons();
    }
}

// Fungsi untuk menampilkan data di tabel
function displayData(type, data, headers) {
    const tableBody = document.getElementById(`${type}TableBody`);

    const colCount = headers?.length || (type === 'rkh' ? 7 : 6);

    if (!data || data.length === 0) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="${colCount}" class="px-3 py-8 text-center text-gray-500">
                    <i class="fas fa-search text-3xl text-gray-300 mb-2 block"></i>
                    <p>Tidak ada data ditemukan untuk periode ini</p>
                </td>
            </tr>
        `;
        return;
    }

    let html = '';
    data.forEach((row, index) => {
        const fadeClass = index < 10 ? 'fade-in' : '';
        html += `<tr class="${fadeClass}" style="animation-delay: ${index * 0.05}s;">`;

        headers.forEach(header => {
            const value = (row && row[header] !== undefined && row[header] !== null) ? String(row[header]) : '';

            if (header === 'nilai') {
                const nilaiClass = value ? `nilai-${value}` : '';
                html += `<td class="px-3 py-2 whitespace-nowrap ${nilaiClass} text-center font-medium">${value}</td>`;
            } else {
                html += `<td class="px-3 py-2 whitespace-nowrap">${value}</td>`;
            }
        });

        html += '</tr>';
    });

    tableBody.innerHTML = html;
}

// Export data ke Excel
function exportData(type) {
    const tableBody = document.getElementById(`${type}TableBody`);
    const rows = tableBody.querySelectorAll('tr');

    // cek placeholder
    if (rows.length === 1 && rows[0].querySelector('td[colspan]')) {
        alert(`Tidak ada data ${type} untuk di-export`);
        return;
    }

    const wsData = [];

    let headers;
    if (type === 'rkh') {
        headers = ['DIVISI_ID', 'NIK', 'REPORT_DATE', 'SEND_DATE', 'SEND_TIME', 'NILAI', 'NAMA'];
    } else {
        headers = ['NIK', 'REPORT_DATE', 'SEND_DATE', 'SEND_TIME', 'NILAI', 'NAMA'];
    }
    wsData.push(headers);

    rows.forEach(row => {
        const cols = row.querySelectorAll('td');
        const expected = (type === 'rkh') ? 7 : 6;
        if (cols.length === expected) {
            wsData.push(Array.from(cols).map(col => col.textContent));
        }
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, type === 'pusingan' ? 'Data_Pusingan' : 'Data_RKH');

    const fileName = `Laporan_Harian_${type}_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
}

function exportAllData() {
    const wb = XLSX.utils.book_new();

    // Pusingan
    appendSheetFromTable_({
        wb,
        sheetName: 'Data_Pusingan',
        tableBodyId: 'pusinganTableBody',
        headerRow: ['NIK', 'REPORT_DATE', 'SEND_DATE', 'SEND_TIME', 'NILAI', 'NAMA'],
        expectedCols: 6
    });

    // RKH
    appendSheetFromTable_({
        wb,
        sheetName: 'Data_RKH',
        tableBodyId: 'rkhTableBody',
        headerRow: ['DIVISI_ID', 'NIK', 'REPORT_DATE', 'SEND_DATE', 'SEND_TIME', 'NILAI', 'NAMA'],
        expectedCols: 7
    });

    if (wb.SheetNames.length === 0) {
        alert('Tidak ada data untuk di-export');
        return;
    }

    const fileName = `Laporan_Harian_Semua_Data_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
}

function appendSheetFromTable_({ wb, sheetName, tableBodyId, headerRow, expectedCols }) {
    const tb = document.getElementById(tableBodyId);
    const rows = tb.querySelectorAll('tr');
    if (rows.length <= 1 || rows[0].querySelector('td[colspan]')) return;

    const data = [headerRow];
    rows.forEach(r => {
        const cols = r.querySelectorAll('td');
        if (cols.length === expectedCols) {
            data.push(Array.from(cols).map(c => c.textContent));
        }
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
}
