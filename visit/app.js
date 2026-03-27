const DB_NAME = 'visit-magang-db';
const DB_VERSION = 1;
const STORE_PARTICIPANTS = 'participants';
const STORE_REPORTS = 'reports';
const STORE_SETTINGS = 'settings';

let db;
let participantsCache = [];
let reportsCache = [];
let selectedParticipants = [];
let editingReportId = null;

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

document.addEventListener('DOMContentLoaded', async () => {
  bindNavigation();
  bindTheme();
  await initDb();
  await seedDefaultParticipantsIfEmpty();
  await refreshAll();
  bindEvents();
  $('#visitDate').valueAsDate = new Date();
});

function bindNavigation() {
  $$('.nav-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      $$('.nav-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const viewId = btn.dataset.view;
      $$('.view').forEach(v => v.classList.remove('active'));
      $('#' + viewId).classList.add('active');
      const titleMap = {
        dashboardView: ['Dashboard', 'Ringkasan hasil visit dan data peserta magang.'],
        visitFormView: ['Form Visit', 'Catat kunjungan, observasi, output, dan tindak lanjut.'],
        reportsView: ['Laporan Visit', 'Kelola, edit, cari, dan kirim ringkasan ke WhatsApp.'],
        masterView: ['Master Peserta', 'Data peserta tersimpan offline di IndexedDB.'],
        guideView: ['Panduan Visit', 'Panduan coaching lapangan yang bisa dibuka setiap saat.']
      };
      $('#pageTitle').textContent = titleMap[viewId][0];
      $('#pageSubtitle').textContent = titleMap[viewId][1];
    });
  });
}

function bindTheme() {
  const saved = localStorage.getItem('visit-magang-theme') || 'light';
  document.body.classList.toggle('dark', saved === 'dark');
  $('#themeToggle').checked = saved === 'dark';
  $('#themeToggle').addEventListener('change', () => {
    const mode = $('#themeToggle').checked ? 'dark' : 'light';
    document.body.classList.toggle('dark', mode === 'dark');
    localStorage.setItem('visit-magang-theme', mode);
  });
}

function bindEvents() {
  $('#participantSearch').addEventListener('input', handleParticipantSearch);
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.field')) {
      $('#participantSuggestions').classList.add('hidden');
    }
  });

  $('#visitForm').addEventListener('submit', saveVisitReport);
  $('#clearFormBtn').addEventListener('click', resetForm);
  $('#previewWhatsappBtn').addEventListener('click', updateWhatsappPreview);
  $('#copyWhatsappBtn').addEventListener('click', async () => {
    updateWhatsappPreview();
    const text = $('#whatsappPreview').value;
    if (!text.trim()) return alert('Preview WhatsApp masih kosong.');
    await navigator.clipboard.writeText(text);
    alert('Teks WhatsApp berhasil disalin.');
  });

  $('#reportSearch').addEventListener('input', renderReports);
  $('#reportUnitFilter').addEventListener('change', renderReports);
  $('#exportXlsxBtn').addEventListener('click', exportReportsToXlsx);

  $('#loadDefaultBtn').addEventListener('click', loadDefaultParticipants);
  $('#masterUpload').addEventListener('change', importMasterFile);

  $('#resetAppBtn').addEventListener('click', resetAppData);
  $('#openGuideBtn').addEventListener('click', () => openGuideView());
  const guideFormBtn = $('#openGuideFromFormBtn');
  if (guideFormBtn) guideFormBtn.addEventListener('click', () => openGuideView());
}


function openGuideView() {
  const btn = document.querySelector('[data-view="guideView"]');
  if (btn) btn.click();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function initDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = (event) => {
      db = event.target.result;
      if (!db.objectStoreNames.contains(STORE_PARTICIPANTS)) {
        const store = db.createObjectStore(STORE_PARTICIPANTS, { keyPath: 'nik' });
        store.createIndex('nama', 'nama', { unique: false });
        store.createIndex('unit', 'unit', { unique: false });
      }
      if (!db.objectStoreNames.contains(STORE_REPORTS)) {
        const store = db.createObjectStore(STORE_REPORTS, { keyPath: 'id' });
        store.createIndex('visitDate', 'visitDate', { unique: false });
        store.createIndex('location', 'location', { unique: false });
      }
      if (!db.objectStoreNames.contains(STORE_SETTINGS)) {
        db.createObjectStore(STORE_SETTINGS, { keyPath: 'key' });
      }
    };
    request.onsuccess = (event) => {
      db = event.target.result;
      resolve();
    };
    request.onerror = () => reject(request.error);
  });
}

function tx(storeName, mode = 'readonly') {
  return db.transaction(storeName, mode).objectStore(storeName);
}

function getAll(storeName) {
  return new Promise((resolve, reject) => {
    const req = tx(storeName).getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

function put(storeName, data) {
  return new Promise((resolve, reject) => {
    const req = tx(storeName, 'readwrite').put(data);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

function bulkPut(storeName, rows) {
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(storeName, 'readwrite');
    const store = transaction.objectStore(storeName);
    rows.forEach(row => store.put(row));
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
}

function clearStore(storeName) {
  return new Promise((resolve, reject) => {
    const req = tx(storeName, 'readwrite').clear();
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

async function seedDefaultParticipantsIfEmpty() {
  const current = await getAll(STORE_PARTICIPANTS);
  if (current.length) return;
  try {
    const res = await fetch('data/default-participants.json');
    const rows = await res.json();
    await bulkPut(STORE_PARTICIPANTS, sanitizeParticipants(rows));
  } catch (error) {
    console.warn('Gagal memuat data bawaan', error);
  }
}

async function refreshAll() {
  participantsCache = await getAll(STORE_PARTICIPANTS);
  reportsCache = await getAll(STORE_REPORTS);
  participantsCache.sort((a, b) => String(a.nama).localeCompare(String(b.nama), 'id'));
  reportsCache.sort((a, b) => String(b.visitDate).localeCompare(String(a.visitDate), 'id'));
  renderMasterTable();
  renderDashboard();
  renderReports();
  renderUnitFilter();
}

function sanitizeParticipants(rows) {
  return rows.map((row) => ({
    nik: String(row.nik ?? '').trim(),
    nama: String(row.nama ?? '').trim(),
    jenis_pelatihan: String(row.jenis_pelatihan ?? '').trim(),
    tahun: String(row.tahun ?? '').trim(),
    lokasi_ojt: String(row.lokasi_ojt ?? '').trim(),
    unit: String(row.unit ?? '').trim(),
    region: String(row.region ?? '').trim(),
    group: String(row.group ?? '').trim()
  })).filter(row => row.nik && row.nama);
}

function renderDashboard() {
  $('#statParticipants').textContent = participantsCache.length;
  $('#statReports').textContent = reportsCache.length;
  $('#statUnits').textContent = new Set(participantsCache.map(p => p.unit).filter(Boolean)).size;
  $('#statLastVisit').textContent = reportsCache[0]?.visitDate ? formatDateID(reportsCache[0].visitDate) : '-';

  const recent = reportsCache.slice(0, 5);
  const wrap = $('#recentReports');
  if (!recent.length) {
    wrap.className = 'stack-list empty-state';
    wrap.textContent = 'Belum ada laporan visit.';
    return;
  }
  wrap.className = 'stack-list';
  wrap.innerHTML = recent.map(r => `
    <div class="report-card">
      <h4>${escapeHtml(r.location || '-')}</h4>
      <div class="report-meta">${formatDateID(r.visitDate)} • ${escapeHtml((r.mentees || []).map(x => x.nama).join(', '))}</div>
      <div>${escapeHtml(shorten(r.resultsObtained || r.summary || '-', 180))}</div>
    </div>
  `).join('');
}

function renderMasterTable() {
  const body = $('#masterTableBody');
  if (!participantsCache.length) {
    body.innerHTML = '<tr><td colspan="8" class="empty-state">Belum ada data master peserta.</td></tr>';
    return;
  }
  body.innerHTML = participantsCache.map(p => `
    <tr>
      <td>${escapeHtml(p.nik)}</td>
      <td>${escapeHtml(p.nama)}</td>
      <td>${escapeHtml(p.jenis_pelatihan)}</td>
      <td>${escapeHtml(p.tahun)}</td>
      <td>${escapeHtml(p.lokasi_ojt)}</td>
      <td>${escapeHtml(p.unit)}</td>
      <td>${escapeHtml(p.region)}</td>
      <td>${escapeHtml(p.group)}</td>
    </tr>
  `).join('');
}

function renderUnitFilter() {
  const units = Array.from(new Set(participantsCache.map(p => p.unit).filter(Boolean))).sort();
  const select = $('#reportUnitFilter');
  const current = select.value;
  select.innerHTML = '<option value="">Semua unit</option>' + units.map(u => `<option value="${escapeHtmlAttr(u)}">${escapeHtml(u)}</option>`).join('');
  select.value = units.includes(current) ? current : '';
}

function handleParticipantSearch(e) {
  const q = e.target.value.trim().toLowerCase();
  const box = $('#participantSuggestions');
  if (!q) {
    box.classList.add('hidden');
    box.innerHTML = '';
    return;
  }
  const selectedNiks = new Set(selectedParticipants.map(p => p.nik));
  const results = participantsCache.filter(p => {
    const hay = `${p.nik} ${p.nama} ${p.unit} ${p.lokasi_ojt} ${p.region}`.toLowerCase();
    return hay.includes(q) && !selectedNiks.has(p.nik);
  }).slice(0, 8);

  if (!results.length) {
    box.innerHTML = '<div class="suggestion-item">Data tidak ditemukan.</div>';
    box.classList.remove('hidden');
    return;
  }

  box.innerHTML = results.map(p => `
    <button type="button" class="suggestion-item" data-nik="${escapeHtmlAttr(p.nik)}">
      <strong>${escapeHtml(p.nama)}</strong><br>
      <small>${escapeHtml([p.nik, p.unit, p.lokasi_ojt, p.region].filter(Boolean).join(' • '))}</small>
    </button>
  `).join('');
  box.classList.remove('hidden');

  $$('.suggestion-item[data-nik]').forEach(btn => {
    btn.addEventListener('click', () => {
      const participant = participantsCache.find(p => p.nik === btn.dataset.nik);
      if (!participant) return;
      selectedParticipants.push(participant);
      renderSelectedParticipants();
      $('#participantSearch').value = '';
      box.classList.add('hidden');
      box.innerHTML = '';
      updateWhatsappPreview();
    });
  });
}

function renderSelectedParticipants() {
  const wrap = $('#selectedParticipants');
  if (!selectedParticipants.length) {
    wrap.className = 'selected-chips empty-box';
    wrap.textContent = 'Belum ada peserta dipilih.';
    return;
  }
  wrap.className = 'selected-chips';
  wrap.innerHTML = selectedParticipants.map(p => `
    <div class="chip">
      <button type="button" data-remove-nik="${escapeHtmlAttr(p.nik)}">×</button>
      <strong>${escapeHtml(p.nama)}</strong>
      <small>${escapeHtml([p.nik, p.unit, p.lokasi_ojt, p.region].filter(Boolean).join(' • '))}</small>
    </div>
  `).join('');

  $$('[data-remove-nik]').forEach(btn => {
    btn.addEventListener('click', () => {
      selectedParticipants = selectedParticipants.filter(p => p.nik !== btn.dataset.removeNik);
      renderSelectedParticipants();
      updateWhatsappPreview();
    });
  });
}

async function saveVisitReport(e) {
  e.preventDefault();
  if (!selectedParticipants.length) {
    return alert('Pilih minimal satu peserta terlebih dahulu.');
  }
  const payload = collectFormPayload();
  if (!payload.visitDate || !payload.location || !payload.summary || !payload.resultsObtained) {
    return alert('Mohon lengkapi field wajib.');
  }
  await put(STORE_REPORTS, payload);
  editingReportId = payload.id;
  await refreshAll();
  updateWhatsappPreview();
  $('#editingBadge').classList.remove('hidden');
  alert('Laporan visit berhasil disimpan.');
}

function collectFormPayload() {
  const nowIso = new Date().toISOString();
  return {
    id: editingReportId || `visit-${Date.now()}`,
    visitDate: $('#visitDate').value,
    location: $('#location').value.trim(),
    mentees: selectedParticipants.map(p => ({
      nik: p.nik, nama: p.nama, unit: p.unit, lokasi_ojt: p.lokasi_ojt, region: p.region
    })),
    summary: $('#summary').value.trim(),
    activityReview: $('#activityReview').value.trim(),
    technicalUnderstanding: $('#technicalUnderstanding').value.trim(),
    fieldObservation: $('#fieldObservation').value.trim(),
    microTeaching: $('#microTeaching').value.trim(),
    livingCondition: $('#livingCondition').value.trim(),
    motivationFuture: $('#motivationFuture').value.trim(),
    resultsObtained: $('#resultsObtained').value.trim(),
    followUp: $('#followUp').value.trim(),
    specialNotes: $('#specialNotes').value.trim(),
    createdAt: editingReportId ? (reportsCache.find(r => r.id === editingReportId)?.createdAt || nowIso) : nowIso,
    updatedAt: nowIso
  };
}

function resetForm() {
  editingReportId = null;
  selectedParticipants = [];
  $('#visitForm').reset();
  $('#visitDate').valueAsDate = new Date();
  $('#whatsappPreview').value = '';
  renderSelectedParticipants();
  $('#editingBadge').classList.add('hidden');
  $('#participantSuggestions').classList.add('hidden');
  $('#participantSuggestions').innerHTML = '';
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function loadReportToForm(id) {
  const report = reportsCache.find(r => r.id === id);
  if (!report) return;
  editingReportId = report.id;
  $('#visitDate').value = report.visitDate || '';
  $('#location').value = report.location || '';
  $('#summary').value = report.summary || '';
  $('#activityReview').value = report.activityReview || '';
  $('#technicalUnderstanding').value = report.technicalUnderstanding || '';
  $('#fieldObservation').value = report.fieldObservation || '';
  $('#microTeaching').value = report.microTeaching || '';
  $('#livingCondition').value = report.livingCondition || '';
  $('#motivationFuture').value = report.motivationFuture || '';
  $('#resultsObtained').value = report.resultsObtained || '';
  $('#followUp').value = report.followUp || '';
  $('#specialNotes').value = report.specialNotes || '';
  selectedParticipants = (report.mentees || []).map(m => participantsCache.find(p => p.nik === m.nik) || m);
  renderSelectedParticipants();
  updateWhatsappPreview();
  $('#editingBadge').classList.remove('hidden');
  document.querySelector('[data-view="visitFormView"]').click();
}

function renderReports() {
  const wrap = $('#reportsContainer');
  const q = $('#reportSearch').value.trim().toLowerCase();
  const unit = $('#reportUnitFilter').value;
  const filtered = reportsCache.filter(r => {
    const menteeNames = (r.mentees || []).map(m => `${m.nama} ${m.unit} ${m.nik}`).join(' ');
    const hay = `${r.visitDate} ${r.location} ${r.summary} ${r.resultsObtained} ${menteeNames}`.toLowerCase();
    const unitMatch = !unit || (r.mentees || []).some(m => m.unit === unit);
    return (!q || hay.includes(q)) && unitMatch;
  });

  if (!filtered.length) {
    wrap.className = 'report-list empty-state';
    wrap.textContent = 'Tidak ada laporan yang cocok.';
    return;
  }
  wrap.className = 'report-list';
  wrap.innerHTML = filtered.map(r => `
    <div class="report-card">
      <div class="panel-head between wrap">
        <div>
          <h4>${escapeHtml(r.location || '-')}</h4>
          <div class="report-meta">${formatDateID(r.visitDate)} • ${escapeHtml((r.mentees || []).map(m => m.nama).join(', '))}</div>
        </div>
        <div class="tag">${escapeHtml((r.mentees || []).map(m => m.unit).filter(Boolean).join(', ') || '-')}</div>
      </div>
      <div><strong>Ringkasan:</strong> ${escapeHtml(r.summary || '-')}</div>
      <div style="margin-top:8px;"><strong>Hasil:</strong> ${escapeHtml(r.resultsObtained || '-')}</div>
      <div class="report-actions">
        <button class="secondary" data-edit-id="${escapeHtmlAttr(r.id)}">Edit</button>
        <button class="ghost" data-wa-id="${escapeHtmlAttr(r.id)}">WhatsApp</button>
      </div>
    </div>
  `).join('');

  $$('[data-edit-id]').forEach(btn => btn.addEventListener('click', () => loadReportToForm(btn.dataset.editId)));
  $$('[data-wa-id]').forEach(btn => btn.addEventListener('click', () => shareToWhatsapp(btn.dataset.waId)));
}

function buildWhatsappText(report) {
  const mentees = (report.mentees || []).map((m, i) => `${i + 1}. ${m.nama}${m.unit ? ' (' + m.unit + ')' : ''}`).join('\n') || '-';
  return [
    '*LAPORAN VISIT PESERTA MAGANG*',
    `Tanggal: ${formatDateID(report.visitDate)}`,
    `Lokasi: ${report.location || '-'}`,
    '',
    '*Mentee:*',
    mentees,
    '',
    '*Ringkasan Kunjungan:*',
    report.summary || '-',
    '',
    '*Hasil yang Diperoleh:*',
    report.resultsObtained || '-',
    report.followUp ? '\n*Tindak Lanjut:*\n' + report.followUp : ''
  ].join('\n');
}

function updateWhatsappPreview() {
  const payload = collectFormPayload();
  $('#whatsappPreview').value = buildWhatsappText(payload);
}

function shareToWhatsapp(reportId) {
  const report = reportsCache.find(r => r.id === reportId);
  if (!report) return;
  const text = buildWhatsappText(report);
  const url = 'https://wa.me/?text=' + encodeURIComponent(text);
  window.open(url, '_blank');
}

async function loadDefaultParticipants() {
  try {
    const res = await fetch('data/default-participants.json');
    const rows = await res.json();
    await clearStore(STORE_PARTICIPANTS);
    await bulkPut(STORE_PARTICIPANTS, sanitizeParticipants(rows));
    await refreshAll();
    alert('Data peserta bawaan berhasil dimuat.');
  } catch (error) {
    console.error(error);
    alert('Gagal memuat data bawaan.');
  }
}

async function importMasterFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  try {
    const rows = await parseXlsxFile(file);
    const normalized = sanitizeParticipants(rows);
    if (!normalized.length) throw new Error('Tidak ada data valid.');
    await clearStore(STORE_PARTICIPANTS);
    await bulkPut(STORE_PARTICIPANTS, normalized);
    await refreshAll();
    alert(`Master peserta berhasil diimpor: ${normalized.length} data.`);
  } catch (error) {
    console.error(error);
    alert('Gagal membaca file .xlsx. Pastikan format kolom sesuai master peserta.');
  } finally {
    e.target.value = '';
  }
}

async function resetAppData() {
  const ok = confirm('Semua master peserta dan laporan visit di aplikasi ini akan dihapus. Lanjutkan?');
  if (!ok) return;
  await clearStore(STORE_PARTICIPANTS);
  await clearStore(STORE_REPORTS);
  await refreshAll();
  resetForm();
  alert('Semua data aplikasi berhasil dihapus.');
}

function formatDateID(value) {
  if (!value) return '-';
  const date = new Date(value + 'T00:00:00');
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }).format(date);
}

function shorten(text, max) {
  return text && text.length > max ? text.slice(0, max - 1) + '…' : (text || '');
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function escapeHtmlAttr(value) {
  return escapeHtml(value).replaceAll('`', '&#96;');
}

// =========================
// XLSX IMPORT (browser-only)
// =========================
async function parseXlsxFile(file) {
  const buffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const workbookXml = await zip.file('xl/workbook.xml').async('string');
  const workbookDoc = new DOMParser().parseFromString(workbookXml, 'application/xml');
  const firstSheet = workbookDoc.querySelector('sheet');
  const sheetName = firstSheet?.getAttribute('name');
  const relId = firstSheet?.getAttribute('r:id');
  if (!relId) throw new Error('Worksheet tidak ditemukan.');

  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const relsDoc = new DOMParser().parseFromString(relsXml, 'application/xml');
  const relNode = Array.from(relsDoc.getElementsByTagName('Relationship')).find(n => n.getAttribute('Id') === relId);
  const target = relNode?.getAttribute('Target');
  if (!target) throw new Error('Relasi worksheet tidak ditemukan.');

  const worksheetPath = target.startsWith('/') ? target.slice(1) : ('xl/' + target.replace(/^\/?/, ''));
  const sharedStrings = await readSharedStrings(zip);
  const sheetXml = await zip.file(worksheetPath).async('string');
  const sheetDoc = new DOMParser().parseFromString(sheetXml, 'application/xml');

  const rows = Array.from(sheetDoc.getElementsByTagName('row')).map(row => {
    const cells = Array.from(row.getElementsByTagName('c')).map(cell => readCell(cell, sharedStrings));
    return cells;
  }).filter(row => row.length);

  if (!rows.length) return [];
  const headers = rows[0].map(h => String(h || '').trim());
  return rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((header, idx) => {
      obj[header] = r[idx] ?? '';
    });
    return obj;
  });
}

async function readSharedStrings(zip) {
  const entry = zip.file('xl/sharedStrings.xml');
  if (!entry) return [];
  const xml = await entry.async('string');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return Array.from(doc.getElementsByTagName('si')).map(si =>
    Array.from(si.getElementsByTagName('t')).map(t => t.textContent || '').join('')
  );
}

function readCell(cell, sharedStrings) {
  const type = cell.getAttribute('t');
  if (type === 'inlineStr') {
    return Array.from(cell.getElementsByTagName('t')).map(t => t.textContent || '').join('');
  }
  const v = cell.getElementsByTagName('v')[0]?.textContent ?? '';
  if (type === 's') return sharedStrings[Number(v)] ?? '';
  return v;
}

// =========================
// XLSX EXPORT (minimal OOXML)
// =========================
async function exportReportsToXlsx() {
  const rows = reportsCache.map((r, idx) => ({
    No: idx + 1,
    Tanggal: r.visitDate,
    Lokasi: r.location,
    Mentee: (r.mentees || []).map(m => m.nama).join('; '),
    Unit: Array.from(new Set((r.mentees || []).map(m => m.unit).filter(Boolean))).join('; '),
    Ringkasan_Kunjungan: r.summary,
    Aktivitas_Peserta: r.activityReview,
    Pemahaman_Teknis: r.technicalUnderstanding,
    Observasi_Lapangan: r.fieldObservation,
    Micro_Teaching: r.microTeaching,
    Adaptasi_Tempat_Tinggal_Makan_UangSaku: r.livingCondition,
    Motivasi_Masa_Depan: r.motivationFuture,
    Hasil_Diperoleh: r.resultsObtained,
    Tindak_Lanjut: r.followUp,
    Catatan_Khusus: r.specialNotes,
    Dibuat_Pada: r.createdAt,
    Diupdate_Pada: r.updatedAt
  }));

  if (!rows.length) {
    return alert('Belum ada laporan untuk diexport.');
  }

  const participantRows = participantsCache.map((p, idx) => ({
    No: idx + 1,
    NIK: p.nik,
    Nama: p.nama,
    Jenis_Pelatihan: p.jenis_pelatihan,
    Tahun: p.tahun,
    Lokasi_OJT: p.lokasi_ojt,
    Unit: p.unit,
    Region: p.region,
    Group: p.group
  }));

  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypesXml());
  zip.folder('_rels').file('.rels', rootRelsXml());
  zip.folder('docProps').file('app.xml', appXml(['Laporan Visit', 'Master Peserta']));
  zip.folder('docProps').file('core.xml', coreXml());

  const xl = zip.folder('xl');
  xl.file('workbook.xml', workbookXml());
  xl.folder('_rels').file('workbook.xml.rels', workbookRelsXml());
  xl.folder('worksheets').file('sheet1.xml', worksheetXml(rows));
  xl.folder('worksheets').file('sheet2.xml', worksheetXml(participantRows));
  xl.file('styles.xml', stylesXml());

  const blob = await zip.generateAsync({ type: 'blob' });
  downloadBlob(blob, `laporan-visit-magang-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function worksheetXml(rows) {
  const headers = Object.keys(rows[0] || {});
  const matrix = [headers, ...rows.map(row => headers.map(h => row[h] ?? ''))];
  const maxWidths = headers.map((header, colIndex) => {
    return Math.min(45, Math.max(
      String(header).length,
      ...rows.map(r => String(r[header] ?? '').length)
    ) + 2);
  });

  const colsXml = maxWidths.map((w, i) => `<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join('');
  const rowsXml = matrix.map((row, rowIndex) => {
    const cells = row.map((value, colIndex) => {
      const ref = colName(colIndex + 1) + (rowIndex + 1);
      const styleId = rowIndex === 0 ? 1 : 0;
      return `<c r="${ref}" t="inlineStr" s="${styleId}"><is><t xml:space="preserve">${xmlEscape(String(value ?? ''))}</t></is></c>`;
    }).join('');
    return `<row r="${rowIndex+1}">${cells}</row>`;
  }).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetViews><sheetView workbookViewId="0"/></sheetViews>
    <cols>${colsXml}</cols>
    <sheetData>${rowsXml}</sheetData>
  </worksheet>`;
}

function contentTypesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  </Types>`;
}

function rootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  </Relationships>`;
}

function workbookXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
      <sheet name="Laporan Visit" sheetId="1" r:id="rId1"/>
      <sheet name="Master Peserta" sheetId="2" r:id="rId2"/>
    </sheets>
  </workbook>`;
}

function workbookRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  </Relationships>`;
}

function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="2">
      <font><sz val="11"/><name val="Calibri"/></font>
      <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    </fonts>
    <fills count="3">
      <fill><patternFill patternType="none"/></fill>
      <fill><patternFill patternType="gray125"/></fill>
      <fill><patternFill patternType="solid"><fgColor rgb="FF2457D6"/><bgColor indexed="64"/></patternFill></fill>
    </fills>
    <borders count="1">
      <border><left/><right/><top/><bottom/><diagonal/></border>
    </borders>
    <cellStyleXfs count="1">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="2">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
      <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    </cellXfs>
    <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  </styleSheet>`;
}

function appXml(sheetNames) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>Visit Magang Offline App</Application>
    <TitlesOfParts>
      <vt:vector size="${sheetNames.length}" baseType="lpstr">
        ${sheetNames.map(n => `<vt:lpstr>${xmlEscape(n)}</vt:lpstr>`).join('')}
      </vt:vector>
    </TitlesOfParts>
  </Properties>`;
}

function coreXml() {
  const iso = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:title>Laporan Visit Magang</dc:title>
    <dc:creator>ChatGPT</dc:creator>
    <cp:lastModifiedBy>ChatGPT</cp:lastModifiedBy>
    <dcterms:created xsi:type="dcterms:W3CDTF">${iso}</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified>
  </cp:coreProperties>`;
}

function colName(n) {
  let s = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function xmlEscape(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
