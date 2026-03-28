const DB_NAME = 'visit-magang-db';
const DB_VERSION = 3;
const STORE_PARTICIPANTS = 'participants';
const STORE_REPORTS = 'reports';
const STORE_SETTINGS = 'settings';
const STORE_SYNC_QUEUE = 'sync_queue';
const STORE_SYNC_FAILED = 'sync_failed';
const STORE_SYNC_CONFLICTS = 'sync_conflicts';

const SYNC_SETTINGS_KEY = 'sync_config';
const GAS_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbxTOPuN4FJTedfU_hpEbxVcB1XH6cAUVUABr5wAc0WI0O5AnvXukCETb8pwdB19y8p-/exec';
const DEFAULT_SYNC_CONFIG = {
  deviceName: getDefaultDeviceName(),
  autoSyncOnSave: true,
  lastPushAt: '',
  lastPullAt: ''
};

let db;
let participantsCache = [];
let reportsCache = [];
let selectedParticipants = [];
let editingReportId = null;
let syncConfig = { ...DEFAULT_SYNC_CONFIG };
let postMessageResolvers = new Map();
let syncInProgress = false;
let syncProgressState = { total: 0, done: 0 };

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

document.addEventListener('DOMContentLoaded', async () => {
  window.addEventListener('message', handlePostMessageResponse);
  bindNavigation();
  bindTheme();
  await initDb();
  await loadSyncConfig();
  await seedDefaultParticipantsIfEmpty();
  await refreshAll();
  bindEvents();
  $('#visitDate').valueAsDate = new Date();
  renderSelectedParticipants();
  renderSyncSettings();
  await updateSyncStats();
  await showStartupSyncNotice();
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
        dashboardView: ['Dashboard', 'Ringkasan hasil visit, antrian sinkronisasi, dan data peserta magang.'],
        visitFormView: ['Form Visit', 'Catat kunjungan, observasi, output, tindak lanjut, dan siap sinkron ke database online.'],
        reportsView: ['Laporan Visit', 'Kelola, edit, cari, dan kirim ringkasan ke WhatsApp.'],
        masterView: ['Master Peserta', 'Data peserta tersimpan lokal dan bisa ditarik dari database online.'],
        conflictsView: ['Konflik Data', 'Pantau benturan data antar perangkat sebelum menentukan tindak lanjut.'],
        failedSyncView: ['Gagal Sinkron', 'Data yang gagal sinkron dapat dilihat dan dikirim ulang dari sini.'],
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

  $('#saveSyncConfigBtn')?.addEventListener('click', saveSyncConfigFromForm);
  $('#testConnectionBtn')?.addEventListener('click', testSyncConnection);
  $('#pushSyncBtn')?.addEventListener('click', pushPendingQueue);
  $('#pullSyncBtn')?.addEventListener('click', pullCloudData);
  $('#fullSyncBtn')?.addEventListener('click', fullSync);
  $('#retryAllFailedBtn')?.addEventListener('click', retryAllFailedSync);
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
      if (!db.objectStoreNames.contains(STORE_SYNC_QUEUE)) {
        const store = db.createObjectStore(STORE_SYNC_QUEUE, { keyPath: 'queueKey' });
        store.createIndex('entityType', 'entityType', { unique: false });
        store.createIndex('updatedAt', 'updatedAt', { unique: false });
      }
      if (!db.objectStoreNames.contains(STORE_SYNC_FAILED)) {
        const store = db.createObjectStore(STORE_SYNC_FAILED, { keyPath: 'failedKey' });
        store.createIndex('entityType', 'entityType', { unique: false });
        store.createIndex('failedAt', 'failedAt', { unique: false });
      }
      if (!db.objectStoreNames.contains(STORE_SYNC_CONFLICTS)) {
        const store = db.createObjectStore(STORE_SYNC_CONFLICTS, { keyPath: 'conflictKey' });
        store.createIndex('entityType', 'entityType', { unique: false });
        store.createIndex('detectedAt', 'detectedAt', { unique: false });
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

function getByKey(storeName, key) {
  return new Promise((resolve, reject) => {
    const req = tx(storeName).get(key);
    req.onsuccess = () => resolve(req.result || null);
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

function deleteByKey(storeName, key) {
  return new Promise((resolve, reject) => {
    const req = tx(storeName, 'readwrite').delete(key);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

function clearStore(storeName) {
  return new Promise((resolve, reject) => {
    const req = tx(storeName, 'readwrite').clear();
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

async function loadSyncConfig() {
  const saved = await getByKey(STORE_SETTINGS, SYNC_SETTINGS_KEY);
  syncConfig = { ...DEFAULT_SYNC_CONFIG, ...(saved?.value || {}) };
}

async function saveSyncConfig(config) {
  syncConfig = { ...syncConfig, ...config };
  await put(STORE_SETTINGS, { key: SYNC_SETTINGS_KEY, value: syncConfig });
}

function renderSyncSettings() {
  if ($('#syncEndpointLabel')) $('#syncEndpointLabel').textContent = GAS_WEBAPP_URL;
  if ($('#deviceName')) $('#deviceName').value = syncConfig.deviceName || getDefaultDeviceName();
  if ($('#autoSyncOnSave')) $('#autoSyncOnSave').checked = !!syncConfig.autoSyncOnSave;
  if ($('#syncLastPush')) $('#syncLastPush').textContent = syncConfig.lastPushAt ? formatDateTimeID(syncConfig.lastPushAt) : '-';
  if ($('#syncLastPull')) $('#syncLastPull').textContent = syncConfig.lastPullAt ? formatDateTimeID(syncConfig.lastPullAt) : '-';
}

async function seedDefaultParticipantsIfEmpty() {
  const current = await getAll(STORE_PARTICIPANTS);
  if (current.length) return;
  try {
    const res = await fetch('data/default-participants.json');
    const rows = await res.json();
    const normalized = sanitizeParticipants(rows, { markSynced: false });
    await bulkPut(STORE_PARTICIPANTS, normalized);
  } catch (error) {
    console.warn('Gagal memuat data bawaan', error);
  }
}

async function refreshAll() {
  participantsCache = await getAll(STORE_PARTICIPANTS);
  reportsCache = await getAll(STORE_REPORTS);
  participantsCache.sort((a, b) => String(a.nama).localeCompare(String(b.nama), 'id'));
  reportsCache.sort((a, b) => String(b.visitDate).localeCompare(String(a.visitDate), 'id') || String(b.updatedAt || '').localeCompare(String(a.updatedAt || ''), 'id'));
  renderMasterTable();
  renderDashboard();
  renderReports();
  renderUnitFilter();
  await renderFailedSyncs();
  await renderConflicts();
  await updateSyncStats();
}

function sanitizeParticipants(rows, options = {}) {
  const nowIso = new Date().toISOString();
  return rows.map((row) => {
    const nik = String(row.nik ?? row.NIK ?? '').trim();
    const updatedAt = row.updatedAt || row.updated_at || nowIso;
    const participant = {
      nik,
      nama: String(row.nama ?? row.Nama ?? '').trim(),
      jenis_pelatihan: String(row.jenis_pelatihan ?? row.Jenis_Pelatihan ?? row.jenisPelatihan ?? '').trim(),
      tahun: String(row.tahun ?? row.Tahun ?? '').trim(),
      lokasi_ojt: String(row.lokasi_ojt ?? row.Lokasi_OJT ?? row.lokasiOjt ?? '').trim(),
      unit: String(row.unit ?? row.Unit ?? '').trim(),
      region: String(row.region ?? row.Region ?? '').trim(),
      group: String(row.group ?? row.Group ?? '').trim(),
      createdAt: row.createdAt || row.created_at || updatedAt,
      updatedAt,
      syncStatus: options.markSynced ? 'synced' : (row.syncStatus || 'local')
    };
    return participant;
  }).filter(row => row.nik && row.nama);
}

function renderDashboard() {
  $('#statParticipants').textContent = participantsCache.length;
  $('#statReports').textContent = reportsCache.length;
  $('#statUnits').textContent = new Set(participantsCache.map(p => p.unit).filter(Boolean)).size;
  $('#statLastVisit').textContent = reportsCache[0]?.visitDate ? formatDateShortID(reportsCache[0].visitDate) : '-';

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
      <div class="report-meta">${formatDateSlashID(r.visitDate)} • ${escapeHtml((r.mentees || []).map(x => x.nama).join(', '))}</div>
      <div>${escapeHtml(shorten(r.resultsObtained || r.summary || '-', 180))}</div>
      <div class="report-meta" style="margin-top:8px;">Status sinkron: ${escapeHtml(humanSyncStatus(r.syncStatus))}</div>
    </div>
  `).join('');
}

function renderMasterTable() {
  const body = $('#masterTableBody');
  if (!participantsCache.length) {
    body.innerHTML = '<tr><td colspan="9" class="empty-state">Belum ada data master peserta.</td></tr>';
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
      <td>${escapeHtml(humanSyncStatus(p.syncStatus))}</td>
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
  payload.syncStatus = 'pending';
  await put(STORE_REPORTS, payload);
  await enqueueSync('report', payload.id, payload.updatedAt);
  editingReportId = payload.id;
  await refreshAll();
  updateWhatsappPreview();
  $('#editingBadge').classList.remove('hidden');

  let message = 'Laporan visit berhasil disimpan dan masuk antrian sinkronisasi.';
  if (syncConfig.autoSyncOnSave && hasGasUrl()) {
    try {
      const result = await pushPendingQueue({ silent: true, stopOnError: true });
      if (result.pushed > 0) message = 'Laporan visit berhasil disimpan dan langsung tersinkron ke database online.';
    } catch (error) {
      console.warn(error);
      message = 'Laporan visit berhasil disimpan. Sinkron otomatis belum berhasil, tetapi data sudah masuk antrian.';
    }
  }
  alert(message);
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
    updatedAt: nowIso,
    sourceDevice: syncConfig.deviceName || getDefaultDeviceName(),
    syncStatus: 'pending'
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
          <div class="report-meta">${formatDateSlashID(r.visitDate)} • ${escapeHtml((r.mentees || []).map(m => m.nama).join(', '))}</div>
        </div>
        <div class="tag">${escapeHtml((r.mentees || []).map(m => m.unit).filter(Boolean).join(', ') || '-')}</div>
      </div>
      <div><strong>Ringkasan:</strong> ${escapeHtml(r.summary || '-')}</div>
      <div style="margin-top:8px;"><strong>Hasil:</strong> ${escapeHtml(r.resultsObtained || '-')}</div>
      <div class="report-meta" style="margin-top:8px;">Sinkron: ${escapeHtml(humanSyncStatus(r.syncStatus))}</div>
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
    const normalized = sanitizeParticipants(rows, { markSynced: false }).map(row => ({ ...row, updatedAt: new Date().toISOString(), syncStatus: 'pending' }));
    await clearStore(STORE_PARTICIPANTS);
    await bulkPut(STORE_PARTICIPANTS, normalized);
    for (const row of normalized) {
      await enqueueSync('participant', row.nik, row.updatedAt);
    }
    await refreshAll();
    alert('Data peserta bawaan berhasil dimuat dan masuk antrian sinkronisasi.');
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
    const nowIso = new Date().toISOString();
    const normalized = sanitizeParticipants(rows, { markSynced: false }).map(row => ({ ...row, updatedAt: nowIso, createdAt: row.createdAt || nowIso, syncStatus: 'pending' }));
    if (!normalized.length) throw new Error('Tidak ada data valid.');
    await clearStore(STORE_PARTICIPANTS);
    await bulkPut(STORE_PARTICIPANTS, normalized);
    for (const row of normalized) {
      await enqueueSync('participant', row.nik, row.updatedAt);
    }
    await refreshAll();
    alert(`Master peserta berhasil diimpor: ${normalized.length} data. Semua data masuk antrian sinkronisasi.`);
  } catch (error) {
    console.error(error);
    alert('Gagal membaca file .xlsx. Pastikan format kolom sesuai master peserta.');
  } finally {
    e.target.value = '';
  }
}

async function resetAppData() {
  const ok = confirm('Semua master peserta, laporan visit, dan antrian sinkronisasi di aplikasi ini akan dihapus. Lanjutkan?');
  if (!ok) return;
  await clearStore(STORE_PARTICIPANTS);
  await clearStore(STORE_REPORTS);
  await clearStore(STORE_SYNC_QUEUE);
  await refreshAll();
  resetForm();
  alert('Semua data aplikasi berhasil dihapus.');
}

function parseDateFlexible(value) {
  if (!value) return null;
  const raw = String(value).trim();
  if (!raw) return null;

  const isoDateOnly = /^\d{4}-\d{2}-\d{2}$/;
  const slashDate = /^\d{2}\/\d{2}\/\d{4}$/;

  let date;
  if (isoDateOnly.test(raw)) {
    date = new Date(raw + 'T00:00:00');
  } else if (slashDate.test(raw)) {
    const [dd, mm, yyyy] = raw.split('/');
    date = new Date(`${yyyy}-${mm}-${dd}T00:00:00`);
  } else {
    date = new Date(raw);
  }

  return Number.isNaN(date.getTime()) ? null : date;
}

function formatDateID(value) {
  if (!value) return '-';
  const date = parseDateFlexible(value);
  if (!date) return value;
  return new Intl.DateTimeFormat('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }).format(date);
}

function formatDateShortID(value) {
  if (!value) return '-';
  const date = parseDateFlexible(value);
  if (!date) return value;
  return new Intl.DateTimeFormat('id-ID', { day: '2-digit', month: 'short', year: 'numeric' }).format(date);
}

function formatDateSlashID(value) {
  if (!value) return '-';
  const date = parseDateFlexible(value);
  if (!date) return value;
  return new Intl.DateTimeFormat('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(date);
}

function formatDateTimeID(value) {
  if (!value) return '-';
  const date = parseDateFlexible(value);
  if (!date) return value;
  return new Intl.DateTimeFormat('id-ID', {
    day: '2-digit', month: 'long', year: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit'
  }).format(date);
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

function humanSyncStatus(status) {
  const map = {
    pending: 'Menunggu sinkron',
    synced: 'Sudah sinkron',
    pulled: 'Ditarik dari online',
    error: 'Gagal sinkron',
    conflict: 'Konflik antar perangkat',
    local: 'Lokal'
  };
  return map[status] || '-';
}

function getDefaultDeviceName() {
  return `Device-${Math.random().toString(36).slice(2, 7).toUpperCase()}`;
}

function hasGasUrl() {
  return /^https:\/\//i.test(GAS_WEBAPP_URL || '');
}

async function saveSyncConfigFromForm() {
  const deviceName = $('#deviceName').value.trim() || getDefaultDeviceName();
  const autoSyncOnSave = !!$('#autoSyncOnSave').checked;
  await saveSyncConfig({ deviceName, autoSyncOnSave });
  renderSyncSettings();
  await updateSyncStats();
  alert('Pengaturan sinkronisasi berhasil disimpan.');
}

async function updateSyncStats() {
  const queue = await getAll(STORE_SYNC_QUEUE);
  const pendingReports = queue.filter(q => q.entityType === 'report').length;
  const pendingParticipants = queue.filter(q => q.entityType === 'participant').length;
  if ($('#pendingReportsCount')) $('#pendingReportsCount').textContent = String(pendingReports);
  if ($('#pendingParticipantsCount')) $('#pendingParticipantsCount').textContent = String(pendingParticipants);
  const failed = await getAll(STORE_SYNC_FAILED);
  const conflicts = await getAll(STORE_SYNC_CONFLICTS);
  if ($('#syncTotalQueue')) $('#syncTotalQueue').textContent = String(queue.length);
  if ($('#failedSyncCount')) $('#failedSyncCount').textContent = String(failed.length);
  if ($('#conflictCount')) $('#conflictCount').textContent = String(conflicts.length);
  renderSyncSettings();
}

async function enqueueSync(entityType, entityId, updatedAt) {
  await put(STORE_SYNC_QUEUE, {
    queueKey: `${entityType}:${entityId}`,
    entityType,
    entityId,
    updatedAt,
    createdAt: new Date().toISOString(),
    action: 'upsert'
  });
}

async function showStartupSyncNotice() {
  const queue = await getAll(STORE_SYNC_QUEUE);
  const failed = await getAll(STORE_SYNC_FAILED);
  const note = $('#startupSyncNotice');
  if (!note) return;
  if (!queue.length && !failed.length) {
    note.classList.add('hidden');
    return;
  }
  note.classList.remove('hidden');
  note.innerHTML = `Saat aplikasi dibuka terdapat <strong>${queue.length}</strong> antrian sinkron dan <strong>${failed.length}</strong> data gagal sinkron. Jalankan sinkron agar database lokal dan online kembali selaras.`;
}

async function testSyncConnection() {
  try {
    ensureSyncReady();
    setSyncStatus('Menguji koneksi ke Apps Script dan Google Sheet...', 'info');
    const res = await appsScriptRequest({
      action: 'init',
      t: Date.now()
    });
    if (!res.ok) throw new Error(res.message || 'Koneksi gagal.');
    setSyncStatus(`Koneksi berhasil. Sheet online siap dipakai. Reports: ${res.stats?.reports ?? 0}, Participants: ${res.stats?.participants ?? 0}`, 'success');
  } catch (error) {
    console.error(error);
    setSyncStatus(error.message || 'Gagal menguji koneksi.', 'error');
    alert(error.message || 'Gagal menguji koneksi.');
  }
}

function setSyncStatus(message, tone = 'info') {
  const box = $('#syncStatusBox');
  if (!box) return;
  box.className = `sync-status ${tone}`;
  box.textContent = message;
}

function ensureSyncReady() {
  if (!hasGasUrl()) throw new Error('Endpoint Apps Script belum valid di paket aplikasi.');
}

function startSyncProgress(title, total) {
  syncProgressState = { total: Math.max(0, Number(total) || 0), done: 0 };
  const card = $('#syncProgressCard');
  if (card) card.classList.remove('hidden');
  if ($('#syncProgressTitle')) $('#syncProgressTitle').textContent = title || 'Sinkronisasi berjalan';
  updateSyncProgress(0, total, 'Menyiapkan proses...');
}

function updateSyncProgress(done, total, stepText) {
  syncProgressState = { total: Math.max(0, Number(total) || 0), done: Math.max(0, Number(done) || 0) };
  const safeTotal = syncProgressState.total || 1;
  const pct = Math.max(0, Math.min(100, Math.round((syncProgressState.done / safeTotal) * 100)));
  if ($('#syncProgressFill')) $('#syncProgressFill').style.width = pct + '%';
  if ($('#syncProgressPercent')) $('#syncProgressPercent').textContent = pct + '%';
  if ($('#syncProgressStep')) $('#syncProgressStep').textContent = stepText || 'Memproses...';
  if ($('#syncProgressCount')) $('#syncProgressCount').textContent = `${syncProgressState.done}/${syncProgressState.total}`;
}

function finishSyncProgress(message) {
  updateSyncProgress(syncProgressState.total, syncProgressState.total, message || 'Selesai');
  setTimeout(() => { $('#syncProgressCard')?.classList.add('hidden'); }, 900);
}

async function markFailedSync(item, reason, record) {
  await put(STORE_SYNC_FAILED, {
    failedKey: item.queueKey,
    queueKey: item.queueKey,
    entityType: item.entityType,
    entityId: item.entityId,
    updatedAt: item.updatedAt,
    failedAt: new Date().toISOString(),
    reason: String(reason || 'Gagal sinkron'),
    snapshot: record || null
  });
}

async function clearFailedSync(queueKey) {
  await deleteByKey(STORE_SYNC_FAILED, queueKey);
}

async function recordConflict(entityType, entityId, localRow, remoteRow, reason) {
  const conflictKey = `${entityType}:${entityId}`;
  await put(STORE_SYNC_CONFLICTS, {
    conflictKey,
    entityType,
    entityId,
    detectedAt: new Date().toISOString(),
    reason: String(reason || 'Konflik data terdeteksi'),
    localUpdatedAt: localRow?.updatedAt || '',
    remoteUpdatedAt: remoteRow?.updatedAt || '',
    localDevice: localRow?.sourceDevice || '',
    remoteDevice: remoteRow?.sourceDevice || '',
    localSnapshot: localRow || null,
    remoteSnapshot: remoteRow || null
  });
}

async function clearConflict(entityType, entityId) {
  await deleteByKey(STORE_SYNC_CONFLICTS, `${entityType}:${entityId}`);
}

function normalizeEntityType(entityType) {
  return entityType === 'report' ? 'laporan' : 'master peserta';
}

function comparableRecord(row) {
  if (!row) return null;
  const clone = JSON.parse(JSON.stringify(row));
  delete clone.syncStatus;
  return clone;
}

function rowsAreDifferent(a, b) {
  return JSON.stringify(comparableRecord(a)) !== JSON.stringify(comparableRecord(b));
}

async function fetchRemoteEntity(entityType, entityId) {
  const res = await appsScriptRequest({
    action: 'getOne',
    entityType,
    entityId,
    t: Date.now()
  });
  if (!res.ok) throw new Error(res.message || 'Gagal membaca data online.');
  return res.data || { exists: false };
}

async function pushPendingQueue(options = {}) {
  if (syncInProgress) return { pushed: 0 };
  ensureSyncReady();
  syncInProgress = true;
  const silent = !!options.silent;
  const stopOnError = options.stopOnError !== false;
  try {
    let queue = await getAll(STORE_SYNC_QUEUE);
    queue = queue.sort((a, b) => String(a.updatedAt).localeCompare(String(b.updatedAt), 'id'));
    if (!queue.length) {
      if (!silent) setSyncStatus('Tidak ada antrian sinkronisasi. Semua data lokal sudah bersih.', 'success');
      finishSyncProgress('Tidak ada antrian.');
      return { pushed: 0 };
    }

    startSyncProgress('Mengirim antrian ke database online', queue.length);
    let pushed = 0;
    let processed = 0;
    for (const item of queue) {
      const storeName = item.entityType === 'report' ? STORE_REPORTS : STORE_PARTICIPANTS;
      const record = await getByKey(storeName, item.entityId);
      if (!record) {
        await deleteByKey(STORE_SYNC_QUEUE, item.queueKey);
        processed += 1;
        updateSyncProgress(processed, queue.length, `Melewati data lokal yang sudah tidak ada: ${item.entityId}`);
        continue;
      }

      updateSyncProgress(processed, queue.length, `Memeriksa ${normalizeEntityType(item.entityType)} ${item.entityId}`);
      const remote = await fetchRemoteEntity(item.entityType, item.entityId);
      if (remote.exists && remote.row && remote.row.updatedAt && isIncomingNewer(remote.row.updatedAt, record.updatedAt) && rowsAreDifferent(remote.row, record) && String(remote.row.sourceDevice || '') !== String(record.sourceDevice || '')) {
        record.syncStatus = 'conflict';
        await put(storeName, record);
        await recordConflict(item.entityType, item.entityId, record, remote.row, 'Versi online lebih baru dari perangkat lain.');
        await markFailedSync(item, 'Konflik: versi online lebih baru dari perangkat lain.', record);
        processed += 1;
        updateSyncProgress(processed, queue.length, `Konflik terdeteksi pada ${item.entityId}`);
        continue;
      }

      if (!silent) setSyncStatus(`Sinkron ${normalizeEntityType(item.entityType)}: ${item.entityId}`, 'info');
      try {
        const res = await postViaIframe({
          action: 'upsert',
          entityType: item.entityType,
          deviceName: syncConfig.deviceName,
          payload: JSON.stringify(record)
        });
        if (!res.ok) throw new Error(res.message || `Gagal sinkron ${item.entityId}`);
        record.syncStatus = 'synced';
        await put(storeName, record);
        await deleteByKey(STORE_SYNC_QUEUE, item.queueKey);
        await clearFailedSync(item.queueKey);
        await clearConflict(item.entityType, item.entityId);
        pushed += 1;
      } catch (error) {
        record.syncStatus = 'error';
        await put(storeName, record);
        await markFailedSync(item, error.message || `Gagal sinkron ${item.entityId}`, record);
        if (stopOnError && !silent) {
          throw error;
        }
      }
      processed += 1;
      updateSyncProgress(processed, queue.length, `Selesai memproses ${item.entityId}`);
    }

    const nowIso = new Date().toISOString();
    await saveSyncConfig({ lastPushAt: nowIso });
    await refreshAll();
    finishSyncProgress(`Sinkron kirim selesai. ${pushed} data berhasil.`);
    if (!silent) setSyncStatus(`Sinkron kirim selesai. ${pushed} data berhasil diunggah ke Google Sheet.`, 'success');
    return { pushed };
  } catch (error) {
    console.error(error);
    finishSyncProgress('Sinkron terhenti karena ada kendala.');
    setSyncStatus(error.message || 'Gagal sinkron kirim.', 'error');
    if (!silent && stopOnError) throw error;
    return { pushed: 0, error };
  } finally {
    syncInProgress = false;
  }
}

async function pullCloudData() {
  try {
    ensureSyncReady();
    setSyncStatus('Menarik data dari Google Sheet...', 'info');
    const res = await appsScriptRequest({
      action: 'pull',
      t: Date.now()
    });
    if (!res.ok) throw new Error(res.message || 'Gagal menarik data online.');

    const participants = sanitizeParticipants(res.data?.participants || [], { markSynced: true });
    const reports = (res.data?.reports || []).map(normalizePulledReport);
    const total = participants.length + reports.length;
    startSyncProgress('Menarik data terbaru dari online', total || 1);

    let done = 0;
    done += await mergePulledParticipants(participants, total, done);
    done += await mergePulledReports(reports, total, done);

    const nowIso = new Date().toISOString();
    await saveSyncConfig({ lastPullAt: nowIso });
    await refreshAll();
    finishSyncProgress('Tarik data selesai.');
    setSyncStatus(`Tarik data selesai. Reports: ${reports.length}, Participants: ${participants.length}. Data online sudah masuk ke penyimpanan lokal.`, 'success');
  } catch (error) {
    console.error(error);
    finishSyncProgress('Tarik data gagal.');
    setSyncStatus(error.message || 'Gagal menarik data online.', 'error');
    alert(error.message || 'Gagal menarik data online.');
  }
}

async function fullSync() {
  try {
    ensureSyncReady();
    setSyncStatus('Menjalankan sinkron penuh: kirim antrian lalu tarik data terbaru...', 'info');
    await pushPendingQueue({ silent: true, stopOnError: false });
    await pullCloudData();
    setSyncStatus('Sinkron penuh selesai.', 'success');
  } catch (error) {
    console.error(error);
    setSyncStatus(error.message || 'Sinkron penuh gagal.', 'error');
    alert(error.message || 'Sinkron penuh gagal.');
  }
}

async function mergePulledParticipants(incomingRows, total = incomingRows.length, offset = 0) {
  let processed = 0;
  for (const row of incomingRows) {
    const existing = await getByKey(STORE_PARTICIPANTS, row.nik);
    const pending = await getByKey(STORE_SYNC_QUEUE, `participant:${row.nik}`);
    if (existing && pending && rowsAreDifferent(existing, row) && String(existing.sourceDevice || '') !== String(row.sourceDevice || '') && isIncomingNewer(row.updatedAt, existing.updatedAt)) {
      await recordConflict('participant', row.nik, existing, row, 'Data online lebih baru saat pull dan masih ada perubahan lokal yang belum sinkron.');
      await markFailedSync({ queueKey: `participant:${row.nik}`, entityType: 'participant', entityId: row.nik, updatedAt: existing.updatedAt }, 'Konflik saat pull data participant.', existing);
      existing.syncStatus = 'conflict';
      await put(STORE_PARTICIPANTS, existing);
    } else if (!existing || isIncomingNewer(row.updatedAt, existing.updatedAt)) {
      await put(STORE_PARTICIPANTS, { ...existing, ...row, syncStatus: 'pulled' });
      await deleteByKey(STORE_SYNC_QUEUE, `participant:${row.nik}`);
      await clearFailedSync(`participant:${row.nik}`);
      await clearConflict('participant', row.nik);
    }
    processed += 1;
    updateSyncProgress(offset + processed, total, `Menarik master peserta ${row.nik}`);
  }
  return processed;
}

async function mergePulledReports(incomingRows, total = incomingRows.length, offset = 0) {
  let processed = 0;
  for (const row of incomingRows) {
    const existing = await getByKey(STORE_REPORTS, row.id);
    const queueKey = `report:${row.id}`;
    const pending = await getByKey(STORE_SYNC_QUEUE, queueKey);
    if (existing && pending && rowsAreDifferent(existing, row) && String(existing.sourceDevice || '') !== String(row.sourceDevice || '') && isIncomingNewer(row.updatedAt, existing.updatedAt)) {
      await recordConflict('report', row.id, existing, row, 'Data online lebih baru saat pull dan versi lokal masih pending.');
      await markFailedSync({ queueKey, entityType: 'report', entityId: row.id, updatedAt: existing.updatedAt }, 'Konflik saat pull data report.', existing);
      existing.syncStatus = 'conflict';
      await put(STORE_REPORTS, existing);
    } else if (!existing || isIncomingNewer(row.updatedAt, existing.updatedAt)) {
      await put(STORE_REPORTS, { ...existing, ...row, syncStatus: 'pulled' });
      await deleteByKey(STORE_SYNC_QUEUE, queueKey);
      await clearFailedSync(queueKey);
      await clearConflict('report', row.id);
    }
    processed += 1;
    updateSyncProgress(offset + processed, total, `Menarik laporan ${row.id}`);
  }
  return processed;
}

async function renderFailedSyncs() {
  const wrap = $('#failedSyncContainer');
  if (!wrap) return;
  const rows = (await getAll(STORE_SYNC_FAILED)).sort((a, b) => String(b.failedAt || '').localeCompare(String(a.failedAt || ''), 'id'));
  if (!rows.length) {
    wrap.className = 'stack-list empty-state';
    wrap.textContent = 'Belum ada data gagal sinkron.';
    return;
  }
  wrap.className = 'stack-list';
  wrap.innerHTML = rows.map(r => `
    <div class="report-card">
      <h4>${escapeHtml(normalizeEntityType(r.entityType))} • ${escapeHtml(r.entityId)}</h4>
      <div class="report-meta">Gagal: ${escapeHtml(formatDateTimeID(r.failedAt))}</div>
      <div>${escapeHtml(r.reason || '-')}</div>
      <div class="report-actions">
        <span class="status-chip error">Gagal sinkron</span>
        <button type="button" class="secondary retry-failed-btn" data-qk="${escapeHtmlAttr(r.queueKey)}">Sinkron ulang</button>
        <button type="button" class="ghost clear-failed-btn" data-qk="${escapeHtmlAttr(r.queueKey)}">Hapus dari daftar gagal</button>
      </div>
    </div>
  `).join('');
  $$('.retry-failed-btn').forEach(btn => btn.addEventListener('click', () => retryFailedSync(btn.dataset.qk)));
  $$('.clear-failed-btn').forEach(btn => btn.addEventListener('click', async () => {
    await clearFailedSync(btn.dataset.qk);
    await refreshAll();
  }));
}

async function renderConflicts() {
  const wrap = $('#conflictsContainer');
  if (!wrap) return;
  const rows = (await getAll(STORE_SYNC_CONFLICTS)).sort((a, b) => String(b.detectedAt || '').localeCompare(String(a.detectedAt || ''), 'id'));
  if (!rows.length) {
    wrap.className = 'stack-list empty-state';
    wrap.textContent = 'Belum ada konflik data.';
    return;
  }
  wrap.className = 'stack-list';
  wrap.innerHTML = rows.map(r => `
    <div class="report-card">
      <h4>${escapeHtml(normalizeEntityType(r.entityType))} • ${escapeHtml(r.entityId)}</h4>
      <div class="report-meta">Terdeteksi: ${escapeHtml(formatDateTimeID(r.detectedAt))}</div>
      <div>${escapeHtml(r.reason || '-')}</div>
      <div class="report-meta" style="margin-top:8px;">Lokal: ${escapeHtml(r.localDevice || '-')} (${escapeHtml(formatDateTimeID(r.localUpdatedAt) || '-')})</div>
      <div class="report-meta">Online: ${escapeHtml(r.remoteDevice || '-')} (${escapeHtml(formatDateTimeID(r.remoteUpdatedAt) || '-')})</div>
      <div class="report-actions">
        <span class="status-chip conflict">Konflik data</span>
      </div>
    </div>
  `).join('');
}

async function retryFailedSync(queueKey) {
  const failed = await getByKey(STORE_SYNC_FAILED, queueKey);
  if (!failed) return;
  await clearFailedSync(queueKey);
  const [entityType, entityId] = String(queueKey).split(':');
  const storeName = entityType === 'report' ? STORE_REPORTS : STORE_PARTICIPANTS;
  const row = await getByKey(storeName, entityId);
  if (row) {
    row.syncStatus = 'pending';
    await put(storeName, row);
    await enqueueSync(entityType, entityId, row.updatedAt || new Date().toISOString());
  }
  await refreshAll();
  await pushPendingQueue({ silent: false, stopOnError: false });
}

async function retryAllFailedSync() {
  const failed = await getAll(STORE_SYNC_FAILED);
  for (const item of failed) {
    const [entityType, entityId] = String(item.queueKey).split(':');
    const storeName = entityType === 'report' ? STORE_REPORTS : STORE_PARTICIPANTS;
    const row = await getByKey(storeName, entityId);
    if (!row) continue;
    row.syncStatus = 'pending';
    await put(storeName, row);
    await enqueueSync(entityType, entityId, row.updatedAt || new Date().toISOString());
    await clearFailedSync(item.queueKey);
  }
  await refreshAll();
  await pushPendingQueue({ silent: false, stopOnError: false });
}

function normalizePulledReport(row) {
  let mentees = [];
  try {
    mentees = Array.isArray(row.mentees) ? row.mentees : JSON.parse(row.mentees_json || '[]');
  } catch (_) {
    mentees = [];
  }
  return {
    id: String(row.id || '').trim(),
    visitDate: String(row.visitDate || '').trim(),
    location: String(row.location || '').trim(),
    mentees,
    summary: String(row.summary || '').trim(),
    activityReview: String(row.activityReview || '').trim(),
    technicalUnderstanding: String(row.technicalUnderstanding || '').trim(),
    fieldObservation: String(row.fieldObservation || '').trim(),
    microTeaching: String(row.microTeaching || '').trim(),
    livingCondition: String(row.livingCondition || '').trim(),
    motivationFuture: String(row.motivationFuture || '').trim(),
    resultsObtained: String(row.resultsObtained || '').trim(),
    followUp: String(row.followUp || '').trim(),
    specialNotes: String(row.specialNotes || '').trim(),
    createdAt: String(row.createdAt || row.updatedAt || new Date().toISOString()).trim(),
    updatedAt: String(row.updatedAt || new Date().toISOString()).trim(),
    sourceDevice: String(row.sourceDevice || '').trim(),
    syncStatus: 'pulled'
  };
}

function isIncomingNewer(incoming, existing) {
  return new Date(incoming || 0).getTime() >= new Date(existing || 0).getTime();
}

async function appsScriptRequest(params) {
  try {
    return await postViaIframe(params, { timeoutMs: 30000 });
  } catch (postError) {
    try {
      return await jsonpGetLegacy(params);
    } catch (jsonpError) {
      const err = new Error(
        'Gagal terhubung ke Apps Script. POST iframe gagal: ' + (postError?.message || postError) +
        ' | JSONP cadangan gagal: ' + (jsonpError?.message || jsonpError)
      );
      err.postError = postError;
      err.jsonpError = jsonpError;
      throw err;
    }
  }
}

function jsonpGetLegacy(params) {
  return new Promise((resolve, reject) => {
    const callbackName = '__jsonp_cb_' + Math.random().toString(36).slice(2);
    const script = document.createElement('script');
    const url = new URL(GAS_WEBAPP_URL);
    Object.entries({ ...params, callback: callbackName, _: Date.now() }).forEach(([k, v]) => url.searchParams.set(k, v));
    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error('JSONP timeout. Pastikan URL Web App benar dan deploy Apps Script sudah publik.'));
    }, 20000);

    function cleanup() {
      clearTimeout(timeout);
      delete window[callbackName];
      script.remove();
    }

    window[callbackName] = (data) => {
      cleanup();
      resolve(data);
    };
    script.onerror = () => {
      cleanup();
      reject(new Error('Gagal memanggil Apps Script via JSONP.'));
    };
    script.src = url.toString();
    document.body.appendChild(script);
  });
}

function postViaIframe(params, options = {}) {
  return new Promise((resolve, reject) => {
    const opId = 'op_' + Math.random().toString(36).slice(2);
    const iframe = document.createElement('iframe');
    const form = document.createElement('form');
    iframe.name = opId;
    iframe.style.display = 'none';
    form.method = 'POST';
    form.action = GAS_WEBAPP_URL;
    form.target = opId;
    form.style.display = 'none';

    const payload = { ...params, opId };
    Object.entries(payload).forEach(([key, value]) => {
      const input = document.createElement('input');
      input.type = 'hidden';
      input.name = key;
      input.value = String(value ?? '');
      form.appendChild(input);
    });

    const timeoutMs = Number(options.timeoutMs || 30000);
    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error('Sinkron POST timeout. Cek deploy Apps Script dan koneksi internet.'));
    }, timeoutMs);

    function cleanup() {
      clearTimeout(timeout);
      postMessageResolvers.delete(opId);
      iframe.remove();
      form.remove();
    }

    postMessageResolvers.set(opId, {
      resolve: (data) => { cleanup(); resolve(data); },
      reject: (error) => { cleanup(); reject(error); }
    });

    document.body.appendChild(iframe);
    document.body.appendChild(form);
    form.submit();
  });
}

function handlePostMessageResponse(event) {
  const data = event.data;
  if (!data || data.__visitSync !== true || !data.opId) return;
  const pending = postMessageResolvers.get(data.opId);
  if (!pending) return;
  pending.resolve(data.payload || { ok: false, message: 'Respon kosong dari Apps Script.' });
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
  const headers = rows.shift().map(h => String(h).trim());
  return rows.map(r => {
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
    Status_Sinkron: humanSyncStatus(r.syncStatus),
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
    Group: p.group,
    Status_Sinkron: humanSyncStatus(p.syncStatus)
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
  const maxWidths = headers.map((header) => Math.min(45, Math.max(
    String(header).length,
    ...rows.map(r => String(r[header] ?? '').length)
  ) + 2));

  const colsXml = maxWidths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');
  const rowsXml = matrix.map((row, rowIndex) => {
    const cells = row.map((value, colIndex) => {
      const ref = colName(colIndex + 1) + (rowIndex + 1);
      const styleId = rowIndex === 0 ? 1 : 0;
      return `<c r="${ref}" t="inlineStr" s="${styleId}"><is><t xml:space="preserve">${xmlEscape(String(value ?? ''))}</t></is></c>`;
    }).join('');
    return `<row r="${rowIndex + 1}">${cells}</row>`;
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
