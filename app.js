/** ===========================================================
 * Laporan Harian Mentee - Frontend (OFFLINE-FIRST)
 * =========================================================== */

/* ===================== 0) CONFIG ===================== */
// >>>>>>> GANTI INI DENGAN URL WEB APP MU <<<<<<<
const GAS_URL = 'https://script.google.com/macros/s/AKfycbypYA9dA9uQbMjYuMYdVTtu6nkMLaVMizeQMmrWDjI8m9Hhz5zxdFzcyrwogFpz_Pt05g/exec';

/* ===================== 1) LOADER ===================== */
const Loader = (() => {
  let overlay, box, spinner, label, progressWrap, progressBar;
  let active = 0;
  function init(){
    if (overlay) return;
    overlay = document.createElement('div');
    overlay.id = 'global-loader';
    overlay.style.cssText = `
      display:none; position:fixed; inset:0; z-index:3000;
      background: rgba(0,0,0,.25); backdrop-filter: blur(1px);
      align-items:center; justify-content:center; font-family:inherit;`;
    box = document.createElement('div');
    box.style.cssText = `width:min(420px,90vw); background:#fff; border-radius:12px; box-shadow:0 10px 30px rgba(0,0,0,.25); padding:18px;`;
    const row = document.createElement('div');
    row.style.cssText = `display:flex; gap:12px; align-items:center; margin-bottom:12px;`;
    spinner = document.createElement('div');
    spinner.style.cssText = `width:28px; height:28px; border-radius:50%; border:3px solid #e0e0e0; border-top-color:var(--primary); animation:spin .8s linear infinite;`;
    const style = document.createElement('style'); style.textContent = `@keyframes spin{to{transform:rotate(360deg)}}`; document.head.appendChild(style);
    label = document.createElement('div'); label.textContent = 'Memproses...'; label.style.cssText = `font-weight:600; color:#333;`;
    progressWrap = document.createElement('div'); progressWrap.style.cssText = `height:8px; background:#f1f1f1; border-radius:999px; overflow:hidden;`;
    progressBar = document.createElement('div'); progressBar.style.cssText = `height:100%; width:0%; background:linear-gradient(90deg,var(--primary),var(--primary-light)); transition:width .25s ease;`;
    progressWrap.appendChild(progressBar);
    row.appendChild(spinner); row.appendChild(label);
    box.appendChild(row); box.appendChild(progressWrap);
    overlay.appendChild(box); document.body.appendChild(overlay);
  }
  function show(text){ init(); active++; label.textContent = text || 'Memproses...'; progressBar.style.width = '10%'; overlay.style.display = 'flex'; }
  function setProgress(pct){ if (!progressBar) return; progressBar.style.width = `${Math.max(0,Math.min(100,+pct||0))}%`; }
  function hide(){ active = Math.max(0, active-1); if (active===0 && overlay){ overlay.style.display='none'; setProgress(0); } }
  async function withLoading(text, promiseLike){ show(text); try{ return await promiseLike; } finally{ hide(); } }
  return { init, show, hide, setProgress, withLoading };
})();
document.addEventListener('DOMContentLoaded', Loader.init);

function toast(msg){ alert(msg); }
function $(sel){ return document.querySelector(sel); }
function $all(sel){ return Array.from(document.querySelectorAll(sel)); }

function lockButton(btn){
  if (!btn) return () => {};
  const prev = { disabled: btn.disabled, pe: btn.style.pointerEvents, op: btn.style.opacity };
  btn.disabled = true; btn.style.pointerEvents = 'none'; btn.style.opacity = '0.6';
  return () => { btn.disabled = prev.disabled; btn.style.pointerEvents = prev.pe; btn.style.opacity = prev.op; };
}

/* ===================== 2) API ===================== */
async function apiGet(params = {}) {
  const url = new URL(GAS_URL);
  Object.entries(params).forEach(([k,v]) => url.searchParams.set(k, v));
  const r = await fetch(url.toString(), { method:'GET' });
  return r.json();
}
async function apiPost(body = {}) {
  const form = new URLSearchParams();
  form.append('payload', JSON.stringify(body));
  const r = await fetch(GAS_URL, { method:'POST', body: form });
  return r.json();
}

/* ===================== 3) UTIL TANGGAL ===================== */
function dd(v){ return String(v).padStart(2,'0'); }
function toDDMMYYYY(d){
  const dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt)) return '';
  return `${dd(dt.getDate())}/${dd(dt.getMonth()+1)}/${dt.getFullYear()}`;
}
function parseSendDateTime(str){
  // Input dari flatpickr format: "dd/mm/yyyy HH:ii"
  const m = String(str||'').match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})$/);
  if (m) {
    return { send_date: `${m[1]}/${m[2]}/${m[3]}`, send_time: `${m[4]}:${m[5]}` };
  }
  // fallback: sekarang
  const now = new Date();
  return { send_date: toDDMMYYYY(now), send_time: `${dd(now.getHours())}:${dd(now.getMinutes())}` };
}
function fmtWIBddmmyyyy(input){
  if (!input) return '';
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(input)) return input;
  const d = new Date(input);
  if (!isNaN(d)) return toDDMMYYYY(new Date(d.getTime()+7*3600*1000));
  const t = String(input).replace(/[-.]/g,'/'); const parts = t.split('/');
  if (parts.length===3){
    const day = dd(parts[0]); const mon = dd(parts[1]); const yr = parts[2].length===2 ? ('20'+parts[2]) : parts[2];
    return `${day}/${mon}/${yr}`;
  }
  return input;
}

/* ===================== 3a) UTIL WAKTU & SORT ===================== */
// Paksa apapun (string ISO, Date, angka fraksi Excel, "HH:mm(:ss)") jadi "HH:mm"
function coerceHHMM(val){
  if (val instanceof Date && !isNaN(val)) return `${dd(val.getHours())}:${dd(val.getMinutes())}`;
  const s = String(val||'').trim();

  // ISO / "1899-..." → ambil jam-menitnya
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(s) || /^1899-/.test(s)){
    const d = new Date(s);
    if (!isNaN(d)) return `${dd(d.getHours())}:${dd(d.getMinutes())}`;
  }

  // "H:MM" atau "HH:MM(:SS)"
  const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (m){
    const h = Math.max(0, Math.min(23, +m[1]||0));
    const i = Math.max(0, Math.min(59, +m[2]||0));
    return `${dd(h)}:${dd(i)}`;
  }

  // Angka fraksi (Excel time) 0..1
  const num = Number(s);
  if (!Number.isNaN(num) && num>=0 && num<1){
    const minutes = Math.round(num*24*60);
    const h = Math.floor(minutes/60);
    const i = minutes%60;
    return `${dd(h)}:${dd(i)}`;
  }

  return ''; // tidak bisa dipaksa
}


// gabungkan dd/mm/yyyy + hh:mm → Date lokal
function dateFrom(ddmmyyyy, hhmm){
  const d = parseDDMMYYYY_toLocalDate(ddmmyyyy);
  if (!d) return null;
  const m = String(hhmm||'').match(/^(\d{2}):(\d{2})$/);
  const H = m? +m[1] : 0, I = m? +m[2] : 0;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), H, I, 0, 0);
}

// state sorting untuk tabel tabular
const SORT = {
  col: 'send_dt', // send_dt | nik | nama | report_date | send_date | send_time | score
  dir: 'desc'     // 'asc' | 'desc'
};

// kecil: update info antrean pending
function renderQueueInfo(){
  const n = DB.getPending().length;
  const el1 = document.getElementById('pending-count');
  if (el1) el1.textContent = String(n);
  const el2 = document.getElementById('pending-count-tabular');
  if (el2) el2.textContent = String(n);
}


/* ===================== 3b) UTIL FILTER & TANGGAL BULAN ===================== */
const FILTER = {
  program: '', group: '', region: '', unit: '', divisi: '',
  month: null,  // Date object (tanggal apapun dalam bulan tersebut)
};

function firstDayOfMonth(dt){ return new Date(dt.getFullYear(), dt.getMonth(), 1); }
function daysInMonth(dt){ return new Date(dt.getFullYear(), dt.getMonth()+1, 0).getDate(); }
function isSunday(y,m,d){ return new Date(y, m, d).getDay() === 0; } // m = 0..11

function uniq(arr){ return Array.from(new Set(arr.filter(v => v && String(v).trim() !== ''))); }
function byNikMap(arr){ const m={}; arr.forEach(x=>{ if(x && x.nik) m[x.nik]=x; }); return m; }

function applyReportFilters(rows){
  // Join atribut peserta
  const pMap = byNikMap(DB.getParticipants());
  const enriched = rows.map(r => ({...r, ...(pMap[r.nik]||{})}));

  // Filter bulan
  let out = enriched;
  if (FILTER.month){
    const y = FILTER.month.getFullYear();
    const m = FILTER.month.getMonth(); // 0..11
    out = out.filter(r=>{
      const rd = parseDDMMYYYY_toLocalDate(fmtWIBddmmyyyy(r.report_date));
      return rd && rd.getFullYear()===y && rd.getMonth()===m;
    });
  }
  // Filter dimensi
  const by = { program:'program', group:'group', region:'region', unit:'unit', divisi:'divisi' };
  Object.entries(by).forEach(([fk, field])=>{
    const val = (FILTER[fk]||'').trim();
    if (val) out = out.filter(x => String(x[field]||'').toLowerCase() === val.toLowerCase());
  });
  return out;
}

function buildFilterOptionsFromParticipants(){
  const ps = DB.getParticipants().filter(p => String(p.is_active)!=='false');
  const programs = uniq(ps.map(p=>p.program));
  const groups   = uniq(ps.map(p=>p.group));
  const regions  = uniq(ps.map(p=>p.region));
  const units    = uniq(ps.map(p=>p.unit));
  const divis    = uniq(ps.map(p=>p.divisi));

  function fillSelect(id, list){
    const el = document.getElementById(id); if (!el) return;
    const cur = el.value;
    el.innerHTML = `<option value="">Semua</option>` + list.map(v=>`<option value="${String(v)}">${String(v)}</option>`).join('');
    if (cur) el.value = cur; // pertahankan pilihan bila ada
  }
  fillSelect('filter-program', programs);
  fillSelect('filter-group',   groups);
  fillSelect('filter-region',  regions);
  fillSelect('filter-unit',    units);
  fillSelect('filter-divisi',  divis);
}

function getSelectedMonthRange(){
  if (!FILTER.month){
    const now = new Date();
    FILTER.month = new Date(now.getFullYear(), now.getMonth(), 1);
  }
  const y = FILTER.month.getFullYear();
  const m = FILTER.month.getMonth();
  const nDays = daysInMonth(FILTER.month);
  return { y, m, nDays };
}



/* ======= PENILAIAN (SCORE) ======= */
// Parser dd/mm/yyyy → Date lokal (WIB)
function parseDDMMYYYY_toLocalDate(s){
  const m = String(s||'').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return null;
  const d = parseInt(m[1],10), mo = parseInt(m[2],10)-1, y = parseInt(m[3],10);
  return new Date(y, mo, d, 0, 0, 0, 0);
}
// Builder dd/mm/yyyy + "HH:MM[:SS]" → Date lokal
function buildLocalDateTime(ddmmyyyy, hhmmss){
  const base = parseDDMMYYYY_toLocalDate(ddmmyyyy);
  if (!base) return null;
  const t = String(hhmmss||'').trim();
  let hh=0, mm=0, ss=0;
  if (/^\d{2}:\d{2}(:\d{2})?$/.test(t)){
    const p = t.split(':'); hh = +p[0]; mm = +p[1]; ss = p[2]? +p[2] : 0;
  }
  return new Date(base.getFullYear(), base.getMonth(), base.getDate(), hh, mm, ss, 0);
}
// Hitung skor & warna harian berdasarkan report_date, send_date, send_time
function computeDailyScore(report_date_ddmmyyyy, send_date_ddmmyyyy, send_time_hhmm){
  // Jika tidak ada tanggal kirim → tidak bernilai & tanpa warna
  if (!send_date_ddmmyyyy || !send_time_hhmm) return { score: '', color: 'none' };

  const H = parseDDMMYYYY_toLocalDate(report_date_ddmmyyyy);
  const sent = buildLocalDateTime(send_date_ddmmyyyy, send_time_hhmm);
  if (!H || !sent) return { score: '', color: 'none' };

  // Ambang H+1 (08:00, 12:00, 23:59:59)
  const H1_0800 = new Date(H); H1_0800.setDate(H1_0800.getDate()+1); H1_0800.setHours(8,0,0,0);
  const H1_1200 = new Date(H); H1_1200.setDate(H1_1200.getDate()+1); H1_1200.setHours(12,0,0,0);
  const H1_2359 = new Date(H); H1_2359.setDate(H1_2359.getDate()+1); H1_2359.setHours(23,59,59,999);

  if (sent <= H1_0800) return { score: 4, color: 'green' };
  if (sent <= H1_1200) return { score: 3, color: 'yellow' };
  if (sent <= H1_2359) return { score: 2, color: 'red' };
  return { score: 1, color: 'black' };
}
// Kelas/inline style untuk badge harian (4/3/2/1/none)
function dailyScoreBadge(score){
  if (score === 4) return `<span class="badge badge-green">4</span>`;
  if (score === 3) return `<span class="badge badge-yellow">3</span>`;
  if (score === 2) return `<span class="badge badge-red">2</span>`;
  if (score === 1) return `<span class="badge" style="background:#000;color:#fff;">1</span>`;
  return ''; // tanpa warna
}
// Warna kolom Nilai total (rekap peserta)
function totalScoreColor(total){
  if (total >= 90) return 'green';
  if (total >= 76) return 'yellow';
  if (total >= 60) return 'red';
  if (total > 0)  return 'black';
  return 'none';
}
function totalScoreBadge(total){
  const c = totalScoreColor(total);
  if (c === 'green')  return `<span class="badge badge-green">${total}</span>`;
  if (c === 'yellow') return `<span class="badge badge-yellow">${total}</span>`;
  if (c === 'red')    return `<span class="badge badge-red">${total}</span>`;
  if (c === 'black')  return `<span class="badge" style="background:#000;color:#fff;">${total}</span>`;
  return String(total||0);
}

function localId(){ return 'loc_' + Math.random().toString(36).slice(2,10); }

// ⬇️ ADD HERE: build lightweight modal for editing a report
function openEditReportModal({ key, id }) {
  const data = DB.getReportByKey(key);
  if (!data) return toast('Data tidak ditemukan.');

  // backdrop
  const backdrop = document.createElement('div');
  backdrop.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:2100;display:flex;align-items:center;justify-content:center;';

  // kartu modal
  const card = document.createElement('div');
  card.style.cssText = 'width:min(520px,95vw);background:#fff;border-radius:10px;box-shadow:0 10px 30px rgba(0,0,0,.25);overflow:hidden;';
  card.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center;padding:12px 16px;border-bottom:1px solid #e5e7eb;">
      <h3 style="margin:0;font-size:1.05rem;color:var(--primary);">Edit Laporan</h3>
      <button class="btn btn-outline btn-sm" id="edit-close">Tutup</button>
    </div>
    <div style="padding:16px;">
      <form id="edit-form">
        <div class="form-group"><label>NIK</label><input class="form-control" id="e-nik" value="${data.nik||''}" disabled></div>
        <div class="form-group"><label>Tanggal Laporan</label><input class="form-control" id="e-report-date" placeholder="dd/mm/yyyy" value="${fmtWIBddmmyyyy(data.report_date)||''}"></div>
        <div class="form-group"><label>Tanggal & Jam Kirim</label><input class="form-control" id="e-send-datetime" placeholder="dd/mm/yyyy hh:mm" value=""></div>
        <div class="form-group"><label>Nilai</label><input class="form-control" id="e-score" type="number" min="0" max="100" step="1" value="${data.score||''}"></div>
        <div class="form-group" style="display:flex;gap:8px;justify-content:flex-end;">
          <button type="button" class="btn btn-outline" id="edit-cancel">Batal</button>
          <button type="submit" class="btn btn-primary"><i class="fas fa-save"></i> Simpan</button>
        </div>
      </form>
    </div>
  `;
  backdrop.appendChild(card);
  document.body.appendChild(backdrop);

  const doClose = ()=> backdrop.remove();
  card.querySelector('#edit-close').onclick = doClose;
  card.querySelector('#edit-cancel').onclick = doClose;

  // ===== Inisialisasi flatpickr yang aman =====
  if (window.flatpickr) {
    flatpickr.localize(flatpickr.l10ns.id);

    // Tanggal laporan
    const rpInput = card.querySelector('#e-report-date');
    flatpickr(rpInput, {
      dateFormat:'d/m/Y',
      allowInput:true,
      defaultDate: fmtWIBddmmyyyy(data.report_date) || undefined
    });

    // Tanggal & jam kirim (gunakan Date object jika valid)
    const sdInput     = card.querySelector('#e-send-datetime');
    const sendDateStr = fmtWIBddmmyyyy(data.send_date) || fmtWIBddmmyyyy(data.report_date);
    const sendTimeStr = coerceHHMM(data.send_time);
    const defDT       = (()=>{
      if (!sendDateStr || !sendTimeStr) return null;
      const dt = dateFrom(sendDateStr, sendTimeStr);
      return (dt && !isNaN(dt)) ? dt : null;
    })();

    flatpickr(sdInput, {
      enableTime:true,
      time_24hr:true,
      dateFormat:'d/m/Y H:i',
      allowInput:true,
      defaultDate: defDT || undefined
    });

    // Jika tidak ada Date valid, tampilkan gabungan string agar user bisa koreksi manual
    if (!defDT) sdInput.value = [sendDateStr, sendTimeStr].filter(Boolean).join(' ').trim();
  }

  // ===== submit edit =====
  card.querySelector('#edit-form').addEventListener('submit', async (e)=>{
    e.preventDefault();
    const btn = e.submitter || card.querySelector('button[type="submit"]');
    const unlock = lockButton(btn);

    // 1) Ambil nilai input
    const newReportDate = card.querySelector('#e-report-date').value.trim();
    const newSendDT     = card.querySelector('#e-send-datetime').value.trim();
    const manualInput   = card.querySelector('#e-score').value.trim(); // opsional

    if (!newReportDate) { unlock(); return toast('Tanggal Laporan wajib diisi.'); }

    // 2) Normalisasi tanggal & jam
    //    - parse gabungan "dd/mm/yyyy HH:mm"
    const parsed = parseSendDateTime(newSendDT);
    const send_date = fmtWIBddmmyyyy(parsed.send_date);
    const send_time = coerceHHMM(parsed.send_time);
    const report_date = fmtWIBddmmyyyy(newReportDate);
    const nik = data.nik;

    // 3) Hitung skor otomatis → pakai jam yg sudah dinormalisasi
    const autoScore = computeDailyScore(report_date, send_date, send_time).score; // 4/3/2/1 atau ''
    const finalScore = (manualInput !== '') ? Number(manualInput) : autoScore;

    // 4) Jika tanggal laporan berubah, _key juga berubah
    const newKey = nik + '|' + report_date;

    // 5) Update lokal (optimistic)
    DB.deleteReportByKey(key);
    DB.upsertReportLocal({
      ...data,
      report_date,
      send_date,
      send_time,
      score: finalScore,
      _key: newKey,
      isSynced: false,
      synced_at: ''
    });
    renderReportsTable();
    renderStats();
    if (typeof renderMonitoringTable === 'function') renderMonitoringTable();

    // 6) (opsional) kirim ke server sekarang; kalau offline akan tetap aman di lokal
    try{
      const res = await Loader.withLoading('Menyimpan perubahan...', apiPost({
        action:'upsertReport',
        nik, report_date, send_date, send_time, score: finalScore
      }));
      if (res && res.ok){
        DB.updateReportByKey(newKey, {
          id: res.id || data.id,
          isSynced: true,
          synced_at: new Date().toISOString()
        });
        renderReportsTable(); renderStats();
        if (typeof renderMonitoringTable === 'function') renderMonitoringTable();
        toast('Perubahan disimpan.');
      } else {
        toast('Perubahan disimpan lokal. Akan disinkron saat online.');
      }
    } catch (_){
      toast('Offline: perubahan disimpan lokal dan akan disinkron.');
    } finally {
      unlock();
      doClose();
    }
  });
}

// ⬇️ ADD HERE: delete flow
async function deleteReportFlow({ key, id }) {
  const row = DB.getReportByKey(key);
  if (!row) return toast('Data tidak ditemukan.');
  if (!confirm(`Hapus laporan NIK ${row.nik} tanggal ${fmtWIBddmmyyyy(row.report_date)} ?`)) return;

  // Hapus lokal dulu (optimistic)
  DB.deleteReportByKey(key);
  renderReportsTable();
  renderStats();

  // Coba hapus ke server (gunakan id bila ada; fallback nik+report_date)
  try{
    const res = await Loader.withLoading('Menghapus laporan...', apiPost({
      action: 'deleteReport',
      id: row.id || id || '',
      nik: row.nik,
      report_date: row.report_date
    }));
    if (res && res.ok){
      toast('Laporan dihapus.');
    } else {
      // kalau backend belum dukung delete: informasikan
      toast('Laporan dihapus lokal. Backend tidak mengonfirmasi penghapusan.');
    }
  } catch(_){
    toast('Offline: laporan terhapus di perangkat. Sinkron nanti bisa mengembalikan dari server jika belum terhapus di sana.');
  }
}


/* ===================== 4) DB (localStorage) ===================== */
const DB = (() => {
  const KEYS = { P:'kmp.participants', R:'kmp.reports', RP:'kmp.reports_pending', RF:'kmp.reports_failed', U:'kmp.users', M:'kmp.meta' };
  const read  = (k,def)=>{ try{ return JSON.parse(localStorage.getItem(k)||JSON.stringify(def)); }catch(_){ return def; } };
  const write = (k,v)=> localStorage.setItem(k, JSON.stringify(v));

  const getParticipants = () => read(KEYS.P, []);
  const setParticipants = (rows) => write(KEYS.P, Array.isArray(rows)? rows: []);
  const getReports = () => read(KEYS.R, []);
  const setReports = (rows) => write(KEYS.R, Array.isArray(rows)? rows: []);
  function upsertReportLocal(obj){
    const rows = getReports();
    const key = obj._key || (obj.nik + '|' + obj.report_date);
    const i = rows.findIndex(r => (r._key || (r.nik+'|'+r.report_date)) === key);
    if (i>-1) rows[i] = {...rows[i], ...obj}; else rows.push({...obj});
    setReports(rows);
  }
  const getPending = () => read(KEYS.RP, []);
  const setPending = (rows) => write(KEYS.RP, Array.isArray(rows)? rows: []);
  const addPending = (obj) => setPending([...getPending(), obj]);
  const removePendingByKey = (_key) => setPending(getPending().filter(r => r._key !== _key));

  // ⬇️ NEW: antrean gagal
  const getFailed  = () => read(KEYS.RF, []);
  const setFailed  = (rows) => write(KEYS.RF, Array.isArray(rows)? rows: []);
  const addFailed  = (obj) => setFailed([...getFailed(), obj]);
  const removeFailedByKey = (_key) => setFailed(getFailed().filter(r => r._key !== _key));

  const getUsers = () => read(KEYS.U, []);
  const setUsers = (rows) => write(KEYS.U, Array.isArray(rows)? rows: []);
  const getMeta = () => read(KEYS.M, { lastSyncAt:null, version:1 });
  const setMeta = (meta) => write(KEYS.M, {...getMeta(), ...meta});

    function getReportByKey(_key){
    return getReports().find(r => (r._key || (r.nik+'|'+r.report_date)) === _key);
    }
    function deleteReportByKey(_key){
    const rows = getReports().filter(r => (r._key || (r.nik+'|'+r.report_date)) !== _key);
    setReports(rows);
    }
    function updateReportByKey(_key, patch){
    const rows = getReports();
    const i = rows.findIndex(r => (r._key || (r.nik+'|'+r.report_date)) === _key);
    if (i > -1) {
        rows[i] = { ...rows[i], ...patch };
        setReports(rows);
    }
    }

  return { getParticipants, setParticipants, getReportByKey, deleteReportByKey, updateReportByKey, getReports, setReports, upsertReportLocal, getPending, setPending, addPending, removePendingByKey,  getFailed, setFailed, addFailed, removeFailedByKey, getUsers, setUsers, getMeta, setMeta };
})();

/* ===================== 5) RENDER ===================== */
function renderStats(){
  const reports = DB.getReports();
  const totalSync = reports.filter(r => r.isSynced || r.synced_at).length;
  const totalPeserta = DB.getParticipants().filter(p => String(p.is_active)!=='false').length;
  const scores = reports.map(r => Number(r.score)).filter(n => !Number.isNaN(n));
  const avg = scores.length ? (scores.reduce((a,b)=>a+b,0)/scores.length).toFixed(1) : '-';
  const cards = $all('.stat-card .stat-info h3');
  if (cards[0]) cards[0].textContent = String(totalPeserta||0);
  if (cards[1]) cards[1].textContent = String(avg);
  if (cards[2]) cards[2].textContent = String(reports.length - totalSync);
  if (cards[3]) cards[3].textContent = String(totalSync);
}
function renderReportsTable(){
  const tbody = document.querySelector('#tabular-content table.data-table tbody'); if (!tbody) return;

  // ambil & filter
  const allRows = DB.getReports();
  let rows = applyReportFilters(allRows);

  // siapkan join nama peserta
  const pMap = byNikMap(DB.getParticipants());

  // siapkan field bantu untuk sorting default: send_dt (gabungan tanggal kirim + jam)
  rows = rows.map(r=>{
    const send_time_fmt = coerceHHMM(r.send_time||'');
    const send_date_fmt = fmtWIBddmmyyyy(r.send_date||'');
    const send_dt = dateFrom(send_date_fmt, send_time_fmt) || dateFrom(fmtWIBddmmyyyy(r.report_date), '00:00');
    const report_dt = parseDDMMYYYY_toLocalDate(fmtWIBddmmyyyy(r.report_date));
    return { ...r, _send_time_fmt: send_time_fmt, _send_date_fmt: send_date_fmt, _send_dt: send_dt, _report_dt: report_dt };
  });

  // sorting
  const cmp = {
    'send_dt': (a,b)=> (a._send_dt||0) - (b._send_dt||0),
    'report_date': (a,b)=> (a._report_dt||0) - (b._report_dt||0),
    'nik': (a,b)=> String(a.nik||'').localeCompare(String(b.nik||'')),
    'nama': (a,b)=> String(pMap[a.nik]?.nama || a.nama || '').localeCompare(String(pMap[b.nik]?.nama || b.nama || '')),
    'send_date': (a,b)=> String(a._send_date_fmt||'').localeCompare(String(b._send_date_fmt||'')),
    'send_time': (a,b)=> String(a._send_time_fmt||'').localeCompare(String(b._send_time_fmt||'')),
    'score': (a,b)=> Number(a.score||0) - Number(b.score||0)
  };
  const sorter = cmp[SORT.col] || cmp['send_dt'];
  rows.sort(sorter);
  if (SORT.dir==='desc') rows.reverse();

  // render
  tbody.innerHTML = '';
  rows.forEach((r, idx)=>{
    const key = r._key || (r.nik + '|' + r.report_date);
    const P = pMap[r.nik] || {};
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${idx+1}</td>
      <td>${r.nik||''}</td>
      <td>${P.nama||r.nama||''}</td>
      <td>${fmtWIBddmmyyyy(r.report_date)||''}</td>
      <td>${r._send_date_fmt||''}</td>
      <td>${r._send_time_fmt||''}</td>
      <td>${dailyScoreBadge(Number(r.score))}</td>
      <td><span class="badge ${(r.isSynced||r.synced_at)?'badge-green':'badge-yellow'}">
        ${(r.isSynced||r.synced_at)?'Tersinkron':'Belum Sync'}</span></td>
      <td>
        <button class="btn btn-outline btn-sm" data-act="edit" data-id="${r.id||''}" data-key="${key}"><i class="fas fa-edit"></i></button>
        <button class="btn btn-outline btn-sm" data-act="delete" data-id="${r.id||''}" data-key="${key}"><i class="fas fa-trash"></i></button>
      </td>`;
    tbody.appendChild(tr);
  });

  const info = document.querySelector('#tabular-content .pagination-info');
  if (info) info.textContent = `Menampilkan ${rows.length} data`;
  renderQueueInfo();
}


// ⬇️ ADD HERE: action handler for Edit/Delete in Tabular table
document.addEventListener('DOMContentLoaded', () => {
  const table = document.querySelector('#tabular-content table.data-table');
  if (!table) return;

  table.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-act]');
    if (!btn) return;
    const act = btn.getAttribute('data-act');
    const key = btn.getAttribute('data-key');   // nik|report_date
    const id  = btn.getAttribute('data-id') || '';

    if (act === 'edit') {
      openEditReportModal({ key, id });
    } else if (act === 'delete') {
      deleteReportFlow({ key, id });
    }
  });

  // sorting header pada Tabular
    const thSel = '#tabular-content table.data-table thead th';
    const ths = document.querySelectorAll(thSel);
    if (ths && ths.length){
    // mapping kolom -> key sort
    const map = { 1:'nik', 2:'nama', 3:'report_date', 4:'send_date', 5:'send_time', 6:'score' };
    ths.forEach((th, idx)=>{
        if (map[idx]) {
        th.style.cursor = 'pointer';
        th.setAttribute('title','Klik untuk sort');
        th.addEventListener('click', ()=>{
            if (SORT.col === map[idx]) {
            SORT.dir = SORT.dir === 'asc' ? 'desc' : 'asc';
            } else {
            SORT.col = map[idx]; SORT.dir = 'asc';
            }
            renderReportsTable();
        });
        }
    });
    }

});

function renderUsersTable(){
  const tbody = $('#users-tbody'); if (!tbody) return;
  const rows = DB.getUsers(); tbody.innerHTML = '';
  rows.forEach(u=>{
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${u.username||''}</td>
      <td>${u.nama||''}</td>
      <td>${u.role||''}</td>
      <td><span class="badge ${String(u.status||'').toLowerCase()==='aktif'?'badge-green':'badge-yellow'}">${u.status||''}</span></td>
      <td><button class="btn btn-outline btn-sm" data-act="edit" data-username="${u.username||''}"><i class="fas fa-pen"></i></button></td>`;
    tbody.appendChild(tr);
  });
}

function renderFailedSyncPage(){
  const tbody = document.querySelector('#failed-sync-tbody'); 
  if (!tbody) return;
  const rows = DB.getFailed();
  tbody.innerHTML = '';
  rows.forEach((r, i)=>{
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${i+1}</td>
      <td>${r.nik}</td>
      <td>${fmtWIBddmmyyyy(r.report_date)}</td>
      <td>${fmtWIBddmmyyyy(r.send_date)}</td>
      <td>${coerceHHMM(r.send_time)}</td>
      <td>${r.score||''}</td>
      <td>${String(r.reason||'')}</td>
      <td>
        <button class="btn btn-outline btn-sm" data-retry="${r._key}"><i class="fas fa-redo"></i></button>
        <button class="btn btn-outline btn-sm" data-remove="${r._key}"><i class="fas fa-times"></i></button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  // action
  tbody.onclick = async (e)=>{
    const btn = e.target.closest('button');
    if (!btn) return;
    if (btn.dataset.retry){
      const key = btn.dataset.retry;
      const row = DB.getFailed().find(x=>x._key===key);
      if (!row) return;
      try{
        const res = await Loader.withLoading('Mencoba sync ulang...', apiPost({
          action:'upsertReport',
          nik:row.nik, report_date:row.report_date, send_date:row.send_date, send_time:row.send_time, score:row.score||'',
          markSynced: true
        }));
        if (res && res.ok){
          DB.removeFailedByKey(key);
          DB.upsertReportLocal({ ...row, id: res.id || row.id, isSynced:true, synced_at:new Date().toISOString() });
          renderFailedSyncPage(); renderReportsTable(); renderStats(); renderQueueInfo();
          toast('Berhasil disinkron.');
        } else {
          toast('Masih gagal: ' + (res.error||'unknown'));
        }
      } catch(err){
        toast('Jaringan bermasalah: ' + String(err));
      }
    } else if (btn.dataset.remove){
      DB.removeFailedByKey(btn.dataset.remove);
      renderFailedSyncPage(); renderQueueInfo();
    }
  };
}

function renderMonitoringHead(){
  const thead = document.getElementById('monitoring-thead'); if(!thead) return;
  const { y, m, nDays } = getSelectedMonthRange();
  // baris header: kolom tetap + kolom tanggal 1..n + Nilai
  const fixedCols = ['No','NIK','Nama','Program','Divisi','Unit','Region','Group'];
  let ths = fixedCols.map(h=>`<th>${h}</th>`).join('');
  for (let d=1; d<=nDays; d++){ ths += `<th>${d}</th>`; }
  ths += `<th>Nilai</th>`;
  thead.innerHTML = `<tr>${ths}</tr>`;
}

function renderMonitoringTable(){
  renderMonitoringHead(); // bangun thead sesuai bulan
  const tbody = document.querySelector('#monitoring-content table.data-table tbody');
  if (!tbody) return;

  const { y, m, nDays } = getSelectedMonthRange();
  const participants = DB.getParticipants().filter(p => String(p.is_active)!=='false');

  // Index laporan bulan-terpilih: map[nk|dd/mm/yyyy] = score
  const reports = applyReportFilters(DB.getReports()); // sudah terfilter dimensi + bulan
  const repMap = {};
  for (const r of reports){
    const key = r.nik + '|' + fmtWIBddmmyyyy(r.report_date);
    repMap[key] = Number(r.score);
  }

  // Terapkan filter peserta (program/group/region/unit/divisi)
  const by = { program:'program', group:'group', region:'region', unit:'unit', divisi:'divisi' };
  const f = (p) => {
    return Object.entries(by).every(([fk, field])=>{
      const val = (FILTER[fk]||'').trim();
      return !val || String(p[field]||'').toLowerCase() === val.toLowerCase();
    });
  };
  const filteredP = participants.filter(f);

  tbody.innerHTML = '';
  filteredP.forEach((p, i) => {
    let total = 0;
    const tds = [];

    for (let day=1; day<=nDays; day++){
  const sdd = dd(day);      // pakai fungsi util dd() yang sudah ada
  const smm = dd(m+1);
  const yyyy = y;

  const ddmmyyyy = `${sdd}/${smm}/${yyyy}`;
  const key = p.nik + '|' + ddmmyyyy;
  const sc = repMap[key];

  const sunday = isSunday(y, m, day);
  const sunClass = sunday && (sc===undefined || sc==='' || Number.isNaN(sc)) ? 'sun-pink' : '';

  if (typeof sc === 'number' && !Number.isNaN(sc)){
    total += sc;
    let txtClass = 'txt-none';
    if (sc===4) txtClass='txt-green';
    else if (sc===3) txtClass='txt-yellow';
    else if (sc===2) txtClass='txt-red';
    else if (sc===1) txtClass='txt-black';

    tds.push(`
      <td class="cell-score ${sunClass}">
        ${dailyScoreBadge(sc)}
        <span class="score-text ${txtClass}">${sc}</span>
      </td>
    `);
  } else {
    tds.push(`<td class="cell-score ${sunClass}"></td>`);
  }
}
    const capped = Math.max(0, Math.min(100, Math.round(total)));
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${i+1}</td>
      <td>${p.nik||''}</td>
      <td>${p.nama||''}</td>
      <td>${p.program||''}</td>
      <td>${p.divisi||''}</td>
      <td>${p.unit||''}</td>
      <td>${p.region||''}</td>
      <td>${p.group||''}</td>
      ${tds.join('')}
      <td>${totalScoreBadge(capped)}</td>
    `;
    tbody.appendChild(tr);
  });

  // update info
  const info = document.querySelector('#monitoring-content .pagination-info');
  if (info) info.textContent = `Menampilkan ${filteredP.length} peserta untuk ${FILTER.month.toLocaleDateString('id-ID',{ month:'long', year:'numeric' })}`;
}


/* ===================== 6) AUTOSUGGEST ===================== */
async function runAutosuggest(q){
  const list = DB.getParticipants();
  const lower = q.toLowerCase();
  let result = list
    .filter(p => String(p.is_active)!=='false')
    .filter(p => (`${p.nik||''} ${p.nama||''}`).toLowerCase().includes(lower))
    .slice(0,25);

  if (!result.length && q.length >= 2){
    const res = await apiGet({action:'listParticipants', q});
    if (res.ok){
      const merged = [...list];
      (res.data||[]).forEach(p=>{
        const i = merged.findIndex(x=>x.nik===p.nik);
        if (i>-1) merged[i] = {...merged[i], ...p, is_active:true};
        else merged.push({...p, is_active:true});
      });
      DB.setParticipants(merged);
      result = res.data||[];
    }
  }
  return result;
}

/* ===================== 7) ROUTING & TABS ===================== */
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.sidebar-menu a').forEach(link => {
    link.addEventListener('click', function(e){
      e.preventDefault();
      document.querySelectorAll('.sidebar-menu a').forEach(i => i.classList.remove('active'));
      this.classList.add('active');
      const pageId = this.getAttribute('data-page') + '-page';
      document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
      const target = document.getElementById(pageId);
      if (target) target.classList.add('active');
      if (pageId==='settings-page') renderUsersTable();
      if (pageId==='report-page') renderReportsTable();
      if (pageId==='dashboard-page') renderStats();
      if (pageId==='monitoring-page') renderMonitoringTable();
      if (pageId==='failedsync-page') renderFailedSyncPage();
    });
  });
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', function(){
      const tabId = this.getAttribute('data-tab');
      document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
      this.classList.add('active');
      document.querySelectorAll('.tab-content').forEach(c=>c.classList.remove('active'));
      const target = document.getElementById(tabId + '-content');
      if (target) target.classList.add('active');
      if (tabId==='monitoring') renderMonitoringTable();
    });
  });
  // Bersihkan dummy angka/tabel awal
  $all('#dashboard-page table.data-table tbody, #report-page table.data-table tbody, #monitoring-content table.data-table tbody')
    .forEach(tb => tb.innerHTML = '');
  $all('.stat-card .stat-info h3').forEach(el => el.textContent = '0');
});

/* ===================== 8) FLATPICKR (Tanggal) ===================== */
document.addEventListener('DOMContentLoaded', () => {
  // Tanggal Laporan (default: kemarin)
  const rpt = $('#report-date');
  if (rpt && window.flatpickr) {
    const kemarin = new Date(); kemarin.setDate(kemarin.getDate()-1);
    flatpickr.localize(flatpickr.l10ns.id);
    flatpickr(rpt, {
      dateFormat: 'd/m/Y',
      defaultDate: kemarin,
      allowInput: true
    });
  }
  // Tanggal & Jam Kirim (default: sekarang)
  const snd = $('#send-datetime');
  if (snd && window.flatpickr) {
    flatpickr.localize(flatpickr.l10ns.id);
    flatpickr(snd, {
      enableTime: true,
      time_24hr: true,
      dateFormat: 'd/m/Y H:i',
      defaultDate: new Date(),
      minuteIncrement: 1,
      allowInput: true
    });
  }

    // Filter Bulan (Monitoring & Tabular)
  const fMonth = document.getElementById('filter-month');
  if (fMonth && window.flatpickr){
    flatpickr.localize(flatpickr.l10ns.id);
    flatpickr(fMonth, {
      dateFormat: 'F Y', altInput: false,
      defaultDate: new Date(),
      onChange: (sel)=>{
        const d = sel && sel[0] ? sel[0] : new Date();
        FILTER.month = new Date(d.getFullYear(), d.getMonth(), 1);
        renderReportsTable(); renderMonitoringTable(); renderStats();
      }
    });
    // set awal
    FILTER.month = new Date();
  }

  // Event perubahan filter lain
  ['filter-program','filter-group','filter-region','filter-unit','filter-divisi'].forEach(id=>{
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', ()=>{
      FILTER[id.replace('filter-','')] = el.value || '';
      renderReportsTable(); renderMonitoringTable(); renderStats();
    });
  });

});



/* ===================== 9) INPUT FORM ===================== */
const nikInput = $('#nik');
const nikSug   = $('#nik-suggestions');
const infoBox  = $('#participant-info');

if (nikInput && nikSug && infoBox) {
  nikInput.addEventListener('input', async function(){
    const q = this.value.trim();
    if (q.length < 2) { nikSug.style.display = 'none'; return; }
    const list = await runAutosuggest(q);
    if (!list.length){ nikSug.style.display='none'; return; }
    nikSug.innerHTML = '';
    list.forEach(p=>{
      const item = document.createElement('div');
      item.className = 'autosuggest-item';
      item.textContent = `${p.nik} - ${p.nama}`;
      item.addEventListener('click', ()=>{
        nikInput.value = p.nik;
        $('#info-nama').textContent    = p.nama || '-';
        $('#info-program').textContent = p.program || '-';
        $('#info-divisi').textContent  = p.divisi || '-';
        $('#info-unit').textContent    = p.unit || '-';
        $('#info-region').textContent  = p.region || '-';
        $('#info-group').textContent   = p.group || '-';
        infoBox.style.display = 'block';
        nikSug.style.display  = 'none';
      });
      nikSug.appendChild(item);
    });
    nikSug.style.display = 'block';
  });
}

const reportForm = $('#report-form');
if (reportForm) {
  reportForm.addEventListener('submit', async function(e){
    e.preventDefault();
    const submitBtn = this.querySelector('button[type="submit"]');
    const unlockBtn = lockButton(submitBtn);

    const nik = ( $('#nik')?.value || '' ).trim();
    const reportDateRaw = ( $('#report-date')?.value || '' ).trim();
    const sendDateTimeRaw = ( $('#send-datetime')?.value || '' ).trim();

    if (!nik || !reportDateRaw){
      unlockBtn(); return toast('NIK dan Tanggal Laporan wajib diisi.');
    }

    // Normalisasi tanggal & jam (jam wajib tampil hh:mm)
    const report_date = fmtWIBddmmyyyy(reportDateRaw); // dd/mm/yyyy
    const p = parseSendDateTime(sendDateTimeRaw);
    const send_date = fmtWIBddmmyyyy(p.send_date);
    const send_time = coerceHHMM(p.send_time);

    const { score } = computeDailyScore(report_date, send_date, send_time);
    const _key = nik + '|' + report_date;

    const localObj = {
      id: localId(),
      nik, report_date, send_date, send_time, score,
      isSynced: false, synced_at: '', _key,
      created_locally_at: new Date().toISOString()
    };

    // ⬇️ HANYA LOKAL + masukkan ke antrean pending
    DB.upsertReportLocal(localObj);
    DB.addPending(localObj);
    renderReportsTable(); renderStats(); renderQueueInfo();
    toast('Tersimpan lokal. Masuk antrean sinkron.');

    // reset form
    this.reset();
    if (infoBox) infoBox.style.display = 'none';
    const kemarin = new Date(); kemarin.setDate(kemarin.getDate()-1);
    if (window.flatpickr) {
      if ($('#report-date') && $('#report-date')._flatpickr) $('#report-date')._flatpickr.setDate(kemarin, true);
      if ($('#send-datetime') && $('#send-datetime')._flatpickr) $('#send-datetime')._flatpickr.setDate(new Date(), true);
    }
    unlockBtn();
  });
}


/* ===================== 10) SYNC (PUSH pending → PULL semua) ===================== */
async function syncAll(){
  const pending = DB.getPending();
  const failed  = DB.getFailed();

  // Kalau tidak ada yang pending, JANGAN push—langsung pull
  if (pending.length){
    Loader.show('Mengirim data pending...');
    let done = 0;
    for (const r of pending){
      try{
        // kirim dengan markSynced:true agar backend memberi synced_at
        const res = await apiPost({ 
          action:'upsertReport',
          nik:r.nik, report_date:r.report_date, send_date:r.send_date, send_time:r.send_time, score:r.score||'',
          markSynced: true
        });
        if (res && res.ok){
          // tandai lokal sukses
          DB.removePendingByKey(r._key);
          DB.upsertReportLocal({ ...r, id: res.id || r.id, isSynced:true, synced_at:new Date().toISOString() });
        } else {
          // pindah ke antrean gagal
          DB.removePendingByKey(r._key);
          DB.addFailed({ ...r, reason: (res && res.error) || 'unknown' });
        }
      } catch(err){
        // jaringan error → pindahkan ke failed
        DB.removePendingByKey(r._key);
        DB.addFailed({ ...r, reason: String(err) });
      } finally {
        done++; Loader.setProgress( Math.min(90, Math.round((done/pending.length)*40)+10) );
      }
    }
    Loader.hide();
  }

  // PULL sesudah push — dan JANGAN merusak status lokal yg sudah Tersinkron
  await Loader.withLoading('Menarik data dari server...', (async ()=>{
    const p = await apiGet({action:'listParticipants', q:''});
    if (p.ok) DB.setParticipants(p.data||[]);

    const r = await apiGet({action:'listReports'});
    if (r.ok) {
      const local = DB.getReports();
      const rows = (r.data||[]).map(x => {
        const _key = x._key || (x.nik + '|' + x.report_date);
        const localRow = local.find(y => (y._key || (y.nik+'|'+y.report_date)) === _key);
        const scoreNum = (x.score===0 || x.score==='0' || x.score) ? Number(x.score) :
            computeDailyScore(fmtWIBddmmyyyy(x.report_date), fmtWIBddmmyyyy(x.send_date), coerceHHMM(x.send_time)).score;

        // isSynced: kalau backend sudah isi synced_at ATAU lokal sudah true, pertahankan true
        const syncedFlag = Boolean(x.synced_at && String(x.synced_at).trim() !== '') || (localRow && (localRow.isSynced || localRow.synced_at));

        return { ...x, send_time: coerceHHMM(x.send_time), score: scoreNum, isSynced: syncedFlag, _key };
      });
      DB.setReports(rows);
    }

    const u = await apiGet({action:'listUsers'});
    if (u.ok) DB.setUsers(u.data||[]);

    DB.setMeta({ lastSyncAt: new Date().toISOString() });
  })());

  renderReportsTable(); renderStats(); renderUsersTable(); renderMonitoringTable(); buildFilterOptionsFromParticipants(); renderQueueInfo();

  const fcount = DB.getFailed().length;
  if (fcount) toast(`Ada ${fcount} data gagal sinkron. Buka halaman "Gagal Sync" untuk mencoba ulang.`);
  else toast('Sinkronisasi selesai.');
}

document.addEventListener('DOMContentLoaded', () => {
  const btnSyncPull = $('#btn-sync-pull');
  const btnSyncTab  = $('#btn-sync-tabular');
  if (btnSyncPull) btnSyncPull.addEventListener('click', async (e)=>{ const unlock = lockButton(e.currentTarget); await syncAll(); unlock(); toast('Sinkronisasi selesai.'); });
  if (btnSyncTab)  btnSyncTab .addEventListener('click', async (e)=>{ const unlock = lockButton(e.currentTarget); await syncAll(); unlock(); toast('Sinkronisasi selesai.'); });
});

/* ===================== 11) IMPORT MASTER ===================== */
document.addEventListener('DOMContentLoaded', () => {
  const btnUpload = $('#btn-upload-master');
  const fileInput = $('#file-master');
  if (!btnUpload || !fileInput) return;

  btnUpload.addEventListener('click', ()=> fileInput.click());
  fileInput.addEventListener('change', async function(){
    if (!this.files || !this.files[0]) return;
    const file = this.files[0];
    const unlockBtn = lockButton(btnUpload);
    const unlockInp = lockButton(fileInput);
    try{
      Loader.show('Membaca file master...'); Loader.setProgress(20);
      const rows = await readTableFile(file);
      if (!rows.length){ toast('File kosong atau kolom tidak dikenali.'); return; }
      DB.setParticipants(rows); Loader.setProgress(50);
      const res = await apiPost({ action:'importParticipants', rows });
      Loader.setProgress(85);
      if (!res.ok) toast('Impor lokal berhasil, tapi gagal kirim ke server: ' + (res.error||'unknown'));
      else toast(`Impor selesai: ${res.inserted||0} baru, ${res.updated||0} update, ${res.skipped||0} dilewati`);
      renderStats();
    }catch(err){
      console.error(err); toast('Gagal memproses file: ' + err.message);
    }finally{
      Loader.hide(); this.value=''; unlockBtn(); unlockInp();
    }
  });
});

async function readTableFile(file){
  const buf = await file.arrayBuffer();
  const name = (file.name||'').toLowerCase();
  let rows = [];
  if (name.endsWith('.csv')) {
    const text = new TextDecoder('utf-8').decode(buf);
    rows = parseCSV(text);
  } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
    if (typeof XLSX === 'undefined') throw new Error('Library XLSX belum termuat');
    const wb = XLSX.read(buf, {type:'array'});
    const sh = wb.Sheets[wb.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(sh, {defval:''});
  } else {
    throw new Error('Format tidak didukung. Gunakan .csv atau .xlsx');
  }
  const norm = rows.map(r => normalizeParticipantRow(r));
  return norm.filter(r => r.nik && r.nama);
}
function parseCSV(text){
  const lines = text.replace(/\r/g,'').split('\n').filter(x=>x.trim()!=='');
  if (!lines.length) return [];
  const headers = lines[0].split(',').map(h=>h.trim());
  return lines.slice(1).map(line=>{
    const cols = line.split(',');
    const obj = {}; headers.forEach((h,i)=> obj[h] = (cols[i]||'').trim());
    return obj;
  });
}
function normalizeParticipantRow(r){
  const get = (...keys) => {
    for (const k of keys) {
      const f = Object.keys(r).find(x => x.toLowerCase().trim() === k.toLowerCase());
      if (f) return String(r[f]).trim();
    }
    return '';
  };
  const yes = (v) => {
    const s = String(v||'').toLowerCase();
    return ['1','true','ya','yes','aktif','active','y'].includes(s);
  };
  return {
    nik:     get('nik'),
    nama:    get('nama','name'),
    program: get('program'),
    divisi:  get('divisi','division'),
    unit:    get('unit','estate','unit kerja'),
    region:  get('region','rayon'),
    group:   get('group','kelompok'),
    is_active: yes(get('is_active','aktif','status'))
  };
}

/* ===================== 12) USER MGMT (Modal) ===================== */
document.addEventListener('DOMContentLoaded', () => {
  const btnAdd = $('#btn-add-user');
  const modal = $('#user-modal');
  const closeA = $('#user-modal-close');
  const cancelA = $('#user-modal-cancel');
  const form = $('#user-form');
  if (btnAdd && modal) btnAdd.addEventListener('click', ()=> modal.style.display='block');
  [closeA, cancelA].forEach(btn=> btn && btn.addEventListener('click', ()=> modal.style.display='none'));
  if (form){
    form.addEventListener('submit', async (e)=>{
      e.preventDefault();
      const submitBtn = form.querySelector('button[type="submit"]');
      const unlock = lockButton(submitBtn);
      const payload = {
        action:'upsertUser',
        username: $('#u-username').value.trim(),
        nama:     $('#u-nama').value.trim(),
        role:     $('#u-role').value.trim(),
        status:   $('#u-status').value.trim()
      };
      if (!payload.username || !payload.nama){ unlock(); return toast('Username dan Nama wajib diisi'); }
      try{
        const res = await Loader.withLoading('Menyimpan pengguna...', apiPost(payload));
        if (res && res.ok){
          const list = DB.getUsers();
          const i = list.findIndex(x=>x.username===payload.username);
          const now = new Date().toISOString();
          const obj = { ...payload, updated_at: now, created_at: (i>-1? list[i].created_at : now) };
          if (i>-1) list[i] = obj; else list.push(obj);
          DB.setUsers(list); renderUsersTable();
          toast(res.created ? 'User ditambahkan' : 'User diperbarui');
          modal.style.display='none'; form.reset();
        } else {
          toast('Gagal simpan user: ' + ((res && res.error) || 'unknown'));
        }
      } finally { unlock(); }
    });
  }
});

/* ===================== 13) BOOTSTRAP ===================== */
document.addEventListener('DOMContentLoaded', async function(){
  renderReportsTable(); renderStats(); renderUsersTable(); renderMonitoringTable(); buildFilterOptionsFromParticipants(); renderFailedSyncPage();


  // Chart stub (opsional)
  const t1 = $('#topScoresChart'); const t2 = $('#scoreDistributionChart');
  if (window.Chart) {
    if (t1) new Chart(t1.getContext('2d'), {
      type:'bar', data:{ labels:[], datasets:[{label:'Nilai', data:[]}] },
      options:{ responsive:true, scales:{y:{beginAtZero:true,max:100}} }
    });
    if (t2) new Chart(t2.getContext('2d'), {
      type:'doughnut',
      data:{ labels:['Hijau (90-100)','Kuning (76-89)','Merah (60-75)','Hitam (<60)'], datasets:[{ data:[0,0,0,0] }] },
      options:{ responsive:true, plugins:{ legend:{ position:'bottom' } } }
    });
  }

  // Warm-up awal kalau lokal kosong
  const needWarmup = DB.getParticipants().length===0 || DB.getReports().length===0 || DB.getUsers().length===0;
  if (needWarmup) { try{ await syncAll(); }catch(_){ /* offline ok */ } }
});

/* ===================== MONITORING → CETAK PDF (A4 Landscape) ===================== */
// Pastikan fungsi/variabel berikut sudah ada: FILTER, DB, dd(), fmtWIBddmmyyyy(), applyReportFilters(), getSelectedMonthRange()

// 1) Utility ringkas untuk ringkasan filter aktif
function summarizeActiveFilters(){
  const map = { program:'Program', group:'Group', region:'Region', unit:'Unit', divisi:'Divisi' };
  const parts = [];
  Object.entries(map).forEach(([k, label])=>{
    const v = (FILTER[k]||'').trim();
    if (v) parts.push(`${label}: ${v}`);
  });
  const bulanLabel = FILTER.month
    ? FILTER.month.toLocaleDateString('id-ID',{ month:'long', year:'numeric' })
    : new Date().toLocaleDateString('id-ID',{ month:'long', year:'numeric' });
  return { bulanLabel, text: parts.join(' • ') || 'Semua' };
}

// 2) Konstruksi data monitoring hasil filter (tanpa render ke layar)
function collectMonitoringDataForPrint(){
  const { y, m, nDays } = getSelectedMonthRange();
  const participants = DB.getParticipants().filter(p => String(p.is_active)!=='false');

  // Index laporan bulan terpilih
  const reports = applyReportFilters(DB.getReports());
  const repMap = {};
  for (const r of reports){
    const key = r.nik + '|' + fmtWIBddmmyyyy(r.report_date);
    repMap[key] = Number(r.score);
  }

  // Terapkan filter peserta dimensi (kecuali month sudah ditangani di applyReportFilters)
  const by = { program:'program', group:'group', region:'region', unit:'unit', divisi:'divisi' };
  const pass = (p) => Object.entries(by).every(([fk, field])=>{
    const val = (FILTER[fk]||'').trim();
    return !val || String(p[field]||'').toLowerCase() === val.toLowerCase();
  });
  const list = participants.filter(pass);

  // Bangun row hasil: { no, nik, nama, program, unit, region, days:[...], total }
  const rows = list.map((p, i)=>{
    let total = 0;
    const days = [];
    for (let d=1; d<=nDays; d++){
      const ddm = dd(d) + '/' + dd(m+1) + '/' + y;
      const sc = repMap[p.nik + '|' + ddm];
      const v = (typeof sc === 'number' && !Number.isNaN(sc)) ? sc : '';
      if (typeof v === 'number') total += v;
      days.push(v);
    }
    // Kap total 0..100 (sesuai tampilan)
    const capped = Math.max(0, Math.min(100, Math.round(total)));
    return {
      no: i+1,
      nik: p.nik || '',
      nama: p.nama || '',
      program: p.program || '',
      unit: p.unit || '',
      region: p.region || '',
      days, total: capped
    };
  });

  return { rows, nDays };
}

// === PATCH: Ganti seluruh fungsi buildMonitoringPrintHTML() dengan ini ===
function buildMonitoringPrintHTML(){
  const { bulanLabel, text: filterText } = summarizeActiveFilters();
  const { rows, nDays } = collectMonitoringDataForPrint();

  const style = `
    <style>
      @page { size: A4 landscape; margin: 12mm; }
      * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
      html,body{ font-family: Arial, Helvetica, sans-serif; color:#000; }
      h1{ margin:0 0 6px 0; font-size:16px; }
      .meta{ font-size:11px; margin:0 0 12px 0; color:#333; }
      table{ width:100%; border-collapse:collapse; table-layout:fixed; }
      th, td{ border:1px solid #222; padding:4px 6px; font-size:11px; }
      th{ background:#f2f2f2; }
      .mono{ font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
      .c{ text-align:center; }
      .strong{ font-weight:700; }
      .w-no{ width:22px; }
      .w-nik{ width:90px; }
      .w-nm{ width:220px; }
      .w-prog{ width:90px; } .w-unit{ width:80px; } .w-reg{ width:80px; }
      .w-day{ width:20px; }
      .w-nilai{ width:45px; }
      thead th.sticky { position: sticky; top: 0; }
      td, th { word-break:break-word; }

      /* Palet sel harian (1..n) */
      .p4{ background:#d1fae5; color:#065f46; }   /* hijau muda */
      .p3{ background:#fef3c7; color:#92400e; }   /* kuning muda */
      .p2{ background:#fee2e2; color:#991b1b; }   /* merah muda */
      .p1{ background:#111111; color:#ffffff; }   /* hitam (teks putih) */
      .p0{ background:#ffffff; color:#111111; }   /* kosong/default */
      .sun-empty{ background:#ffe4e6; color:#7a1d2b; } /* Minggu kosong */

      /* Palet kolom NILAI (total) */
      .t-green  { background:#d1fae5; color:#065f46; } /* ≥90 */
      .t-yellow { background:#fef3c7; color:#92400e; } /* 76–89 */
      .t-red    { background:#fee2e2; color:#991b1b; } /* 60–75 */
      .t-black  { background:#111111; color:#ffffff; } /* >0–59 */
      .t-none   { background:#ffffff; color:#111111; } /* 0 atau kosong */
    </style>`;

  // Header kolom
  const headHtml = (() => {
    let days = '';
    for (let d=1; d<=nDays; d++) days += `<th class="sticky w-day">${d}</th>`;
    return `
      <tr>
        <th class="sticky w-no">No</th>
        <th class="sticky w-nik">NIK</th>
        <th class="sticky w-nm">Nama</th>
        <th class="sticky w-prog">Program</th>
        <th class="sticky w-unit">Unit</th>
        <th class="sticky w-reg">Region</th>
        ${days}
        <th class="sticky w-nilai">Nilai</th>
      </tr>`;
  })();

  // Helper: kelas warna sel harian
  function dayCellClass(score, isSunday, empty){
    if (empty && isSunday) return 'sun-empty';
    if (score === 4) return 'p4';
    if (score === 3) return 'p3';
    if (score === 2) return 'p2';
    if (score === 1) return 'p1';
    return 'p0';
  }
  // Helper: kelas warna kolom NILAI total
  function totalClass(total){
    if (typeof total !== 'number' || Number.isNaN(total) || total<=0) return 't-none';
    if (total >= 90) return 't-green';
    if (total >= 76) return 't-yellow';
    if (total >= 60) return 't-red';
    return 't-black';
  }

  // Body
  const monthBase = (FILTER.month ? new Date(FILTER.month) : new Date());
  const y = monthBase.getFullYear();
  const m = monthBase.getMonth(); // 0..11

  const bodyHtml = rows.map(r=>{
    // sel harian berwarna
    let dayTds = '';
    for (let d=1; d<=nDays; d++){
      const v = r.days[d-1]; // '' atau 1/2/3/4
      const empty = (v === '' || v === null || typeof v === 'undefined');
      const sunday = (new Date(y, m, d).getDay() === 0);
      const cls = dayCellClass(Number(v), sunday, empty);
      dayTds += `<td class="c w-day ${cls}">${empty ? '' : v}</td>`;
    }
    const tCls = totalClass(r.total);

    return `<tr>
      <td class="c">${r.no}</td>
      <td class="c mono">${r.nik}</td>
      <td>${r.nama}</td>
      <td class="c">${r.program}</td>
      <td class="c">${r.unit}</td>
      <td class="c">${r.region}</td>
      ${dayTds}
      <td class="c strong w-nilai ${tCls}">${r.total}</td>
    </tr>`;
  }).join('');

  const title = `Monitoring Laporan Harian – ${bulanLabel}`;

  return `
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8">${style}
        <title>${title}</title>
      </head>
      <body>
        <h1>${title}</h1>
        <p class="meta">Filter: ${filterText}</p>
        <table>
          <thead>${headHtml}</thead>
          <tbody>${bodyHtml}</tbody>
        </table>
        <script>
          window.addEventListener('load', function(){ setTimeout(function(){ window.print(); }, 300); });
        </script>
      </body>
    </html>
  `;
}

// 4) Buka jendela preview & cetak
function openMonitoringPrintPreview(){
  const html = buildMonitoringPrintHTML();
  const w = window.open('', '_blank');
  if (!w) { alert('Popup diblokir. Izinkan popup untuk mencetak.'); return; }
  w.document.open(); w.document.write(html); w.document.close(); w.focus();
}

// 5) Sisipkan tombol di Tab Monitoring (sekali saja)
function ensureMonitoringPrintButton(){
  if (document.getElementById('btn-print-monitoring')) return;
  // Cari bar area; jika tidak ada, prepend ke container monitoring
  const host =
    document.querySelector('#monitoring-content .toolbar') ||
    document.querySelector('#monitoring-content') ||
    document.getElementById('monitoring-page');

  if (!host) return;
  const wrap = document.createElement('div');
  wrap.style.cssText = 'display:flex;gap:8px;align-items:center;margin-bottom:10px;';
  const btn = document.createElement('button');
  btn.id = 'btn-print-monitoring';
  btn.className = 'btn btn-primary';
  btn.innerHTML = '<i class="fas fa-file-pdf"></i> Cetak PDF';
  btn.addEventListener('click', openMonitoringPrintPreview);
  wrap.appendChild(btn);
  host.prepend(wrap);
}

// 6) Pastikan tombol muncul saat Tab Monitoring aktif/render
document.addEventListener('DOMContentLoaded', ensureMonitoringPrintButton);
// juga panggil setiap selesai render
const _origRenderMonitoringTable = (typeof renderMonitoringTable === 'function') ? renderMonitoringTable : null;
if (_origRenderMonitoringTable) {
  window.renderMonitoringTable = function(){
    _origRenderMonitoringTable.apply(this, arguments);
    ensureMonitoringPrintButton();
  };
} else {
  // jika fungsi belum terdefinisi pada saat file ini dieksekusi,
  // tetap coba memunculkan tombol setelah sedikit jeda
  setTimeout(ensureMonitoringPrintButton, 800);
}
