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

// ==== Prefetch semua peserta dari server ====
// 1) Coba backend baru dengan parameter limit besar
// 2) Jika backend masih versi lama (tetap 25), fallback fan-out query by prefix (0-9 + a-z)
async function fetchAllParticipants(){
  // --- Strategi 1: coba limit besar (butuh patch GAS di bawah)
  try{
    const res = await apiGet({ action:'listParticipants', q:'', limit: 10000 });
    if (res && res.ok && Array.isArray(res.data) && res.data.length >= 26) {
      await DB.setParticipants(res.data);
      return;
    }
  } catch(_) { /* lanjut ke strategi 2 */ }

  // --- Strategi 2: kompatibel backend lama (cap 25) — fan-out by prefix
  const prefixes = ['0','1','2','3','4','5','6','7','8','9']
    .concat(Array.from({length:26}, (_,i)=>String.fromCharCode(97+i))); // a..z
  const merged = byNikMap(DB.getParticipants()); // gunakan helper yg sdh ada

  // Jalankan bertahap agar aman thd rate-limit
  for (const pfx of prefixes){
    try{
      const r = await apiGet({ action:'listParticipants', q: pfx });
      if (r && r.ok && Array.isArray(r.data)){
        r.data.forEach(p=>{
          if (p && p.nik) merged[p.nik] = { ...(merged[p.nik]||{}), ...p, is_active:true };
        });
      }
    } catch(_){}
  }
  await DB.setParticipants(Object.values(merged));
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

function getDefaultReportDate(){
  const d = new Date();
  d.setDate(d.getDate()-1);
  return d;
}
function getDefaultSendDateTime(){
  return new Date();
}
function resetReportEntryFields(){
  const rpt = $('#report-date');
  const snd = $('#send-datetime');
  const nik = $('#nik');
  if (rpt){
    if (rpt._flatpickr) rpt._flatpickr.setDate(getDefaultReportDate(), true);
    else rpt.value = toDDMMYYYY(getDefaultReportDate());
  }
  if (snd){
    if (snd._flatpickr) snd._flatpickr.setDate(getDefaultSendDateTime(), true);
    else {
      const now = getDefaultSendDateTime();
      snd.value = `${toDDMMYYYY(now)} ${dd(now.getHours())}:${dd(now.getMinutes())}`;
    }
  }
  if (nik){
    requestAnimationFrame(() => { nik.focus(); nik.select?.(); });
    setTimeout(() => { nik.focus(); nik.select?.(); }, 50);
  }
}
function upsertPendingReport(row, oldKey){
  const currentKey = row && row._key ? row._key : (row.nik + '|' + row.report_date);
  if (oldKey && oldKey !== currentKey) DB.removePendingByKey(oldKey);
  DB.removePendingByKey(currentKey);
  DB.addPending({ ...row, _key: currentKey });
}

const SCORE_DONUT_COLORS = {
  hijau: '#22c55e',
  kuning: '#facc15',
  merah: '#ef4444',
  hitam: '#111111'
};

function buildReportPayloadForServer(row, extra = {}){
  const report_date = fmtWIBddmmyyyy(row?.report_date || '');
  const send_date = fmtWIBddmmyyyy(row?.send_date || '');
  const send_time = coerceHHMM(row?.send_time || '');
  const nik = String(row?.nik || '').trim();
  const id = String(row?.id || '').trim();
  const score = (row && row.score !== undefined && row.score !== null && row.score !== '') ? Number(row.score) : '';
  const base = {
    id,
    nik,
    report_date,
    send_date,
    send_time,
    score,
    _key: (row && row._key) ? row._key : (nik && report_date ? (nik + '|' + report_date) : '')
  };
  return { ...base, ...extra };
}

function buildDeletePayloadForServer(row, extra = {}){
  return {
    id: String(row?.id || '').trim(),
    nik: String(row?.nik || '').trim(),
    report_date: fmtWIBddmmyyyy(row?.report_date || ''),
    ...extra
  };
}

/* ===================== [HOL] HOLIDAYS UTIL & STYLE ===================== */

// === NORMALISASI DATA LIBUR → selalu simpan array ISO 'YYYY-MM-DD' (unik) ===
function toISOyyyy_mm_dd_any(v){
  if (!v) return '';
  // already ISO
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // dd/mm/yyyy → iso
  const m = s.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/);
  if (m) {
    const d = +m[1], mo = +m[2]-1, y = +m[3];
    const dt = new Date(y, mo, d);
    if (!isNaN(dt)) return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}-${String(dt.getDate()).padStart(2,'0')}`;
  }
  // coba parse bebas
  const dt = new Date(s);
  if (!isNaN(dt)) {
    return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}-${String(dt.getDate()).padStart(2,'0')}`;
  }
  return '';
}

function normalizeHolidayArray(input){
  const out = new Set();
  (input||[]).forEach(row=>{
    // dukung: '2025-01-01' atau {tanggal:'2025-01-01', ...} atau {date:'...'}
    if (typeof row === 'string') {
      const iso = toISOyyyy_mm_dd_any(row);
      if (iso) out.add(iso);
    } else if (row && typeof row === 'object') {
      const iso = toISOyyyy_mm_dd_any(row.tanggal || row.date || row.tgl || row.holiday || row[0]);
      if (iso) out.add(iso);
    }
  });
  return Array.from(out);
}


// Inject gaya minimal jika class belum ada (buat warna libur & sunday header)
(function injectHolidayStyles(){
  const id = 'holiday-styles';
  if (document.getElementById(id)) return;
  const css = `
    /* Header tanggal Minggu: angka tanggal berwarna pink */
    th.sun-head { color:#ec4899; font-weight:700; }

    /* Header tanggal libur nasional: latar abu-abu */
    th.hol-head { background:#e5e7eb !important; color:#111; }

    /* Sel kolom libur nasional (kosong maupun berisi nilai) */
    td.hol-col { background:#f3f4f6 !important; }

    /* Jika Minggu & kosong: sudah ada .sun-pink di sel lama, kita pastikan warnanya lembut */
    td.sun-pink { background:#ffe4e6 !important; }
  `;
  const style = document.createElement('style');
  style.id = id; style.textContent = css;
  document.head.appendChild(style);
})();

// Simpan libur sebagai ISO (yyyy-mm-dd) di localStorage lalu dipakai per-bulan
function isoOf(y,m,d){ return `${y}-${String(m+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`; }

function buildHolidaySetForMonth(y, m){
  const list = DB.getHolidays(); // array ISO 'YYYY-MM-DD'
  const set = new Set();
  (list||[]).forEach(iso=>{
    const dt = new Date(String(iso).trim() + 'T00:00:00');
    if (!isNaN(dt) && dt.getFullYear()===y && dt.getMonth()===m){
      set.add(dt.getDate());
    }
  });
  return set;
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

/* ===================== A) AGGREGATE & PAGINATION HELPERS ===================== */

// Kategori total (untuk donut & tabel)
function scoreBucket(total){
  if (total >= 90) return { key:'A', label:'Hijau (90–100)', range:[90,100] };
  if (total >= 76) return { key:'B', label:'Kuning (76–89)', range:[76,89] };
  if (total >= 60) return { key:'C', label:'Merah (60–75)', range:[60,75] };
  if (total > 0)   return { key:'D', label:'Hitam (<60)', range:[1,59] };
  return { key:'Z', label:'Tanpa nilai (0)', range:[0,0] };
}

// Agregasi nilai per peserta dalam bulan + filter dimensi yang aktif
function aggregateScoresByParticipant(){
  const { y, m, nDays } = getSelectedMonthRange();
  const participants = DB.getParticipants().filter(p => String(p.is_active)!=='false');

  // Filter peserta by dimensi (program, group, region, unit, divisi)
  const by = { program:'program', group:'group', region:'region', unit:'unit', divisi:'divisi' };
  const pass = (p) => Object.entries(by).every(([fk, field])=>{
    const v = (FILTER[fk]||'').trim();
    return !v || String(p[field]||'').toLowerCase() === v.toLowerCase();
  });
  const baseP = participants.filter(pass);

  // Index laporan bulan-terpilih
  const reports = applyReportFilters(DB.getReports()); // sudah tersaring BULAN + dimensi
  const repMap = {};
  for (const r of reports){
    const ddmmyyyy = fmtWIBddmmyyyy(r.report_date);
    const key = r.nik + '|' + ddmmyyyy;
    repMap[key] = Number(r.score);
  }

  // Hitung total (cap 0..100) per peserta
  return baseP.map(p=>{
    let total = 0;
    for (let d=1; d<=nDays; d++){
      const ddmmyyyy = `${dd(d)}/${dd(m+1)}/${y}`;
      const sc = repMap[p.nik + '|' + ddmmyyyy];
      if (typeof sc === 'number' && !Number.isNaN(sc)) total += sc;
    }
    total = Math.max(0, Math.min(100, Math.round(total)));
    return { ...p, total, bucket: scoreBucket(total) };
  });
}

// ===== Pagination (ellipsis) — reusable
function buildEllipsisPages(cur, total, span=1){
  // ex: for total=20, cur=10 → 1 … 9 10 11 … 20
  if (total <= 7) return Array.from({length:total}, (_,i)=>({type:'page', val:i+1}));
  const out = [];
  const push = (t,v)=> out.push({type:t, val:v});
  push('page',1);
  if (cur > 3+span) push('ellipsis', '...');
  const start = Math.max(2, cur - span);
  const end   = Math.min(total-1, cur + span);
  for (let i=start; i<=end; i++) push('page', i);
  if (cur < total-(2+span)) push('ellipsis', '...');
  push('page', total);
  return out;
}

function renderPagination(host, totalRows, pageSize, state, onGoPage){
  const infoEl = host.querySelector('.pagination-info');
  const ctrlEl = host.querySelector('.pagination-controls .page-numbers');
  const sizeSel= host.querySelector('.pagination-controls .page-size-select');

  const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));
  state.page = Math.max(1, Math.min(state.page || 1, totalPages));

  if (infoEl) infoEl.textContent = `Menampilkan ${totalRows} data`;
  if (sizeSel){
    // Pertahankan pilihan
    sizeSel.value = String(state.size || pageSize);
    sizeSel.onchange = () => {
      state.size = parseInt(sizeSel.value,10) || pageSize;
      state.page = 1;
      onGoPage(1, state.size);
    };
  }
  if (ctrlEl){
    ctrlEl.innerHTML = '';
    const nodes = buildEllipsisPages(state.page, totalPages, 1);
    nodes.forEach(n=>{
      if (n.type==='ellipsis'){
        const span = document.createElement('span');
        span.textContent = '…'; span.style.margin = '0 6px'; span.style.userSelect='none';
        ctrlEl.appendChild(span);
      } else {
        const btn = document.createElement('button');
        btn.textContent = n.val;
        if (n.val === state.page) btn.classList.add('active');
        btn.onclick = () => { state.page = n.val; onGoPage(state.page, state.size||pageSize); };
        ctrlEl.appendChild(btn);
      }
    });
  }
  return { totalPages, current: state.page };
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

    // 5) Update lokal (optimistic) + masukkan ke antrean sync
    const updatedRow = {
      ...data,
      id: data.id || id || '',
      report_date,
      send_date,
      send_time,
      score: finalScore,
      _key: newKey,
      isSynced: false,
      synced_at: ''
    };
    DB.deleteReportByKey(key);
    DB.upsertReportLocal(updatedRow);
    upsertPendingReport(updatedRow, key);
    renderReportsTable();
    renderStats();
    renderQueueInfo();
    if (typeof renderMonitoringTable === 'function') renderMonitoringTable();

    toast('Perubahan disimpan lokal. Masuk antrean sinkron.');
    unlock();
    doClose();
  });
}

// ⬇️ ADD HERE: delete flow
async function deleteReportFlow({ key, id }) {
  const row = DB.getReportByKey(key);
  if (!row) return toast('Data tidak ditemukan.');
  if (!confirm(`Hapus laporan NIK ${row.nik} tanggal ${fmtWIBddmmyyyy(row.report_date)} ?`)) return;

  // Hapus lokal dulu (optimistic)
  DB.deleteReportByKey(key);
  DB.removePendingByKey(key);
  DB.addPending({
    _key: `DELETE|${row.id || key}`,
    _op: 'delete',
    id: row.id || id || '',
    nik: row.nik,
    report_date: row.report_date,
    send_date: row.send_date || '',
    send_time: row.send_time || '',
    score: row.score || ''
  });
  renderReportsTable();
  renderStats();
  renderQueueInfo();
  if (typeof renderMonitoringTable === 'function') renderMonitoringTable();
  toast('Laporan dihapus lokal. Masuk antrean sinkron.');
}


/* ===================== 4) DB (IndexedDB + In-Memory Cache) ===================== */
const DB = (() => {
  // --- fallback kalau IndexedDB/Dexie tidak tersedia ---
  const hasIDB = typeof indexedDB !== 'undefined' && typeof Dexie !== 'undefined';

  // kunci lama (untuk migrasi & fallback)
  const KEYS = { P:'kmp.participants', R:'kmp.reports', RP:'kmp.reports_pending', RF:'kmp.reports_failed', U:'kmp.users', M:'kmp.meta', H:'kmp.holidays' };
  const lsRead  = (k,def)=>{ try{ return JSON.parse(localStorage.getItem(k)||JSON.stringify(def)); }catch(_){ return def; } };
  const lsWrite = (k,v)=> localStorage.setItem(k, JSON.stringify(v));

  // --- in-memory cache agar API getter tetap sinkron ---
  const cache = {
    participants: [],
    reports: [],
    reports_pending: [],
    reports_failed: [],
    users: [],
    meta: { lastSyncAt:null, version:1 },
    holidays: []
  };

  // ---------- Fallback PENUH: gunakan localStorage persis seperti dulu ----------
  if (!hasIDB) {
    console.warn('[DB] IndexedDB/Dexie tidak tersedia. Fallback ke localStorage.');
    // inisialisasi cache dari LS agar konsisten
    cache.participants   = lsRead(KEYS.P, []);
    cache.reports        = lsRead(KEYS.R, []);
    cache.reports_pending= lsRead(KEYS.RP, []);
    cache.reports_failed = lsRead(KEYS.RF, []);
    cache.users          = lsRead(KEYS.U, []);
    cache.meta           = lsRead(KEYS.M, { lastSyncAt:null, version:1 });
    cache.holidays       = lsRead(KEYS.H, []);

    // adapter kompatibel (API sama)
    function syncLS(){
      lsWrite(KEYS.P, cache.participants);
      lsWrite(KEYS.R, cache.reports);
      lsWrite(KEYS.RP, cache.reports_pending);
      lsWrite(KEYS.RF, cache.reports_failed);
      lsWrite(KEYS.U, cache.users);
      lsWrite(KEYS.M, cache.meta);
      lsWrite(KEYS.H, cache.holidays);
    }

    // expose
    const api = {
      // lifecycle
      async ready(){}, // no-op
      async migrateFromLocalIfNeeded(){}, // no-op pada fallback

      // collections (getters sinkron, setters update cache + LS)
      getParticipants(){ return cache.participants; },
      setParticipants(rows){ cache.participants = Array.isArray(rows)? rows: []; syncLS(); },

      getReports(){ return cache.reports; },
      setReports(rows){ cache.reports = Array.isArray(rows)? rows: []; syncLS(); },
      upsertReportLocal(obj){
        const key = obj._key || (obj.nik + '|' + obj.report_date);
        const i = cache.reports.findIndex(r => (r._key || (r.nik+'|'+r.report_date)) === key);
        if (i>-1) cache.reports[i] = {...cache.reports[i], ...obj}; else cache.reports.push({...obj});
        syncLS();
      },

      getPending(){ return cache.reports_pending; },
      setPending(rows){ cache.reports_pending = Array.isArray(rows)? rows: []; syncLS(); },
      addPending(obj){ cache.reports_pending.push(obj); syncLS(); },
      removePendingByKey(k){ cache.reports_pending = cache.reports_pending.filter(r => r._key !== k); syncLS(); },

      getFailed(){ return cache.reports_failed; },
      setFailed(rows){ cache.reports_failed = Array.isArray(rows)? rows: []; syncLS(); },
      addFailed(obj){ cache.reports_failed.push(obj); syncLS(); },
      removeFailedByKey(k){ cache.reports_failed = cache.reports_failed.filter(r => r._key !== k); syncLS(); },

      getUsers(){ return cache.users; },
      setUsers(rows){ cache.users = Array.isArray(rows)? rows: []; syncLS(); },

      getMeta(){ return cache.meta; },
      setMeta(meta){ cache.meta = { ...cache.meta, ...(meta||{}) }; syncLS(); },

      getHolidays(){ return cache.holidays; },
      setHolidays(arr){ cache.holidays = Array.isArray(arr)? arr: []; syncLS(); },

      getReportByKey(k){ return cache.reports.find(r => (r._key || (r.nik+'|'+r.report_date)) === k); },
      deleteReportByKey(k){ cache.reports = cache.reports.filter(r => (r._key || (r.nik+'|'+r.report_date)) !== k); syncLS(); },
      updateReportByKey(k, patch){
        const i = cache.reports.findIndex(r => (r._key || (r.nik+'|'+r.report_date)) === k);
        if (i>-1){ cache.reports[i] = { ...cache.reports[i], ...(patch||{}) }; syncLS(); }
      },

      // clear helpers
      clearReports(){ cache.reports=[]; syncLS(); },
      clearParticipants(){ cache.participants=[]; syncLS(); },
      clearQueues(){ cache.reports_pending=[]; cache.reports_failed=[]; syncLS(); },
      clearUsers(){ cache.users=[]; syncLS(); },
      clearMeta(){ cache.meta={ lastSyncAt:null, version:1 }; syncLS(); },
      clearAll(){ cache.participants=[]; cache.reports=[]; cache.reports_pending=[]; cache.reports_failed=[]; cache.users=[]; cache.meta={ lastSyncAt:null, version:1 }; cache.holidays=[]; syncLS(); }
    };
    return api;
  }

  // ---------- Mode IndexedDB (dengan Dexie) ----------
  const db = new Dexie('kmp-db');
  db.version(1).stores({
    participants: 'nik, program, unit, region, divisi, group',
    reports:      '_key, nik, report_date, send_date, send_time, score, isSynced, synced_at',
    reports_pending: '_key',
    reports_failed:  '_key',
    users:        'username',
    meta:         'key',
    holidays:     'iso'
  });

  let _isReady = false;

  async function loadAllIntoCache(){
    const [P,R,RP,RF,U,M,H] = await Promise.all([
      db.participants.toArray(),
      db.reports.toArray(),
      db.reports_pending.toArray(),
      db.reports_failed.toArray(),
      db.users.toArray(),
      db.meta.get('app_meta'),
      db.holidays.toArray()
    ]);
    cache.participants   = P || [];
    cache.reports        = R || [];
    cache.reports_pending= RP|| [];
    cache.reports_failed = RF|| [];
    cache.users          = U || [];
    cache.meta           = M || { key:'app_meta', lastSyncAt:null, version:1, migrated:false };
    cache.holidays       = (H||[]).map(x=>x.iso);
  }

  async function persistMeta(patch){
    const merged = { ...(cache.meta||{}), ...(patch||{}), key:'app_meta' };
    cache.meta = merged;
    await db.meta.put(merged);
  }

  const api = {
    async ready(){
      if (_isReady) return;
      await loadAllIntoCache();
      _isReady = true;
    },

    // sekali waktu: copy dari localStorage -> IndexedDB
    async migrateFromLocalIfNeeded(){
      await api.ready();
      if (cache.meta && cache.meta.migrated) return; // sudah migrasi

      // baca semua sumber lama
      const P  = lsRead(KEYS.P, []);
      const R  = lsRead(KEYS.R, []);
      const RP = lsRead(KEYS.RP, []);
      const RF = lsRead(KEYS.RF, []);
      const U  = lsRead(KEYS.U, []);
      const M  = lsRead(KEYS.M, { lastSyncAt:null, version:1 });
      const H  = lsRead(KEYS.H, []);

      // tulis ke IDB (replace)
      await db.transaction('rw', db.tables, async ()=>{
        await db.participants.clear();     if (P.length)  await db.participants.bulkPut(P);
        await db.reports.clear();          if (R.length)  await db.reports.bulkPut(R.map(x=>({ ...x, _key: x._key || (x.nik+'|'+x.report_date) })));
        await db.reports_pending.clear();  if (RP.length) await db.reports_pending.bulkPut(RP.map(x=>({ ...x, _key: x._key || (x.nik+'|'+x.report_date) })));
        await db.reports_failed.clear();   if (RF.length) await db.reports_failed.bulkPut(RF.map(x=>({ ...x, _key: x._key || (x.nik+'|'+x.report_date) })));
        await db.users.clear();            if (U.length)  await db.users.bulkPut(U);
        await db.holidays.clear();         if (H.length)  await db.holidays.bulkPut((H||[]).map(iso=>({ iso })));
        await db.meta.put({ key:'app_meta', ...M, migrated:true, migratedAt:new Date().toISOString() });
      });

      await loadAllIntoCache();
      // opsional: bersihkan LS lama (aman)
      try{
        [KEYS.P,KEYS.R,KEYS.RP,KEYS.RF,KEYS.U,KEYS.M,KEYS.H].forEach(k=> localStorage.removeItem(k));
      }catch(_){}
    },

    // ------- API kompatibel (getter sinkron — baca dari cache) -------
    getParticipants(){ return cache.participants; },
    async setParticipants(rows){
      cache.participants = Array.isArray(rows)? rows: [];
      await db.participants.clear();
      if (cache.participants.length) await db.participants.bulkPut(cache.participants);
    },

    getReports(){ return cache.reports; },
    async setReports(rows){
      cache.reports = Array.isArray(rows)? rows: [];
      // pastikan setiap baris punya _key
      cache.reports.forEach(r => r._key = r._key || (r.nik + '|' + r.report_date));
      await db.reports.clear();
      if (cache.reports.length) await db.reports.bulkPut(cache.reports);
    },
    async upsertReportLocal(obj){
      const key = obj._key || (obj.nik + '|' + obj.report_date);
      const i = cache.reports.findIndex(r => (r._key || (r.nik+'|'+r.report_date)) === key);
      if (i>-1) cache.reports[i] = { ...cache.reports[i], ...obj, _key:key };
      else cache.reports.push({ ...obj, _key:key });
      await db.reports.put(cache.reports.find(r => r._key===key));
    },

    getPending(){ return cache.reports_pending; },
    async setPending(rows){
      cache.reports_pending = Array.isArray(rows)? rows: [];
      cache.reports_pending.forEach(r => r._key = r._key || (r.nik + '|' + r.report_date));
      await db.reports_pending.clear();
      if (cache.reports_pending.length) await db.reports_pending.bulkPut(cache.reports_pending);
    },
    async addPending(obj){
      const key = obj._key || (obj.nik + '|' + obj.report_date);
      const row = { ...obj, _key:key };
      cache.reports_pending.push(row);
      await db.reports_pending.put(row);
    },
    async removePendingByKey(k){
      cache.reports_pending = cache.reports_pending.filter(r => r._key !== k);
      await db.reports_pending.delete(k);
    },

    getFailed(){ return cache.reports_failed; },
    async setFailed(rows){
      cache.reports_failed = Array.isArray(rows)? rows: [];
      cache.reports_failed.forEach(r => r._key = r._key || (r.nik + '|' + r.report_date));
      await db.reports_failed.clear();
      if (cache.reports_failed.length) await db.reports_failed.bulkPut(cache.reports_failed);
    },
    async addFailed(obj){
      const key = obj._key || (obj.nik + '|' + obj.report_date);
      const row = { ...obj, _key:key };
      cache.reports_failed.push(row);
      await db.reports_failed.put(row);
    },
    async removeFailedByKey(k){
      cache.reports_failed = cache.reports_failed.filter(r => r._key !== k);
      await db.reports_failed.delete(k);
    },

    getUsers(){ return cache.users; },
    async setUsers(rows){
      cache.users = Array.isArray(rows)? rows: [];
      await db.users.clear();
      if (cache.users.length) await db.users.bulkPut(cache.users);
    },

    getMeta(){ return cache.meta; },
    async setMeta(meta){ await persistMeta(meta); },

    getHolidays(){ return cache.holidays; },
    async setHolidays(arr){
      const list = Array.isArray(arr)? arr: [];
      cache.holidays = list;
      await db.holidays.clear();
      if (list.length) await db.holidays.bulkPut(list.map(iso => ({ iso })));
    },

    getReportByKey(k){ return cache.reports.find(r => (r._key || (r.nik+'|'+r.report_date)) === k); },
    async deleteReportByKey(k){
      cache.reports = cache.reports.filter(r => (r._key || (r.nik+'|'+r.report_date)) !== k);
      await db.reports.delete(k);
    },
    async updateReportByKey(k, patch){
      const i = cache.reports.findIndex(r => (r._key || (r.nik+'|'+r.report_date)) === k);
      if (i>-1){
        cache.reports[i] = { ...cache.reports[i], ...(patch||{}) };
        await db.reports.put(cache.reports[i]);
      }
    },

    // clear helpers
    async clearReports(){ cache.reports=[]; await db.reports.clear(); },
    async clearParticipants(){ cache.participants=[]; await db.participants.clear(); },
    async clearQueues(){ cache.reports_pending=[]; cache.reports_failed=[]; await db.reports_pending.clear(); await db.reports_failed.clear(); },
    async clearUsers(){ cache.users=[]; await db.users.clear(); },
    async clearMeta(){ cache.meta={ key:'app_meta', lastSyncAt:null, version:1, migrated:true }; await db.meta.put(cache.meta); },
    async clearAll(){
      await db.transaction('rw', db.tables, async ()=>{
        for (const t of db.tables) await t.clear();
        await db.meta.put({ key:'app_meta', lastSyncAt:null, version:1, migrated:true });
      });
      await loadAllIntoCache();
    }
  };

  return api;
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

/* ===================== B) DASHBOARD RENDERERS ===================== */

const DASH = {
  sortCol: 'total', // total | nik | nama | program | unit | region | divisi
  sortDir: 'desc',
  page: 1,
  size: 20,
  search: ''
};

// Ambil elemen tabel Dashboard
function getDashboardTableRefs(){
  const root = document.querySelector('#dashboard-page');
  return {
    root,
    tbody: root?.querySelector('.card-body .data-table tbody'),
    hostPag: root?.querySelector('.card-body .pagination'),
    sizeSel: root?.querySelector('.card-body .page-size-select'),
    searchInput: root?.querySelector('.card-header input[type="text"]')
  };
}

// Render bar chart Top-10
function renderTop10Chart(agg){
  const ctx = document.getElementById('topScoresChart')?.getContext('2d');
  if (!ctx || !window.Chart) return;

  const top10 = [...agg].sort((a,b)=> b.total - a.total).slice(0,10);
  const labels = top10.map(x => (x.nama || x.nik || '').toString());
  const data   = top10.map(x => x.total);

  // destroy existing if any
  if (renderTop10Chart._chart) { renderTop10Chart._chart.destroy(); }
  renderTop10Chart._chart = new Chart(ctx, {
    type:'bar',
    data:{ labels, datasets:[{ label:'Nilai Total', data }] },
    options:{
      responsive:true,
      scales:{ y:{ beginAtZero:true, max:100 } },
      plugins:{ tooltip:{ callbacks:{ label:(ctx)=> ` ${ctx.raw}` } } }
    }
  });
}

// Render donut Distribusi (klik segmen → buka modal daftar peserta dalam range)
function renderDistributionChart(agg){
  const ctx = document.getElementById('scoreDistributionChart')?.getContext('2d');
  if (!ctx || !window.Chart) return;

  const buckets = [
    {key:'A', label:'Hijau (90–100)', where:(t)=>t>=90},
    {key:'B', label:'Kuning (76–89)', where:(t)=>t>=76 && t<=89},
    {key:'C', label:'Merah (60–75)', where:(t)=>t>=60 && t<=75},
    {key:'D', label:'Hitam (<60)',   where:(t)=>t>0 && t<60}
  ];
  const counts = buckets.map(b => agg.filter(p => b.where(p.total)).length);

  if (renderDistributionChart._chart) renderDistributionChart._chart.destroy();
  const chartColors = [SCORE_DONUT_COLORS.hijau, SCORE_DONUT_COLORS.kuning, SCORE_DONUT_COLORS.merah, SCORE_DONUT_COLORS.hitam];
  const chart = new Chart(ctx, {
    type:'doughnut',
    data:{
      labels:buckets.map(b=>b.label),
      datasets:[{
        data:counts,
        backgroundColor: chartColors,
        borderColor: '#ffffff',
        borderWidth: 2,
        hoverOffset: 8
      }]
    },
    options:{
      responsive:true,
      plugins:{ legend:{ position:'bottom' } },
      onClick: (evt, elements)=>{
        if (!elements.length) return;
        const idx = elements[0].index;
        const b = buckets[idx];
        const rows = agg
          .filter(p => b.where(p.total))
          .sort((a,b)=> b.total - a.total);
        openParticipantListModal({
          title: `Peserta – ${b.label}`,
          rows,
          showTotal:true
        });
      }
    }
  });
  renderDistributionChart._chart = chart;
}

// Modal daftar peserta (klik “Total Peserta” atau dari donut)
function openParticipantListModal({ title='Daftar Peserta', rows=[], showTotal=false }={}){
  const backdrop = document.createElement('div');
  backdrop.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:2200;display:flex;align-items:center;justify-content:center;';
  const card = document.createElement('div');
  card.style.cssText = 'width:min(900px,96vw);max-height:90vh;background:#fff;border-radius:10px;box-shadow:0 10px 30px rgba(0,0,0,.3);overflow:hidden;display:flex;flex-direction:column;';
  card.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:space-between;padding:10px 14px;border-bottom:1px solid #e5e7eb;">
      <h3 style="margin:0;color:var(--primary);font-size:1.05rem;">${title}</h3>
      <div style="display:flex;gap:8px;align-items:center;">
        <select id="modal-program" class="form-control" style="min-width:160px"></select>
        <button class="btn btn-outline btn-sm" id="modal-close">Tutup</button>
      </div>
    </div>
    <div style="padding:12px;overflow:auto;">
      <div class="table-container">
        <table class="data-table">
          <thead>
            <tr>
              <th>No</th><th>NIK</th><th>Nama</th><th>Program</th><th>Divisi</th><th>Unit</th><th>Region</th>${showTotal?'<th>Total</th>':''}
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>
      <div class="pagination" style="margin-top:8px;">
        <div class="pagination-info">Menampilkan 0 data</div>
        <div class="pagination-controls">
          <select class="form-control page-size-select" style="width:auto;">
            <option>20</option><option>40</option><option>80</option><option>100</option><option>500</option>
          </select>
          <div class="page-numbers"><button class="active">1</button></div>
        </div>
      </div>
    </div>
  `;
  backdrop.appendChild(card); document.body.appendChild(backdrop);
  const close = ()=> backdrop.remove();
  card.querySelector('#modal-close').onclick = close;

  // Isi pilihan program
  const progSel = card.querySelector('#modal-program');
  const programs = uniq(DB.getParticipants().filter(p=>String(p.is_active)!=='false').map(p=>p.program));
  progSel.innerHTML = `<option value="">Semua Program</option>` + programs.map(p=>`<option>${p}</option>`).join('');

  const state = { page:1, size:20 };
  const all = rows.length ? rows : DB.getParticipants().filter(p => String(p.is_active)!=='false');

  function applyFilter(data){
    const v = progSel.value || '';
    if (!v) return data;
    return data.filter(p => String(p.program||'').toLowerCase() === v.toLowerCase());
  }

  function renderModalTable(){
    const tbody = card.querySelector('tbody');
    const hostPag = card.querySelector('.pagination');
    const data = applyFilter(all);
    // paging
    const go = (pg, size)=>{ state.page=pg; state.size=size||state.size; renderModalTable(); };
    renderPagination(hostPag, data.length, state.size, state, go);
    const start = (state.page-1)*state.size;
    const pageRows = data.slice(start, start+state.size);

    tbody.innerHTML = '';
    pageRows.forEach((p,i)=>{
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${start+i+1}</td>
        <td>${p.nik||''}</td>
        <td>${p.nama||''}</td>
        <td>${p.program||''}</td>
        <td>${p.divisi||''}</td>
        <td>${p.unit||''}</td>
        <td>${p.region||''}</td>
        ${showTotal?`<td>${p.total||0}</td>`:''}
      `;
      tbody.appendChild(tr);
    });
  }

  progSel.onchange = ()=>{ state.page=1; renderModalTable(); };
  renderModalTable();
}

// Tabel Leaderboard di Dashboard (urut dari total tertinggi + sorting header + search + paging elipsis)
function renderDashboardTable(agg){
  const { tbody, hostPag, searchInput } = getDashboardTableRefs();
  if (!tbody || !hostPag) return;

  // filter pencarian
  const q = (DASH.search||'').toLowerCase();
  let data = agg.filter(x=>{
    const s = `${x.nik||''} ${x.nama||''} ${x.program||''} ${x.divisi||''} ${x.unit||''} ${x.region||''}`.toLowerCase();
    return s.includes(q);
  });

  // sorting
  const cmp = {
    total:  (a,b)=> a.total - b.total,
    nik:    (a,b)=> String(a.nik||'').localeCompare(String(b.nik||'')),
    nama:   (a,b)=> String(a.nama||'').localeCompare(String(b.nama||'')),
    program:(a,b)=> String(a.program||'').localeCompare(String(b.program||'')),
    divisi: (a,b)=> String(a.divisi||'').localeCompare(String(b.divisi||'')),
    unit:   (a,b)=> String(a.unit||'').localeCompare(String(b.unit||'')),
    region: (a,b)=> String(a.region||'').localeCompare(String(b.region||'')),
  };
  const sorter = cmp[DASH.sortCol] || cmp.total;
  data.sort(sorter); if (DASH.sortDir==='desc') data.reverse();

  // paging
  const go = (pg, size)=>{ DASH.page=pg; DASH.size=size||DASH.size; renderDashboardTable(agg); };
  renderPagination(hostPag, data.length, DASH.size, DASH, go);
  const start = (DASH.page-1)*DASH.size;
  const pageRows = data.slice(start, start+DASH.size);

  // render rows
  tbody.innerHTML = '';
  pageRows.forEach((p,i)=>{
    const cat = p.bucket.key;
    const catLabel = p.bucket.label;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${start+i+1}</td>
      <td>${p.nik||''}</td>
      <td>${p.nama||''}</td>
      <td>${p.unit||''}</td>
      <td>${p.region||''}</td>
      <td>${totalScoreBadge(p.total)}</td>
      <td>${catLabel}</td>
    `;
    tbody.appendChild(tr);
  });

  // search binding (sekali)
  if (searchInput && !searchInput._bound){
    searchInput._bound = true;
    let t=null;
    searchInput.addEventListener('input', e=>{
      clearTimeout(t);
      t=setTimeout(()=>{ DASH.search = e.target.value||''; DASH.page=1; renderDashboardTable(agg); }, 180);
    });
  }

  // header sorting (sekali)
  const mapIdx = {1:'nik',2:'nama',3:'unit',4:'region',5:'total',6:'total'}; // 5=nilai badge; 6=kategori (tetap sort 'total' saat klik Kategori)
  const ths = document.querySelectorAll('#dashboard-page .data-table thead th');
  if (ths && !ths._bound){
    ths._bound = true;
    ths.forEach((th, idx)=>{
      const key = mapIdx[idx];
      if (!key) return;
      th.style.cursor='pointer';
      th.title='Klik untuk sort';
      th.addEventListener('click', ()=>{
        if (DASH.sortCol === key) DASH.sortDir = (DASH.sortDir==='asc'?'desc':'asc');
        else { DASH.sortCol = key; DASH.sortDir='asc'; }
        renderDashboardTable(agg);
      });
    });
  }
}

// Render keseluruhan Dashboard (dipanggil saat buka tab Dashboard atau setelah sync/filter)
function renderDashboard(){
  // Pastikan kartu Total Peserta bisa di-klik
  const cardTotal = document.getElementById('stat-total-peserta') || document.querySelector('#dashboard-page .stat-card.stat-green');
  if (cardTotal && !cardTotal._bound){
    cardTotal._bound = true;
    cardTotal.style.cursor = 'pointer';
    cardTotal.title = 'Klik untuk melihat daftar peserta';
    cardTotal.addEventListener('click', ()=>{
      const agg = aggregateScoresByParticipant();
      openParticipantListModal({ title:'Daftar Peserta (semua program)', rows:agg, showTotal:true });
    });
  }

  const agg = aggregateScoresByParticipant();
  renderTop10Chart(agg);
  renderDistributionChart(agg);
  renderDashboardTable(agg);
}

function renderReportsTable(){
  const tbody = document.querySelector('#tabular-content table.data-table tbody'); if (!tbody) return;
  const hostPag = document.querySelector('#tabular-content .pagination');

  // ambil & filter (per laporan)
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

  // === NEW: paging elipsis untuk Tabular ===
  renderReportsTable._state = renderReportsTable._state || { page:1, size:20 };
  const state = renderReportsTable._state;
  const go = (pg,size)=>{ state.page=pg; state.size=size||state.size; renderReportsTable(); };
  const { } = renderPagination(hostPag, rows.length, state.size, state, go);
  const start = (state.page-1)*(state.size);
  const pageRows = rows.slice(start, start+state.size);

  // render
  tbody.innerHTML = '';
  pageRows.forEach((r, idx)=>{
    const key = r._key || (r.nik + '|' + r.report_date);
    const P = pMap[r.nik] || {};
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${start+idx+1}</td>
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
        <button class="btn btn-danger btn-sm" data-act="delete" data-id="${r.id||''}" data-key="${key}"><i class="fas fa-trash"></i></button>
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

  // === NEW: banner ringkasan import terakhir (jika ada)
  const container = tbody.closest('.card-body');
  if (container && !container.querySelector('#import-errors-banner')) {
    const banner = document.createElement('div');
    banner.id = 'import-errors-banner';
    banner.style.cssText = 'margin-bottom:12px;padding:10px;border:1px solid #f5c2c7;background:#f8d7da;color:#842029;border-radius:6px;display:none;';
    container.prepend(banner);
  }
  const meta = DB.getMeta();
  const banner = document.getElementById('import-errors-banner');
  if (banner) {
    const errs = meta.lastImportErrors || [];
    if (errs.length) {
      // tampilkan 3 contoh alasan pertama
      const examples = errs.slice(0, 3).map(e => `• [${e.reason||'error'}] ${e.message||''}`).join('<br>');
      banner.innerHTML = `<strong>${errs.length} baris gagal saat import terakhir.</strong><br>${examples}${errs.length>3?'<br>…':''} <br><small>Periksa tabel di bawah, lalu klik <i class="fas fa-redo"></i> untuk coba sinkron ulang atau <i class="fas fa-times"></i> untuk hapus dari daftar.</small>`;
      banner.style.display = 'block';
    } else {
      banner.style.display = 'none';
    }
  }

  // === render tabel gagal (kode lama tetap) ===
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

  // action (tetap)
  tbody.onclick = async (e)=>{
    const btn = e.target.closest('button');
    if (!btn) return;
    if (btn.dataset.retry){
      const key = btn.dataset.retry;
      const row = DB.getFailed().find(x=>x._key===key);
      if (!row) return;
      try{
        const res = await Loader.withLoading('Mencoba sync ulang...', apiPost(
          row && row._op === 'delete'
            ? {
                action:'deleteReport',
                id:row.id || '',
                nik:row.nik,
                report_date:row.report_date
              }
            : {
                action:'upsertReport',
                id:row.id || '',
                nik:row.nik, report_date:row.report_date, send_date:row.send_date, send_time:row.send_time, score:row.score||'',
                markSynced: true
              }
        ));
        if (res && res.ok){
          DB.removeFailedByKey(key);
          if (!(row && row._op === 'delete')) {
            DB.upsertReportLocal({ ...row, id: res.id || row.id, isSynced:true, synced_at:new Date().toISOString() });
          }
          renderFailedSyncPage(); renderReportsTable(); renderStats(); renderQueueInfo();
          toast('Berhasil disinkron.');
        } else {
          // tampilkan alasan dari backend bila ada
          const reason = (res && (res.message || res.error)) || 'unknown';
          DB.updateReportByKey(key, {}); // no-op: placeholder bila ingin menandai
          toast('Masih gagal: ' + reason);
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
  const holSet = buildHolidaySetForMonth(y, m); // <- NEW

  const fixedCols = ['No','NIK','Nama','Program','Divisi','Unit','Region','Group'];
  let ths = fixedCols.map(h=>`<th>${h}</th>`).join('');

  for (let d=1; d<=nDays; d++){
    const isSun = isSunday(y, m, d);
    const isHol = holSet.has(d);
    const cls = [
      isSun ? 'sun-head' : '',
      isHol ? 'hol-head' : ''
    ].filter(Boolean).join(' ');
    ths += `<th class="${cls}">${d}</th>`;
  }
  ths += `<th>Nilai</th>`;
  thead.innerHTML = `<tr>${ths}</tr>`;
}


function renderMonitoringTable(){
  renderMonitoringHead(); // bangun thead sesuai bulan
  const tbody = document.querySelector('#monitoring-content table.data-table tbody');
  if (!tbody) return;

  const { y, m, nDays } = getSelectedMonthRange();
  const holSet = buildHolidaySetForMonth(y, m);
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
    const sdd = dd(day);
    const smm = dd(m+1);
    const yyyy = y;

    const ddmmyyyy = `${sdd}/${smm}/${yyyy}`;
    const key = p.nik + '|' + ddmmyyyy;
    const sc = repMap[key];

    const sunday  = isSunday(y, m, day);
    const isHol   = holSet.has(day);

    const sunClass = sunday && (sc===undefined || sc==='' || Number.isNaN(sc)) ? 'sun-pink' : '';
    const holClass = isHol ? 'hol-col' : '';

    if (typeof sc === 'number' && !Number.isNaN(sc)){
        total += sc;
        let txtClass = 'txt-none';
        if (sc===4) txtClass='txt-green';
        else if (sc===3) txtClass='txt-yellow';
        else if (sc===2) txtClass='txt-red';
        else if (sc===1) txtClass='txt-black';

        tds.push(`
        <td class="cell-score ${holClass} ${sunClass}">
            ${dailyScoreBadge(sc)}
            <span class="score-text ${txtClass}">${sc}</span>
        </td>
        `);
    } else {
        tds.push(`<td class="cell-score ${holClass} ${sunClass}"></td>`);
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

      // Reset scroll agar halaman baru tampil dari atas
      const contentEl = document.querySelector('.content');
      if (contentEl) contentEl.scrollTop = 0;
      try { window.scrollTo({ top: 0, left: 0, behavior: 'instant' }); } catch(e){ window.scrollTo(0,0); }

      if (pageId==='settings-page') renderUsersTable();
      if (pageId==='report-page') renderReportsTable();
      if (pageId==='dashboard-page'){ renderStats(); renderDashboard(); }
      if (pageId==='monitoring-page') renderMonitoringTable();
      if (pageId==='failedsync-page') renderFailedSyncPage();

      // === Embedded Download (konv/index.html) ===
      // Beberapa browser/server menunda load iframe saat elemen tersembunyi.
      // Jadi kita set src saat page "Download" benar-benar dibuka.
      if (pageId==='extractor-page'){
        const frame = document.getElementById('extractor-frame');
        const loading = document.getElementById('extractor-loading');
        if (frame){
          // pasang handler sekali
          if (!frame.__bindDone){
            frame.addEventListener('load', () => {
              if (loading) loading.style.display = 'none';
            });
            frame.addEventListener('error', () => {
              if (loading){
                loading.style.display = 'flex';
                loading.innerHTML = `<div style="padding:16px;line-height:1.4;">
                  <div style="font-weight:800;color:#b91c1c;">Gagal memuat halaman konv/index.html</div>
                  <div style="margin-top:6px;color:var(--gray);">
                    Pastikan folder <b>konv</b> berada satu level dengan halaman utama dan dapat diakses lewat browser.
                  </div>
                </div>`;
              }
            });
            frame.__bindDone = true;
          }

          // trigger load (sekali / atau paksa reload jika sebelumnya belum sukses)
          const src = frame.getAttribute('data-src') || frame.getAttribute('src') || 'konv/index.html';
          if (!frame.getAttribute('src')){
            if (loading) loading.style.display = 'flex';
            frame.setAttribute('src', src + (src.includes('?') ? '&' : '?') + 'embed=1&_ts=' + Date.now());
          }
        }
      }
      
      attachExportButtons();
      
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
      attachExportButtons();
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
    flatpickr.localize(flatpickr.l10ns.id);
    flatpickr(rpt, {
      dateFormat: 'd/m/Y',
      defaultDate: getDefaultReportDate(),
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
      defaultDate: getDefaultSendDateTime(),
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
    upsertPendingReport(localObj);
    bumpPickCount(nik);
    renderReportsTable(); renderStats(); renderQueueInfo();
    toast('Tersimpan lokal. Masuk antrean sinkron.');

    // reset form + kembalikan tanggal/jam default terbaru
    this.reset();
    if (infoBox) infoBox.style.display = 'none';
    resetReportEntryFields();

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
        let res;
        if (r && r._op === 'delete') {
          res = await apiPost({
            action:'deleteReport',
            ...buildDeletePayloadForServer(r)
          });
          if (res && res.ok){
            DB.removePendingByKey(r._key);
          } else {
            DB.removePendingByKey(r._key);
            DB.addFailed({ ...r, reason: (res && (res.error || res.message)) || 'unknown' });
          }
        } else {
          // kirim dengan markSynced:true agar backend memberi synced_at
          res = await apiPost({ 
            action:'upsertReport',
            ...buildReportPayloadForServer(r, { markSynced: true })
          });
          if (res && res.ok){
            // tandai lokal sukses
            DB.removePendingByKey(r._key);
            DB.upsertReportLocal({ ...r, id: res.id || r.id, isSynced:true, synced_at:new Date().toISOString() });
          } else {
            // pindah ke antrean gagal
            DB.removePendingByKey(r._key);
            DB.addFailed({ ...r, reason: (res && (res.error || res.message)) || 'unknown' });
          }
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
    // [HOL] Tarik master libur nasional (ISO yyyy-mm-dd)
    try{
        const h = await apiGet({ action:'listHolidays' });
        if (h && h.ok && Array.isArray(h.data)) {
            DB.setHolidays( normalizeHolidayArray(h.data) ); // <<— penting
        }
        }catch(_){}

    await fetchAllParticipants();

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

/* ===================== 11a) IMPORT DATA AKTUAL (.xlsx/.xls/.csv) + TEMPLATE ===================== */
document.addEventListener('DOMContentLoaded', () => {
  const btnUploadActual = document.getElementById('btn-upload-actual');
  const inputActual     = document.getElementById('file-actual');
  const btnTplActual    = document.getElementById('btn-download-actual-template');

  if (btnUploadActual && inputActual) {
    btnUploadActual.addEventListener('click', () => inputActual.click());
    inputActual.addEventListener('change', async function(){
      if (!this.files || !this.files[0]) return;
      const file   = this.files[0];
      const unlock = lockButton(btnUploadActual);
      try{
        Loader.show('Membaca file data aktual...'); Loader.setProgress(15);
        const rows = await readActualFile(file); // normalisasi → {nik,report_date,send_date,send_time,score,...}
        if (!rows.length){ toast('File kosong atau kolom tidak dikenali.'); return; }

        // Upsert lokal satu per satu
        let done = 0;
        for (const r of rows){
          DB.upsertReportLocal(r);
          done++;
          if (done % 20 === 0) Loader.setProgress(15 + Math.min(60, Math.round(done/rows.length*60)));
        }
        Loader.setProgress(80);

        // Coba kirim ke server (jika backend mendukung). Kalau gagal, tetap aman di lokal.
        try{
          const res = await apiPost({ action:'importReports', rows });
            Loader.setProgress(95);

            if (res && res.ok){
            // Tandai tersinkron utk baris yang lolos
            const stamped = rows.map(x => ({ ...x, isSynced:true, synced_at:new Date().toISOString() }));
            stamped.forEach(x => DB.upsertReportLocal(x));

            // === NEW: proses error baris (jika ada) → masukkan ke Gagal Sync
            const errs = Array.isArray(res.errors) ? res.errors : [];
            let addedFailed = 0;
            if (errs.length){
                errs.forEach(er => {
                const idx = typeof er.index === 'number' ? er.index : -1;
                const src = rows[idx] || {}; // baris asal (bisa tidak ada)
                const nik = src.nik || er.nik || '';
                const report_date = fmtWIBddmmyyyy(src.report_date || er.report_date || '');
                const send_date   = fmtWIBddmmyyyy(src.send_date || '');
                const send_time   = coerceHHMM(src.send_time || '');
                const score       = (src.score===0 || src.score==='0' || src.score) ? Number(src.score) : '';

                const _key = nik && report_date ? (nik + '|' + report_date) : (src._key || localId());

                DB.addFailed({
                    _key, id: src.id || '',
                    nik, report_date, send_date, send_time, score,
                    reason: (er.message || er.reason || 'Unknown import error'),
                });
                addedFailed++;
                });

                // simpan ringkasan ke meta untuk ditampilkan sebagai banner di halaman Gagal Sync
                const meta = DB.getMeta();
                DB.setMeta({
                ...meta,
                lastImportAt: new Date().toISOString(),
                lastImportErrors: errs.slice(0, 200) // batasi agar ringan
                });
            }

            renderFailedSyncPage(); renderReportsTable(); renderStats(); renderMonitoringTable(); renderQueueInfo();

            const msg = `Impor selesai. ${res.inserted||0} baru, ${res.updated||0} update, ${res.skipped||0} dilewati` +
                        (errs.length ? `, ${errs.length} baris bermasalah → lihat "Gagal Sync".` : '');
            toast(msg);
            } else {
            toast(`Impor selesai lokal. Kirim ke server gagal/skip: ${(res && res.error) || 'unknown'}.`);
            }
        } catch(_){
          toast('Impor selesai lokal. Offline/Backend tidak mendukung import massal.');
        }

        renderReportsTable(); renderStats(); renderMonitoringTable(); renderQueueInfo();
      }catch(err){
        console.error(err);
        toast('Gagal memproses file: ' + err.message);
      }finally{
        Loader.hide(); this.value=''; unlock();
      }
    });
  }

  if (btnTplActual) {
    btnTplActual.addEventListener('click', ()=> {
      try { downloadActualTemplateXLSX(); }
      catch(err){ console.error(err); toast('Gagal membuat template: ' + err.message); }
    });
  }
});

// Baca file aktual → array objek normal (siap upsert)
async function readActualFile(file){
  const buf  = await file.arrayBuffer();
  const name = (file.name||'').toLowerCase();
  let rowsRaw = [];
  if (name.endsWith('.csv')) {
    const text = new TextDecoder('utf-8').decode(buf);
    rowsRaw = parseCSV(text);
  } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
    if (typeof XLSX === 'undefined') throw new Error('Library XLSX belum termuat');

   // Baca date sebagai Date, dan cari sheet "actual_reports" jika ada
   const wb  = XLSX.read(buf, { type:'array', cellDates:true });
   const targetName = wb.SheetNames.includes('actual_reports') ? 'actual_reports' : wb.SheetNames[0];
   const sh  = wb.Sheets[targetName];
   // raw:false biar waktu "07:00:00" jadi string, bukan angka kasar; defval:'' biar tidak undefined
   rowsRaw   = XLSX.utils.sheet_to_json(sh, { defval:'', raw:false });
  } else {
    throw new Error('Format tidak didukung. Gunakan .csv atau .xlsx');
  }

  const norm = rowsRaw.map(r => normalizeActualRow(r))
                      .filter(x => x && x.nik && x.report_date);
  return norm;
}

// Konversi nilai tanggal campuran (Date object / "YYYY-MM-DD" / serial Excel) → "dd/mm/yyyy"
function toDDMMYYYY_Flex(v){
  if (v instanceof Date && !isNaN(v)) return toDDMMYYYY(v);
  const s = String(v||'').trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  const asDate = new Date(s);
  if (!isNaN(asDate)) return toDDMMYYYY(asDate);
  const n = Number(s);
  if (!Number.isNaN(n) && n>0 && n<600000){ // kisaran aman serial Excel
    // Excel epoch 1899-12-30
    const base = Date.UTC(1899, 11, 30);
    const ms = base + Math.round(n*24*60*60*1000);
    return toDDMMYYYY(new Date(ms));
  }
  return fmtWIBddmmyyyy(s);
}

// Normalisasi 1 baris data aktual → konsisten dgn skema app
function normalizeActualRow(r){
  const get = (...keys) => {
    for (const k of keys){
      const f = Object.keys(r).find(x => x.toLowerCase().trim() === k.toLowerCase());
      if (f) return r[f]; // biarkan tipe aslinya (bisa Date/number/string)
    }
    return '';
  };

 const nikRaw      = get('nik','NIK');
 const nik         = (typeof nikRaw === 'number') ? String(Math.trunc(nikRaw)) : String(nikRaw||'').replace(/\.0$/,'');
 const reportDate  = toDDMMYYYY_Flex(get('report_date','tanggal laporan','tgl laporan','tgl_laporan','tanggal'));
 const sendDate    = toDDMMYYYY_Flex(get('send_date','tanggal kirim','tgl kirim','tgl_kirim'));
 const sendTime    = coerceHHMM( get('send_time','jam kirim','jam_kirim','waktu','time') );
  let scoreStr      = get('score','nilai');

  let score = (scoreStr!=='' && scoreStr!=null) ? Number(scoreStr) : computeDailyScore(reportDate, sendDate, sendTime).score;
  const _key = nik + '|' + reportDate;

  return {
    id: localId(),
    nik,
    report_date: reportDate,
    send_date: sendDate,
    send_time: sendTime,
    score: (Number.isFinite(score) ? score : ''),
    isSynced: false,
    synced_at: '',
    _key,
    imported_at: new Date().toISOString()
  };
}


// Buat & unduh template .xlsx contoh import data aktual
function downloadActualTemplateXLSX(){
  if (typeof XLSX === 'undefined') throw new Error('Library XLSX belum termuat');

  // Sheet README
  const readmeAOA = [
    ['Template Import Data Aktual – Laporan Harian'],
    ['Isi kolom sesuai header pada sheet "actual_reports".'],
    ['Format tanggal: dd/mm/yyyy, format jam: HH:MM (24 jam).'],
    ['Kolom "score" opsional. Jika kosong, sistem akan menghitung otomatis:'],
    ['- 4: Dikirim ≤ H+1 08:00 | 3: ≤ H+1 12:00 | 2: ≤ H+1 23:59 | 1: Lewat dari itu'],
    ['Header wajib: nik, report_date, send_date, send_time, score(optional)'],
  ];
  const shReadme = XLSX.utils.aoa_to_sheet(readmeAOA);

  // Sheet contoh data
  const today  = new Date();
  const ddmmmy = toDDMMYYYY(today);
  const sample = [
    { nik:'6501012345', report_date:ddmmmy, send_date:ddmmmy, send_time:'07:45', score:'' },
    { nik:'6501098765', report_date:ddmmmy, send_date:ddmmmy, send_time:'13:10', score:'2' } // contoh isi manual
  ];
  const shData = XLSX.utils.json_to_sheet(sample, {header:['nik','report_date','send_date','send_time','score']});

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, shReadme, 'README');
  XLSX.utils.book_append_sheet(wb, shData,   'actual_reports');
  XLSX.writeFile(wb, 'template_import_aktual.xlsx');
}

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

/* ===================== 12a) SETTINGS – CLEAR LOCAL STORAGE ===================== */
document.addEventListener('DOMContentLoaded', () => {
  const bRep  = document.getElementById('btn-clear-reports');
  const bPar  = document.getElementById('btn-clear-participants');
  const bQue  = document.getElementById('btn-clear-queues');
  const bAll  = document.getElementById('btn-clear-all');

  if (bRep) bRep.addEventListener('click', ()=>{
    if (confirm('Hapus SEMUA laporan lokal?')) {
      DB.clearReports(); renderReportsTable(); renderStats(); renderMonitoringTable(); renderQueueInfo();
      toast('Laporan lokal dihapus.');
    }
  });
  if (bPar) bPar.addEventListener('click', ()=>{
    if (confirm('Hapus master peserta lokal?')) {
      DB.clearParticipants(); buildFilterOptionsFromParticipants(); renderStats();
      toast('Master peserta lokal dihapus.');
    }
  });
  if (bQue) bQue.addEventListener('click', ()=>{
    if (confirm('Hapus antrean Pending & Gagal sync?')) {
      DB.clearQueues(); renderFailedSyncPage(); renderQueueInfo();
      toast('Antrean pending/gagal dihapus.');
    }
  });
  if (bAll) bAll.addEventListener('click', ()=>{
    if (confirm('HAPUS SEMUA DATA LOKAL (laporan, peserta, pengguna, meta, antrean)?')) {
      DB.clearAll(); renderReportsTable(); renderFailedSyncPage(); buildFilterOptionsFromParticipants(); renderStats(); renderMonitoringTable(); renderUsersTable(); renderQueueInfo();
      toast('Semua data lokal dihapus. Lakukan Sinkronisasi untuk menarik ulang dari server.');
    }
  });
});

// ========== EXPORT EXCEL UTIL ==========
function tableToAOA(tableEl){
  const rows = [];
  const trsHead = tableEl.querySelectorAll('thead tr');
  const trsBody = tableEl.querySelectorAll('tbody tr');

  trsHead.forEach(tr=>{
    const row = Array.from(tr.cells).map(td=> td.innerText.trim());
    rows.push(row);
  });
  trsBody.forEach(tr=>{
    const row = Array.from(tr.cells).map(td=> td.innerText.trim());
    rows.push(row);
  });
  return rows;
}

function exportAOAtoXLSX(aoa, fileName='export.xlsx', sheetName='Sheet1'){
  if (typeof XLSX === 'undefined') { alert('Library XLSX belum termuat.'); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, fileName);
}

function exportTableSelectorToXLSX(selector, fileName, sheetName){
  const tbl = document.querySelector(selector);
  if (!tbl){ alert('Tabel tidak ditemukan.'); return; }
  const aoa = tableToAOA(tbl);
  exportAOAtoXLSX(aoa, fileName, sheetName);
}

// Tambahan khusus Monitoring: ekspor dari data yang sedang difilter (lebih cepat & rapi)
function exportMonitoringCurrentView(){
  const host = document.querySelector('#monitoring-content table.data-table');
  if (host){
    const labelBulan = (FILTER.month || new Date())
      .toLocaleDateString('id-ID',{ month:'long', year:'numeric' })
      .replace(/\s+/g,'_');
    exportTableSelectorToXLSX(
      '#monitoring-content table.data-table',
      `monitoring_${labelBulan}.xlsx`,
      'Monitoring'
    );
  } else {
    alert('Tabel Monitoring tidak ditemukan.');
  }
}

function ensureExportBtn({ hostSelector, btnId, tableSelector, label='Export Excel', fileName='export.xlsx', sheetName='Sheet1'}){
  const host = document.querySelector(hostSelector);
  if (!host || document.getElementById(btnId)) return;

  const wrap = document.createElement('div');
  wrap.style.cssText = 'display:flex;gap:8px;align-items:center;margin-bottom:8px;flex-wrap:wrap;';
  const btn = document.createElement('button');
  btn.id = btnId;
  btn.className = 'btn btn-outline';
  btn.innerHTML = `<i class="fas fa-file-excel"></i> ${label}`;
  btn.onclick = () => exportTableSelectorToXLSX(tableSelector, fileName, sheetName);
  wrap.appendChild(btn);

  // sisipkan di paling atas container
  host.prepend(wrap);
}


/* ===================== 13) BOOTSTRAP ===================== */
document.addEventListener('DOMContentLoaded', async function(){
  // 0) Pastikan DB siap dan lakukan migrasi dari localStorage kalau perlu
  if (DB && typeof DB.ready === 'function') await DB.ready();
  if (DB && typeof DB.migrateFromLocalIfNeeded === 'function') await DB.migrateFromLocalIfNeeded();

  // 1) Render awal (semua getter DB sudah aman dipakai)
  renderReportsTable(); renderStats(); renderUsersTable(); renderMonitoringTable(); buildFilterOptionsFromParticipants(); renderFailedSyncPage(); renderDashboard();

  /*// 2) Chart stub (opsional) — biarkan seperti semula
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
  }*/

  // 3) Warm-up awal jika kosong → tarik dari server
  const needWarmup = DB.getParticipants().length===0 || DB.getReports().length===0 || DB.getUsers().length===0;
  if (needWarmup) { try{ await syncAll(); }catch(_){ /* offline ok */ } }
});

// === Tambah tombol Export di berbagai halaman/section ===
function attachExportButtons(){
  // Dashboard: leaderboard (tabel utama)
  ensureExportBtn({
    hostSelector:'#dashboard-page .card-body',
    btnId:'btn-export-dashboard',
    tableSelector:'#dashboard-page .data-table',
    label:'Export Excel',
    fileName:'dashboard_leaderboard.xlsx',
    sheetName:'Leaderboard'
  });

  // Tabular (Laporan)
  ensureExportBtn({
    hostSelector:'#report-page .card-body',
    btnId:'btn-export-tabular',
    tableSelector:'#tabular-content table.data-table',
    label:'Export Excel',
    fileName:'laporan_tabular.xlsx',
    sheetName:'Laporan'
  });

  // Monitoring (pakai tombol khusus di samping Cetak PDF)
  (function ensureMonitoringExportButton(){
    const host =
      document.querySelector('#monitoring-content .toolbar') ||
      document.querySelector('#monitoring-content') ||
      document.getElementById('monitoring-page');
    if (!host || document.getElementById('btn-export-monitoring')) return;

    const btn = document.createElement('button');
    btn.id = 'btn-export-monitoring';
    btn.className = 'btn btn-outline';
    btn.style.marginLeft = '8px';
    btn.innerHTML = '<i class="fas fa-file-excel"></i> Export Excel';
    btn.addEventListener('click', exportMonitoringCurrentView);

    // upayakan berdampingan dengan tombol Cetak PDF jika ada
    const bar = document.getElementById('btn-print-monitoring')?.parentElement || host;
    bar.appendChild(btn);
  })();

  // Users (Settings)
  ensureExportBtn({
    hostSelector:'#settings-page',
    btnId:'btn-export-users',
    tableSelector:'#settings-page table.data-table',
    label:'Export Excel',
    fileName:'users.xlsx',
    sheetName:'Users'
  });

  // Gagal Sync
  ensureExportBtn({
    hostSelector:'#failedsync-page .card-body',
    btnId:'btn-export-failed',
    tableSelector:'#failedsync-page table.data-table',
    label:'Export Excel',
    fileName:'gagal_sync.xlsx',
    sheetName:'Failed'
  });
}

// panggil saat halaman/tab berubah
document.addEventListener('DOMContentLoaded', attachExportButtons);

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

/* ===================== 18) QUICK INPUT PAGE (CLICK ONLY) ===================== */
const QUICK_INPUT = (() => {
  const MEM_KEY = 'kmp.quickInput.memory';
  const AUTH_KEY = 'kmp.quickInput.auth';
  const SESSION_COUNT_KEY = 'kmp.quickInput.sessionCounts';
  const PASSWORD = 'userL1234';

  const state = {
    selectedProgram: '',
    selectedNik: '',
    reportDate: toDDMMYYYY(getDefaultReportDate()),
    sendDate: toDDMMYYYY(getDefaultSendDateTime()),
    sendHour: dd(getDefaultSendDateTime().getHours()),
    sendMinute: dd(getDefaultSendDateTime().getMinutes()),
    viewReportMonth: new Date(getDefaultReportDate().getFullYear(), getDefaultReportDate().getMonth(), 1),
    viewSendMonth: new Date(getDefaultSendDateTime().getFullYear(), getDefaultSendDateTime().getMonth(), 1)
  };

  function loadMemory(){
    try{
      const saved = JSON.parse(localStorage.getItem(MEM_KEY) || '{}');
      if (saved && typeof saved === 'object') Object.assign(state, saved);
    }catch(_){ }
  }
  function saveMemory(){
    try {
      localStorage.setItem(MEM_KEY, JSON.stringify({
        selectedProgram: state.selectedProgram || '',
        selectedNik: state.selectedNik || '',
        reportDate: state.reportDate || '',
        sendDate: state.sendDate || '',
        sendHour: state.sendHour || '00',
        sendMinute: state.sendMinute || '00',
        viewReportMonth: state.viewReportMonth ? state.viewReportMonth.toISOString() : '',
        viewSendMonth: state.viewSendMonth ? state.viewSendMonth.toISOString() : ''
      }));
    } catch(_){ }
  }
  function restoreViewMonths(){
    if (typeof state.viewReportMonth === 'string' && state.viewReportMonth) state.viewReportMonth = new Date(state.viewReportMonth);
    if (!(state.viewReportMonth instanceof Date) || isNaN(state.viewReportMonth)) state.viewReportMonth = new Date(getDefaultReportDate().getFullYear(), getDefaultReportDate().getMonth(), 1);
    if (typeof state.viewSendMonth === 'string' && state.viewSendMonth) state.viewSendMonth = new Date(state.viewSendMonth);
    if (!(state.viewSendMonth instanceof Date) || isNaN(state.viewSendMonth)) state.viewSendMonth = new Date(getDefaultSendDateTime().getFullYear(), getDefaultSendDateTime().getMonth(), 1);
  }


  function loadSessionCounts(){
    try{
      const raw = JSON.parse(sessionStorage.getItem(SESSION_COUNT_KEY) || '{}');
      return raw && typeof raw === 'object' ? raw : {};
    }catch(_){ return {}; }
  }
  function saveSessionCounts(map){
    try{ sessionStorage.setItem(SESSION_COUNT_KEY, JSON.stringify(map || {})); }catch(_){ }
  }
  function getPickCount(nik){
    const map = loadSessionCounts();
    return Math.max(0, Number(map[String(nik||'')] || 0) || 0);
  }
  function bumpPickCount(nik){
    const key = String(nik || '').trim();
    if (!key) return 0;
    const map = loadSessionCounts();
    map[key] = Math.max(0, Number(map[key] || 0) || 0) + 1;
    saveSessionCounts(map);
    return map[key];
  }
  function pickLevelClass(count){
    const n = Math.max(0, Number(count) || 0);
    if (n <= 0) return '';
    return ' picked picked-' + String(Math.min(n, 5));
  }

  function isAuthorized(){ return localStorage.getItem(AUTH_KEY) === '1'; }
  function setAuthorized(v){ localStorage.setItem(AUTH_KEY, v ? '1' : '0'); }

  function participants(){ return DB.getParticipants().filter(p => String(p.is_active) !== 'false'); }
  function getPrograms(){ return uniq(participants().map(p => String(p.program||'').trim()).filter(Boolean)).sort((a,b)=>a.localeCompare(b)); }
  function getSelectedParticipant(){ return participants().find(p => String(p.nik||'') === String(state.selectedNik||'')) || null; }
  function participantsByProgram(){
    if (!state.selectedProgram) return [];
    return participants()
      .filter(p => String(p.program||'').trim() === state.selectedProgram)
      .sort((a,b)=> String(a.nama||'').localeCompare(String(b.nama||'')) || String(a.nik||'').localeCompare(String(b.nik||'')));
  }

  function openPage(){
    document.querySelectorAll('.sidebar-menu a').forEach(i => i.classList.remove('active'));
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    const page = document.getElementById('quickinput-page');
    if (page) page.classList.add('active');
    const contentEl = document.querySelector('.content');
    if (contentEl) contentEl.scrollTop = 0;
    renderAll();
  }
  function goNormalInput(){
    const link = document.querySelector('.sidebar-menu a[data-page="input"]');
    if (link) link.click();
  }

  function showAccessModal(){
    if (isAuthorized()) return openPage();
    const modal = document.getElementById('quick-access-modal');
    const input = document.getElementById('quick-access-password');
    if (modal) modal.style.display = 'flex';
    if (input){ input.value=''; setTimeout(()=>input.focus(), 20); }
  }
  function hideAccessModal(){
    const modal = document.getElementById('quick-access-modal');
    const input = document.getElementById('quick-access-password');
    if (modal) modal.style.display = 'none';
    if (input) input.value = '';
  }
  function submitPassword(){
    const input = document.getElementById('quick-access-password');
    const val = (input?.value || '').trim();
    if (val !== PASSWORD) return toast('Password halaman input cepat tidak sesuai.');
    setAuthorized(true);
    hideAccessModal();
    openPage();
  }

  function renderPrograms(){
    const wrap = document.getElementById('quick-programs');
    if (!wrap) return;
    const list = getPrograms();
    wrap.innerHTML = list.length ? '' : '<div class="quick-muted">Belum ada master program.</div>';
    list.forEach(name => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'quick-chip' + (state.selectedProgram === name ? ' active' : '');
      btn.textContent = name;
      btn.addEventListener('click', ()=>{
        state.selectedProgram = name;
        const rows = participantsByProgram();
        if (!rows.some(p => String(p.nik||'') === String(state.selectedNik||''))) state.selectedNik = rows[0]?.nik || '';
        saveMemory();
        renderPrograms();
        renderParticipants();
        renderSummary();
      });
      wrap.appendChild(btn);
    });
  }

  function renderParticipants(){
    const wrap = document.getElementById('quick-participants');
    const meta = document.getElementById('quick-participants-meta');
    if (!wrap || !meta) return;
    const rows = participantsByProgram();
    meta.textContent = state.selectedProgram
      ? `${rows.length} peserta pada program ${state.selectedProgram}`
      : 'Pilih program terlebih dahulu.';
    wrap.innerHTML = '';
    if (!state.selectedProgram){
      wrap.innerHTML = '<div class="quick-muted">Belum ada program yang dipilih.</div>';
      return;
    }
    if (!rows.length){
      wrap.innerHTML = '<div class="quick-muted">Tidak ada peserta aktif pada program ini.</div>';
      return;
    }
    rows.forEach(p => {
      const card = document.createElement('button');
      card.type = 'button';
      const pickCount = getPickCount(p.nik);
      card.className = 'quick-participant-card' + pickLevelClass(pickCount) + (String(state.selectedNik||'') === String(p.nik||'') ? ' active' : '');
      card.setAttribute('data-picked-count', String(pickCount));
      card.innerHTML = `<div class="q-name">${p.nama || '-'}</div><div class="q-meta">${p.nik || '-'} • ${p.unit || '-'}${p.divisi ? ' • ' + p.divisi : ''}</div>`;
      card.addEventListener('click', ()=>{
        state.selectedNik = p.nik || '';
        saveMemory();
        renderParticipants();
        renderSummary();
      });
      wrap.appendChild(card);
    });
  }

  function parseDDMMYYYY(value){
    const m = String(value||'').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (!m) return null;
    return new Date(+m[3], +m[2]-1, +m[1]);
  }
  function monthTitle(d){ return d.toLocaleDateString('id-ID',{month:'long', year:'numeric'}); }
  function sameDate(a,b){ return a && b && a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }

  function renderCalendar(targetId, valueKey, viewKey){
    const root = document.getElementById(targetId);
    if (!root) return;
    const view = state[viewKey];
    const selected = parseDDMMYYYY(state[valueKey]);
    const today = new Date();
    const first = new Date(view.getFullYear(), view.getMonth(), 1);
    const start = new Date(first);
    const offset = (first.getDay() + 6) % 7; // Monday first
    start.setDate(first.getDate() - offset);

    root.innerHTML = `
      <div class="quick-cal-head">
        <div class="quick-cal-title">${monthTitle(view)}</div>
        <div class="quick-cal-nav">
          <button type="button" data-nav="prev"><i class="fas fa-chevron-left"></i></button>
          <button type="button" data-nav="next"><i class="fas fa-chevron-right"></i></button>
        </div>
      </div>
      <div class="quick-cal-weekdays"><div>Sen</div><div>Sel</div><div>Rab</div><div>Kam</div><div>Jum</div><div>Sab</div><div>Min</div></div>
      <div class="quick-cal-days"></div>
    `;
    root.querySelector('[data-nav="prev"]').addEventListener('click', ()=>{
      state[viewKey] = new Date(view.getFullYear(), view.getMonth()-1, 1);
      saveMemory(); renderCalendar(targetId, valueKey, viewKey);
    });
    root.querySelector('[data-nav="next"]').addEventListener('click', ()=>{
      state[viewKey] = new Date(view.getFullYear(), view.getMonth()+1, 1);
      saveMemory(); renderCalendar(targetId, valueKey, viewKey);
    });
    const daysWrap = root.querySelector('.quick-cal-days');
    for(let i=0;i<42;i++){
      const d = new Date(start);
      d.setDate(start.getDate()+i);
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'quick-day';
      if (d.getMonth() !== view.getMonth()) btn.classList.add('muted');
      if (sameDate(d, today)) btn.classList.add('today');
      if (selected && sameDate(d, selected)) btn.classList.add('active');
      btn.textContent = String(d.getDate());
      btn.addEventListener('click', ()=>{
        state[valueKey] = toDDMMYYYY(d);
        state[viewKey] = new Date(d.getFullYear(), d.getMonth(), 1);
        saveMemory();
        renderCalendar(targetId, valueKey, viewKey);
        renderSummary();
      });
      daysWrap.appendChild(btn);
    }
  }

  function renderTimeGrid(){
    const hourWrap = document.getElementById('quick-hour-grid');
    const minuteWrap = document.getElementById('quick-minute-grid');
    if (hourWrap){
      hourWrap.innerHTML = '';
      for(let i=0;i<24;i++){
        const v = dd(i);
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = state.sendHour === v ? 'active' : '';
        btn.textContent = v;
        btn.addEventListener('click', ()=>{ state.sendHour = v; saveMemory(); renderTimeGrid(); renderSummary(); });
        hourWrap.appendChild(btn);
      }
    }
    if (minuteWrap){
      minuteWrap.innerHTML = '';
      for(let i=0;i<60;i++){
        const v = dd(i);
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = state.sendMinute === v ? 'active' : '';
        btn.textContent = v;
        btn.addEventListener('click', ()=>{ state.sendMinute = v; saveMemory(); renderTimeGrid(); renderSummary(); });
        minuteWrap.appendChild(btn);
      }
    }
  }

  function renderSummary(){
    const box = document.getElementById('quick-summary');
    if (!box) return;
    const p = getSelectedParticipant();
    const sendTime = `${state.sendHour}:${state.sendMinute}`;
    box.innerHTML = `
      <div class="quick-summary-item"><div class="quick-summary-label">Program</div><div class="quick-summary-value">${state.selectedProgram || '-'}</div></div>
      <div class="quick-summary-item"><div class="quick-summary-label">Peserta</div><div class="quick-summary-value">${p ? (p.nama || '-') : '-'}</div></div>
      <div class="quick-summary-item"><div class="quick-summary-label">NIK / Unit</div><div class="quick-summary-value">${p ? ((p.nik||'-') + ' / ' + (p.unit||'-')) : '-'}</div></div>
      <div class="quick-summary-item"><div class="quick-summary-label">Tanggal Laporan</div><div class="quick-summary-value">${state.reportDate || '-'}</div></div>
      <div class="quick-summary-item"><div class="quick-summary-label">Tanggal Kirim</div><div class="quick-summary-value">${state.sendDate || '-'}</div></div>
      <div class="quick-summary-item"><div class="quick-summary-label">Jam Kirim</div><div class="quick-summary-value">${sendTime}</div></div>
    `;
  }

  function renderAll(){
    restoreViewMonths();
    const progs = getPrograms();
    if (!state.selectedProgram || !progs.includes(state.selectedProgram)) state.selectedProgram = progs[0] || '';
    const rows = participantsByProgram();
    if (!rows.some(p => String(p.nik||'') === String(state.selectedNik||''))) state.selectedNik = rows[0]?.nik || '';
    saveMemory();
    renderPrograms();
    renderParticipants();
    renderCalendar('quick-report-calendar', 'reportDate', 'viewReportMonth');
    renderCalendar('quick-send-calendar', 'sendDate', 'viewSendMonth');
    renderTimeGrid();
    renderSummary();
  }

  function resetMemory(){
    const rpt = getDefaultReportDate();
    const snd = getDefaultSendDateTime();
    state.selectedProgram = '';
    state.selectedNik = '';
    state.reportDate = toDDMMYYYY(rpt);
    state.sendDate = toDDMMYYYY(snd);
    state.sendHour = dd(snd.getHours());
    state.sendMinute = dd(snd.getMinutes());
    state.viewReportMonth = new Date(rpt.getFullYear(), rpt.getMonth(), 1);
    state.viewSendMonth = new Date(snd.getFullYear(), snd.getMonth(), 1);
    saveMemory();
    renderAll();
    toast('Pilihan input cepat direset.');
  }

  function saveReport(){
    const p = getSelectedParticipant();
    if (!state.selectedProgram) return toast('Pilih program terlebih dahulu.');
    if (!p || !p.nik) return toast('Pilih peserta terlebih dahulu.');
    if (!state.reportDate || !state.sendDate) return toast('Tanggal laporan dan tanggal kirim wajib dipilih.');
    const send_time = `${state.sendHour}:${state.sendMinute}`;
    const report_date = fmtWIBddmmyyyy(state.reportDate);
    const send_date = fmtWIBddmmyyyy(state.sendDate);
    const nik = String(p.nik || '').trim();
    const { score } = computeDailyScore(report_date, send_date, send_time);
    const _key = nik + '|' + report_date;
    const localObj = {
      id: localId(),
      nik, report_date, send_date, send_time, score,
      isSynced: false, synced_at: '', _key,
      created_locally_at: new Date().toISOString()
    };
    DB.upsertReportLocal(localObj);
    upsertPendingReport(localObj);
    bumpPickCount(nik);
    renderReportsTable(); renderStats(); renderQueueInfo();
    if (typeof renderMonitoringTable === 'function') renderMonitoringTable();
    renderParticipants();
    renderSummary();
    toast('Tersimpan lokal dari Input Cepat. Masuk antrean sinkron.');
    saveMemory();
  }

  function bind(){
    loadMemory();
    restoreViewMonths();
    document.getElementById('quick-access-btn')?.addEventListener('click', showAccessModal);
    document.getElementById('quick-access-cancel')?.addEventListener('click', hideAccessModal);
    document.getElementById('quick-access-submit')?.addEventListener('click', submitPassword);
    document.getElementById('quick-access-password')?.addEventListener('keydown', (e)=>{ if (e.key === 'Enter') submitPassword(); });
    document.getElementById('quick-access-modal')?.addEventListener('click', (e)=>{ if (e.target?.id === 'quick-access-modal') hideAccessModal(); });
    document.getElementById('btn-quick-back-input')?.addEventListener('click', goNormalInput);
    document.getElementById('btn-quick-reset-memory')?.addEventListener('click', resetMemory);
    document.getElementById('btn-quick-save')?.addEventListener('click', saveReport);
  }

  return { bind, renderAll, showAccessModal };
})();

document.addEventListener('DOMContentLoaded', () => {
  QUICK_INPUT.bind();
});
