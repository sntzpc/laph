/** ==========================================================
 *  Laporan Harian - Data Extractor (GAS Backend)
 *  - Source A: Spreadsheet "pusingan" (Pusingan Panen)
 *  - Source B: Spreadsheet "rkh" + master_asisten (mapping NIK/Nama)
 *  - Output supports JSON and JSONP (for CORS-free access from localhost/GitHub Pages)
 *
 *  Actions:
 *   - getPusingan&startDate=YYYY-MM-DD&endDate=YYYY-MM-DD
 *   - getRKH&startDate=YYYY-MM-DD&endDate=YYYY-MM-DD
 *
 *  Optional:
 *   - callback=fnName  (JSONP)
 *  ========================================================== */

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = String(params.action || '').trim();

  // If called without action, serve a simple landing (optional)
  if (!action) {
    return ContentService.createTextOutput(
      'Laporan Harian - Data Extractor backend OK. Use ?action=getPusingan or ?action=getRKH'
    ).setMimeType(ContentService.MimeType.TEXT);
  }

  const result = handleAPIRequest_(params);
  return output_(result, params.callback);
}

function doPost(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const result = handleAPIRequest_(params);
  return output_(result, params.callback);
}

function handleAPIRequest_(params) {
  const action = String(params.action || '').trim();
  const startDate = params.startDate;
  const endDate = params.endDate;

  try {
    if (action === 'getPusingan') return getDataPusingan_(startDate, endDate);
    if (action === 'getRKH') return getDataRKH_(startDate, endDate);
    return { success: false, error: 'Action tidak valid', data: [], headers: [] };
  } catch (err) {
    return { success: false, error: String(err && err.message ? err.message : err), data: [], headers: [] };
  }
}

/** Output JSON / JSONP */
function output_(obj, callback) {
  const json = JSON.stringify(obj);

  if (callback && String(callback).trim()) {
    const cb = String(callback).trim().replace(/[^\w$.]/g, ''); // sanitize
    const js = `${cb}(${json});`;
    return ContentService.createTextOutput(js)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ==========================================================
   CONFIG
========================================================== */
const CFG_ = {
  PUSINGAN: {
    ssId: '1WhXyDcHy6iZG9BJTKygaShQr29cqSJeF0oqwRs39b18',
    sheet: 'pusingan'
  },
  RKH: {
    ssId: '1CnEyvAnCOseyWi2APtWTPr6admFJnfljbC4jF-y_kNU',
    sheet: 'rkh',
    masterSheet: 'master_asisten'
  },
  WIB_OFFSET_HOURS: 7
};

/* ==========================================================
   HELPERS: date parsing + formatting (explicit WIB conversion)
========================================================== */

/** Parse dd-mm-yyyy OR yyyy-mm-dd OR Date -> Date (UTC-based "date only") */
function parseDateOnly_(value) {
  if (!value) return null;

  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    // normalize to date-only using local Y/M/D of the Date object
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0, 0));
  }

  if (typeof value === 'string') {
    const s = value.trim();
    // dd-mm-yyyy
    let m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    if (m) {
      const d = parseInt(m[1], 10);
      const mo = parseInt(m[2], 10) - 1;
      const y = parseInt(m[3], 10);
      const dt = new Date(Date.UTC(y, mo, d, 0, 0, 0, 0));
      return isNaN(dt.getTime()) ? null : dt;
    }
    // yyyy-mm-dd
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) {
      const y = parseInt(m[1], 10);
      const mo = parseInt(m[2], 10) - 1;
      const d = parseInt(m[3], 10);
      const dt = new Date(Date.UTC(y, mo, d, 0, 0, 0, 0));
      return isNaN(dt.getTime()) ? null : dt;
    }

    // fall back: Date parse
    const dt = new Date(s);
    if (!isNaN(dt.getTime())) {
      return new Date(Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate(), 0, 0, 0, 0));
    }
  }

  return null;
}

/** Convert an ISO/Z time (UTC) to "WIB as UTC-shifted date" */
function toWibDateTime_(value) {
  if (!value) return null;
  const utc = (value instanceof Date) ? value : new Date(String(value));
  if (isNaN(utc.getTime())) return null;
  const wib = new Date(utc.getTime() + CFG_.WIB_OFFSET_HOURS * 60 * 60 * 1000);
  return wib;
}

function fmtDDMMYYYY_(dateObjLike) {
  if (!dateObjLike) return '';
  const d = (dateObjLike instanceof Date) ? dateObjLike : new Date(dateObjLike);
  if (isNaN(d.getTime())) return '';
  // Use UTC getters to avoid script timezone surprises.
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const yy = d.getUTCFullYear();
  return `${dd}/${mm}/${yy}`;
}

function fmtHHMM_(dateObjLike) {
  if (!dateObjLike) return '';
  const d = (dateObjLike instanceof Date) ? dateObjLike : new Date(dateObjLike);
  if (isNaN(d.getTime())) return '';
  const hh = String(d.getUTCHours()).padStart(2, '0');
  const mi = String(d.getUTCMinutes()).padStart(2, '0');
  return `${hh}:${mi}`;
}

function keyYMD_(dateOnlyUtc) {
  if (!dateOnlyUtc || !(dateOnlyUtc instanceof Date) || isNaN(dateOnlyUtc.getTime())) return '';
  const y = dateOnlyUtc.getUTCFullYear();
  const m = String(dateOnlyUtc.getUTCMonth() + 1).padStart(2, '0');
  const d = String(dateOnlyUtc.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/* ==========================================================
   PUSINGAN
   Output columns:
     nik; report_date; send_date; send_time; nilai; nama
========================================================== */
function getDataPusingan_(startDate, endDate) {
  const ss = SpreadsheetApp.openById(CFG_.PUSINGAN.ssId);
  const sh = ss.getSheetByName(CFG_.PUSINGAN.sheet);
  if (!sh) throw new Error(`Sheet ${CFG_.PUSINGAN.sheet} tidak ditemukan`);

  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { success: true, data: [], headers: ['nik','report_date','send_date','send_time','nilai','nama'], count: 0 };
  }

  const headers = values[0].map(String);
  const idx = indexMap_(headers);

  const start = startDate ? parseDateOnly_(startDate) : null;
  const end = endDate ? parseDateOnly_(endDate) : null;

  // keep latest updated_at per report date
  const latest = {}; // keyYMD -> {rowObj, updatedWib}
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (!row || row.length === 0) continue;

    const tRaw = row[idx.tanggal];
    const reportDateOnly = parseDateOnly_(tRaw);
    if (!reportDateOnly) continue;

    if (start && reportDateOnly.getTime() < start.getTime()) continue;
    if (end && reportDateOnly.getTime() > end.getTime()) continue;

        const updRaw = row[idx.updated_at];
    const updatedWib = toWibDateTime_(updRaw) || toWibDateTime_(new Date());

    // Dedup per (tanggal, nik): ambil baris dengan updated_at paling akhir
    const nik = safeString_(row[idx.nik_mandor]);
    if (!nik) continue;

    const k = keyYMD_(reportDateOnly) + '|' + nik;
    if (!latest[k] || updatedWib.getTime() > latest[k].updatedWib.getTime()) {
      latest[k] = { row, reportDateOnly, updatedWib };
    }
  }

  const out = [];
  Object.keys(latest).sort().forEach(k => {
    const it = latest[k];
    const row = it.row;
    const reportDateOnly = it.reportDateOnly; // UTC date-only
    const updatedWib = it.updatedWib;         // shifted, treat with UTC getters

    const nik = safeString_(row[idx.nik_mandor]);
    const nama = safeString_(row[idx.nama_mandor]);

    const report_date = fmtDDMMYYYY_(reportDateOnly);

    const send_date = fmtDDMMYYYY_(dateOnlyFromDateTime_(updatedWib));
    const send_time = fmtHHMM_(updatedWib);

    const nilai = nilaiPusingan_(reportDateOnly, updatedWib);

    out.push({ nik, report_date, send_date, send_time, nilai, nama });
  });

  return {
    success: true,
    data: out,
    headers: ['nik','report_date','send_date','send_time','nilai','nama'],
    count: out.length
  };
}

function nilaiPusingan_(reportDateOnlyUtc, updatedWibShifted) {
  if (!reportDateOnlyUtc || !updatedWibShifted) return '';

  // Build thresholds in WIB (represented as UTC with shift already applied)
  // report date 00:00 WIB => use reportDateOnlyUtc as that baseline (date-only)
  const base = new Date(reportDateOnlyUtc.getTime()); // 00:00
  const t1 = new Date(base.getTime());
  t1.setUTCDate(t1.getUTCDate() + 1);
  t1.setUTCHours(8,0,0,0);

  const t2 = new Date(base.getTime());
  t2.setUTCDate(t2.getUTCDate() + 1);
  t2.setUTCHours(12,0,0,0);

  const t3 = new Date(base.getTime());
  t3.setUTCDate(t3.getUTCDate() + 1);
  t3.setUTCHours(23,59,59,999);

  const upd = updatedWibShifted;
  if (upd.getTime() <= t1.getTime()) return '4';
  if (upd.getTime() <= t2.getTime()) return '3';
  if (upd.getTime() <= t3.getTime()) return '2';
  if (upd.getTime() > t3.getTime()) return '1';
  return '';
}

/* ==========================================================
   RKH
   Output columns:
     divisi_id; nik; report_date; send_date; send_time; nilai; nama
========================================================== */
function getDataRKH_(startDate, endDate) {
  const ss = SpreadsheetApp.openById(CFG_.RKH.ssId);
  const sh = ss.getSheetByName(CFG_.RKH.sheet);
  if (!sh) throw new Error(`Sheet ${CFG_.RKH.sheet} tidak ditemukan`);

  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { success: true, data: [], headers: ['divisi_id','nik','report_date','send_date','send_time','nilai','nama'], count: 0 };
  }

  const headers = values[0].map(String);
  const idx = indexMapRkh_(headers);

  // master_asisten map: divisi_id -> {nik,nama}
  const masterMap = loadMasterAsistenMap_(ss);

  const start = startDate ? parseDateOnly_(startDate) : null;
  const end = endDate ? parseDateOnly_(endDate) : null;

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (!row || row.length === 0) continue;

    const reportDateOnly = parseDateOnly_(row[idx.tanggal]);
    if (!reportDateOnly) continue;

    if (start && reportDateOnly.getTime() < start.getTime()) continue;
    if (end && reportDateOnly.getTime() > end.getTime()) continue;

    const divisi_id = safeString_(row[idx.divisi_id]);
    const m = masterMap[divisi_id] || { nik: '', nama: '' };

    const updatedWib = toWibDateTime_(row[idx.updated_at]) || toWibDateTime_(new Date());

    const report_date = fmtDDMMYYYY_(reportDateOnly);
    const send_date = fmtDDMMYYYY_(dateOnlyFromDateTime_(updatedWib));
    const send_time = fmtHHMM_(updatedWib);

    const nilai = nilaiRkh_(reportDateOnly, updatedWib);

    out.push({
      divisi_id,
      nik: m.nik || '',
      report_date,
      send_date,
      send_time,
      nilai,
      nama: m.nama || ''
    });
  }

  return {
    success: true,
    data: out,
    headers: ['divisi_id','nik','report_date','send_date','send_time','nilai','nama'],
    count: out.length
  };
}

function nilaiRkh_(reportDateOnlyUtc, updatedWibShifted) {
  if (!reportDateOnlyUtc || !updatedWibShifted) return '';
  const sendDateOnly = dateOnlyFromDateTime_(updatedWibShifted);
  if (!sendDateOnly) return '';

  // diffDays = sendDate - reportDate (based on date-only)
  const diffDays = Math.round((sendDateOnly.getTime() - reportDateOnlyUtc.getTime()) / (24*60*60*1000));

  if (diffDays === -1) return '4'; // t-1
  if (diffDays === 1) return '1';  // t+1
  return '';
}

/* ==========================================================
   Master Asisten
========================================================== */
function loadMasterAsistenMap_(spreadsheetObj) {
  const sh = spreadsheetObj.getSheetByName(CFG_.RKH.masterSheet);
  if (!sh) {
    // If sheet not found, return empty map (but don't fail hard)
    return {};
  }

  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return {};

  const headers = values[0].map(String);
  const idxNik = headers.indexOf('nik');
  const idxNama = headers.indexOf('nama');
  const idxDiv = headers.indexOf('divisi_id');

  if (idxNik < 0 || idxNama < 0 || idxDiv < 0) return {};

  const map = {};
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const div = safeString_(row[idxDiv]);
    if (!div) continue;
    map[div] = {
      nik: safeString_(row[idxNik]),
      nama: safeString_(row[idxNama])
    };
  }
  return map;
}

/* ==========================================================
   Small utilities
========================================================== */
function safeString_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function dateOnlyFromDateTime_(dt) {
  if (!dt || !(dt instanceof Date) || isNaN(dt.getTime())) return null;
  return new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate(), 0,0,0,0));
}

function indexMap_(headers) {
  const m = {};
  function ix(name){ return headers.indexOf(name); }
  m.tanggal = ix('tanggal');
  m.updated_at = ix('updated_at');
  m.nik_mandor = ix('nik_mandor');
  m.nama_mandor = ix('nama_mandor');

  // Basic validation
  ['tanggal','updated_at','nik_mandor','nama_mandor'].forEach(k => {
    if (m[k] < 0) throw new Error(`Kolom "${k}" tidak ditemukan di sheet pusingan`);
  });

  return m;
}

function indexMapRkh_(headers) {
  const m = {};
  function ix(name){ return headers.indexOf(name); }
  m.tanggal = ix('tanggal');
  m.updated_at = ix('updated_at');
  m.divisi_id = ix('divisi_id');

  ['tanggal','updated_at','divisi_id'].forEach(k => {
    if (m[k] < 0) throw new Error(`Kolom "${k}" tidak ditemukan di sheet rkh`);
  });

  return m;
}
