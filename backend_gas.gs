/************************************
 * Laporan Harian Mentee - GAS Backend
 * Auto-create sheets + header, anti-duplikat,
 * pencarian peserta, list laporan, dan upsert.
 * Timezone: Asia/Jakarta
 *
 * Penyesuaian:
 * 1) Edit frontend kini diutamakan overwrite berdasarkan id.
 * 2) Jika report_date / nik berubah saat edit, _key ikut diperbarui.
 * 3) Jika key baru sudah ada pada baris lain, data lama dioverwrite, bukan append.
 * 4) Delete tetap mendukung id atau nik+report_date.
 * 5) listHolidays diperbaiki agar action lowercase tetap terbaca.
 ************************************/

const CFG = {
  SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(), // atau isi manual kalau perlu
  TZ: 'Asia/Jakarta',
  SHEETS: {
    PARTICIPANTS: 'master_participants',
    REPORTS: 'laporan_harian',
    USERS: 'master_users',
    HOLIDAYS: 'holidays'
  },
  HEADERS: {
    PARTICIPANTS: ['nik','nama','program','divisi','unit','region','group','is_active'],
    REPORTS: ['id','nik','report_date','send_date','send_time','score','synced_at','created_at','updated_at','_key'],
    USERS: ['username','nama','role','status','created_at','updated_at'],
    HOLIDAYS: ['tanggal','keterangan']
  }
};

// ===== Utilities =====
function ss() {
  return SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
}

function ensureSheet_(name, headers) {
  const wb = ss();
  let sh = wb.getSheetByName(name);
  if (!sh) {
    sh = wb.insertSheet(name);
  }
  const firstRow = sh.getRange(1,1,1,headers.length).getValues()[0];
  const hasHeaders = firstRow.some(v => String(v||'').trim() !== '');
  if (!hasHeaders) {
    sh.clear();
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.autoResizeColumns(1, headers.length);
  } else {
    const existing = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
    const toAdd = headers.filter(h => !existing.includes(h));
    if (toAdd.length) {
      sh.insertColumnsAfter(existing.length, toAdd.length);
      sh.getRange(1, existing.length+1, 1, toAdd.length).setValues([toAdd]);
    }
  }
  return sh;
}

function headerMap_(sh) {
  const lastCol = sh.getLastColumn() || 1;
  const hdr = sh.getRange(1,1,1,lastCol).getValues()[0];
  const map = {};
  hdr.forEach((h,i)=> map[String(h||'').trim()] = i+1);
  return map;
}

function nowWIB_() {
  return Utilities.formatDate(new Date(), CFG.TZ, "yyyy-MM-dd' 'HH:mm:ss");
}

function buildKey_(nik, reportDate) {
  return String(nik||'').trim() + '|' + String(reportDate||'').trim();
}

function findRowByKey_(sh, keyColIdx, keyValue) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  const rng = sh.getRange(2, keyColIdx, lastRow-1, 1).getValues();
  for (let i=0;i<rng.length;i++){
    if (String(rng[i][0]||'') === String(keyValue||'')) {
      return i+2;
    }
  }
  return -1;
}

function asJSON_(obj, statusCode=200) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== Extra Helpers (tanggal/jam & skor) =====
function to2_(n){ return ('0' + n).slice(-2); }

function parseDateFlexibleToDDMMYYYY_(s){
  s = String(s||'').trim();
  if (!s) return '';
  let m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return m[1]+'/'+m[2]+'/'+m[3];
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[3]+'/'+m[2]+'/'+m[1];
  return '';
}

function normalizeTimeHHMM_(s){
  const m = String(s||'').trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!m) return '';
  const hh = Math.max(0, Math.min(23, parseInt(m[1],10)||0));
  const mm = Math.max(0, Math.min(59, parseInt(m[2],10)||0));
  return to2_(hh)+':'+to2_(mm);
}

function parseDDMMYYYYToDate_(s){
  const m = String(s||'').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return null;
  const d = parseInt(m[1],10), mo = parseInt(m[2],10)-1, y = parseInt(m[3],10);
  return new Date(y, mo, d, 0, 0, 0, 0);
}

function buildLocalDateTime_(ddmmyyyy, hhmm){
  const base = parseDDMMYYYYToDate_(ddmmyyyy);
  if (!base) return null;
  const m = String(hhmm||'').match(/^(\d{2}):(\d{2})$/);
  const hh = m ? parseInt(m[1],10) : 0;
  const mm = m ? parseInt(m[2],10) : 0;
  return new Date(base.getFullYear(), base.getMonth(), base.getDate(), hh, mm, 0, 0);
}

function computeDailyScoreServer_(report_date_ddmmyyyy, send_date_ddmmyyyy, send_time_hhmm){
  if (!report_date_ddmmyyyy || !send_date_ddmmyyyy || !send_time_hhmm) return '';
  const H = parseDDMMYYYYToDate_(report_date_ddmmyyyy);
  const sent = buildLocalDateTime_(send_date_ddmmyyyy, send_time_hhmm);
  if (!H || !sent) return '';

  const H1_0800 = new Date(H); H1_0800.setDate(H1_0800.getDate()+1); H1_0800.setHours(8,0,0,0);
  const H1_1200 = new Date(H); H1_1200.setDate(H1_1200.getDate()+1); H1_1200.setHours(12,0,0,0);
  const H1_2359 = new Date(H); H1_2359.setDate(H1_2359.getDate()+1); H1_2359.setHours(23,59,59,999);

  if (sent <= H1_0800) return 4;
  if (sent <= H1_1200) return 3;
  if (sent <= H1_2359) return 2;
  return 1;
}

function buildParticipantsIndex_(){
  const sh  = ensureSheet_(CFG.SHEETS.PARTICIPANTS, CFG.HEADERS.PARTICIPANTS);
  const map = headerMap_(sh);
  const last = sh.getLastRow();
  const idx = { known:{}, active:{} };
  if (last >= 2){
    const data = sh.getRange(2,1,last-1, sh.getLastColumn()).getValues();
    const colNik = map['nik']||1;
    const colAct = map['is_active']||0;
    data.forEach(row=>{
      const nik = String(row[colNik-1]||'').trim();
      if (!nik) return;
      idx.known[nik] = true;
      const isAct = String(row[colAct-1]).toLowerCase();
      const active = !(isAct==='false' || isAct==='0' || isAct==='no' || isAct==='nonaktif');
      if (active) idx.active[nik] = true;
    });
  }
  return idx;
}

function normalizeReportPayload_(p){
  const nik = String(p.nik || p.NIK || '').trim();
  const report_date = parseDateFlexibleToDDMMYYYY_(p.report_date || p.reportDate || '');
  const send_date = parseDateFlexibleToDDMMYYYY_(p.send_date || p.sendDate || '');
  const send_time = normalizeTimeHHMM_(p.send_time || p.sendTime || '');
  const id = String(p.id || p.report_id || '').trim();
  const explicitKey = String(p._key || p.key || '').trim();
  const markSynced = String(p.markSynced ?? p.mark_synced ?? '').toLowerCase()==='true' || p.markSynced===true || p.mark_synced===true;

  let scoreVal = (p.score === '' || p.score === null || typeof p.score === 'undefined') ? '' : Number(p.score);
  if (scoreVal === '' || scoreVal === null || Number.isNaN(scoreVal)) {
    const autoS = computeDailyScoreServer_(report_date, send_date, send_time);
    scoreVal = (autoS !== '') ? autoS : '';
  }

  return {
    id,
    nik,
    report_date,
    send_date,
    send_time,
    score: scoreVal,
    markSynced,
    key: explicitKey || buildKey_(nik, report_date)
  };
}

function deleteRowSafely_(sh, rowIdx){
  if (rowIdx > 1 && rowIdx <= sh.getLastRow()) sh.deleteRow(rowIdx);
}

// ====== Public API ======
function doGet(e) {
  try {
    bootstrap_();
    const action = String(e.parameter.action||'').toLowerCase();

    if (action === 'listparticipants') return handleListParticipants_(e);
    if (action === 'listreports') return handleListReports_(e);
    if (action === 'stats') return handleStats_();
    if (action === 'listusers') return handleListUsers_();
    if (action === 'listholidays') return handleListHolidays_();

    return asJSON_({
      ok:true,
      message:'GAS up & running',
      endpoints:['listParticipants','listReports','stats','upsertReport','importParticipants','listUsers','upsertUser','deleteReport','importReports','listHolidays']
    });
  } catch(err) {
    return asJSON_({ok:false, error:String(err)});
  }
}

function doPost(e) {
  try {
    bootstrap_();
    const payload = parseBody_(e);
    const action = String(payload.action||'').toLowerCase();

    if (action === 'upsertreport') return handleUpsertReport_(payload);
    if (action === 'importparticipants') return handleImportParticipants_(payload);
    if (action === 'upsertuser') return handleUpsertUser_(payload);
    if (action === 'deletereport') return handleDeleteReport_(payload);
    if (action === 'importreports') return handleImportReports_(payload);

    return asJSON_({ok:false, error:'Unknown action'});
  } catch(err) {
    return asJSON_({ok:false, error:String(err)});
  }
}

function bootstrap_(){
  ensureSheet_(CFG.SHEETS.PARTICIPANTS, CFG.HEADERS.PARTICIPANTS);
  ensureSheet_(CFG.SHEETS.REPORTS, CFG.HEADERS.REPORTS);
  ensureSheet_(CFG.SHEETS.USERS, CFG.HEADERS.USERS);
  ensureSheet_(CFG.SHEETS.HOLIDAYS, CFG.HEADERS.HOLIDAYS);
}

function handleListParticipants_(e){
  const q = (e.parameter.q||'').toLowerCase();
  const limitParam = parseInt(e.parameter.limit, 10);
  const LIMIT = (!isNaN(limitParam) && limitParam > 0) ? limitParam : 10000;

  const sh = ensureSheet_(CFG.SHEETS.PARTICIPANTS, CFG.HEADERS.PARTICIPANTS);
  const map = headerMap_(sh);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return asJSON_({ok:true, data:[]});

  const data = sh.getRange(2,1,lastRow-1, sh.getLastColumn()).getValues();
  const out = [];
  const idxNik = map['nik'], idxNama = map['nama'], idxProgram = map['program'], idxDiv = map['divisi'],
        idxUnit = map['unit'], idxReg = map['region'], idxGrp = map['group'], idxActive = map['is_active'];

  for (const row of data) {
    const nik   = row[idxNik-1];
    const nama  = row[idxNama-1];
    const program = row[idxProgram-1];
    const div   = row[idxDiv-1];
    const unit  = row[idxUnit-1];
    const reg   = row[idxReg-1];
    const grp   = row[idxGrp-1];
    const active= String(row[idxActive-1]||'true').toLowerCase() !== 'false';

    if (!active) continue;

    const needle = (String(nik||'')+' '+String(nama||'')).toLowerCase();
    if (!q || needle.includes(q)) {
      out.push({nik, nama, program, divisi:div, unit, region:reg, group:grp});
      if (out.length >= LIMIT) break;
    }
  }
  return asJSON_({ok:true, data:out});
}

function handleListHolidays_(){
  const sh = ensureSheet_(CFG.SHEETS.HOLIDAYS, CFG.HEADERS.HOLIDAYS);
  const map = headerMap_(sh);
  const last = sh.getLastRow();
  const out = [];
  if (last >= 2){
    const vals = sh.getRange(2,1,last-1, sh.getLastColumn()).getValues();
    const colTgl = map['tanggal']||1;
    vals.forEach(r=>{
      const raw = r[colTgl-1];
      let iso = '';
      if (raw instanceof Date) {
        iso = Utilities.formatDate(raw, CFG.TZ, 'yyyy-MM-dd');
      } else {
        const s = String(raw||'').trim();
        let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (m) iso = s;
        else {
          m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
          if (m) iso = `${m[3]}-${m[2]}-${m[1]}`;
        }
      }
      if (iso) out.push(iso);
    });
  }
  return asJSON_({ ok:true, data: out });
}

function handleListReports_(e){
  const sh = ensureSheet_(CFG.SHEETS.REPORTS, CFG.HEADERS.REPORTS);
  const map = headerMap_(sh);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return asJSON_({ok:true, data:[]});

  const data = sh.getRange(2,1,lastRow-1, sh.getLastColumn()).getValues();
  const out = [];
  for (const row of data) {
    const obj = {};
    Object.keys(map).forEach(h=>{
      obj[h] = row[map[h]-1];
    });
    out.push(obj);
  }
  return asJSON_({ok:true, data:out});
}

function handleStats_(){
  const sh = ensureSheet_(CFG.SHEETS.REPORTS, CFG.HEADERS.REPORTS);
  const lastRow = sh.getLastRow();
  const total = Math.max(0, lastRow - 1);
  return asJSON_({
    ok:true,
    data:{
      total_peserta: 0,
      rata_rata_nilai: null,
      total_sync: total
    }
  });
}

function handleUpsertReport_(payload){
  const norm = normalizeReportPayload_(payload || {});
  if (!norm.nik || !norm.report_date) {
    return asJSON_({ok:false, error:'nik dan report_date wajib diisi'});
  }

  const idxNik = buildParticipantsIndex_();
  if (!idxNik.known[norm.nik]) {
    return asJSON_({ ok:false, error:'nik_not_found', message:'NIK tidak terdaftar di master_participants' });
  }
  if (!idxNik.active[norm.nik]) {
    return asJSON_({ ok:false, error:'nik_inactive', message:'NIK ada di master_participants namun statusnya nonaktif' });
  }

  const sh = ensureSheet_(CFG.SHEETS.REPORTS, CFG.HEADERS.REPORTS);
  const map = headerMap_(sh);
  const now = nowWIB_();
  const idCol = map['id'];
  const keyCol = map['_key'];

  const rowById = norm.id ? findRowByKey_(sh, idCol, norm.id) : -1;
  const rowByKey = findRowByKey_(sh, keyCol, norm.key);

  // Kasus edit: id ditemukan, tapi key baru bentrok dengan baris lain.
  // Strategi: overwrite baris target key baru, lalu hapus baris lama berdasarkan id.
  if (rowById > -1 && rowByKey > -1 && rowById !== rowByKey) {
    const targetUpdates = {
      id: norm.id,
      nik: norm.nik,
      report_date: norm.report_date,
      send_date: norm.send_date,
      send_time: norm.send_time,
      score: norm.score,
      updated_at: now,
      _key: norm.key
    };
    if (norm.markSynced) targetUpdates.synced_at = now;
    writeObjectRow_(sh, map, rowByKey, targetUpdates);
    deleteRowSafely_(sh, rowById);
    return asJSON_({ok:true, updated:true, merged:true, by:'id+key', key:norm.key, id:norm.id});
  }

  const targetRow = rowById > -1 ? rowById : rowByKey;

  if (targetRow > -1) {
    const updates = {
      id: norm.id || sh.getRange(targetRow, idCol).getValue() || Utilities.getUuid(),
      nik: norm.nik,
      report_date: norm.report_date,
      send_date: norm.send_date,
      send_time: norm.send_time,
      score: norm.score,
      updated_at: now,
      _key: norm.key
    };
    if (norm.markSynced) updates.synced_at = now;
    writeObjectRow_(sh, map, targetRow, updates);
    return asJSON_({ok:true, updated:true, by:(rowById > -1 ? 'id' : 'key'), key:norm.key, id:updates.id});
  }

  const id = norm.id || Utilities.getUuid();
  const rowObj = {
    id,
    nik: norm.nik,
    report_date: norm.report_date,
    send_date: norm.send_date,
    send_time: norm.send_time,
    score: norm.score,
    synced_at: norm.markSynced ? now : '',
    created_at: now,
    updated_at: now,
    _key: norm.key
  };
  appendObjectRow_(sh, map, rowObj);
  return asJSON_({ok:true, created:true, by:'new', key:norm.key, id});
}

function handleImportReports_(payload){
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  if (!rows.length) return asJSON_({ok:false, error:'no rows'});

  const markSynced = (payload.markSynced===undefined) ? true :
                   (String(payload.markSynced||'').toLowerCase()==='true' || payload.markSynced===true);

  const sh  = ensureSheet_(CFG.SHEETS.REPORTS, CFG.HEADERS.REPORTS);
  const map = headerMap_(sh);
  const idxNik = buildParticipantsIndex_();
  const idCol = map['id'];
  const keyCol = map['_key'];

  let inserted = 0, updated = 0, skipped = 0;
  const keysProcessed = [];
  const errors = [];

  rows.forEach(function(r, i){
    try{
      const norm = normalizeReportPayload_(Object.assign({}, r, { markSynced: markSynced }));

      if (!norm.nik){
        skipped++;
        errors.push({ index:i, key:'', nik:'', report_date:'', reason:'missing_nik', message:'Kolom NIK kosong' });
        return;
      }
      if (!norm.report_date){
        skipped++;
        errors.push({ index:i, key:'', nik:norm.nik, report_date:String(r.report_date||''), reason:'invalid_report_date', message:'Format report_date tidak valid (harus dd/mm/yyyy atau ISO)' });
        return;
      }
      if (!idxNik.known[norm.nik]){
        skipped++;
        errors.push({ index:i, key:norm.key, nik:norm.nik, report_date:norm.report_date, reason:'nik_not_found', message:'NIK tidak ditemukan di master_participants' });
        return;
      }
      if (!idxNik.active[norm.nik]){
        skipped++;
        errors.push({ index:i, key:norm.key, nik:norm.nik, report_date:norm.report_date, reason:'nik_inactive', message:'NIK nonaktif di master_participants' });
        return;
      }

      const now = nowWIB_();
      const rowById = norm.id ? findRowByKey_(sh, idCol, norm.id) : -1;
      const rowByKey = findRowByKey_(sh, keyCol, norm.key);

      if (rowById > -1 && rowByKey > -1 && rowById !== rowByKey) {
        const targetUpdates = {
          id: norm.id,
          nik: norm.nik,
          report_date: norm.report_date,
          send_date: norm.send_date,
          send_time: norm.send_time,
          score: norm.score,
          updated_at: now,
          _key: norm.key
        };
        if (markSynced) targetUpdates.synced_at = now;
        writeObjectRow_(sh, map, rowByKey, targetUpdates);
        deleteRowSafely_(sh, rowById);
        updated++; keysProcessed.push(norm.key);
        return;
      }

      const targetRow = rowById > -1 ? rowById : rowByKey;
      if (targetRow > -1){
        const updates = {
          id: norm.id || sh.getRange(targetRow, idCol).getValue() || Utilities.getUuid(),
          nik: norm.nik,
          report_date: norm.report_date,
          send_date: norm.send_date,
          send_time: norm.send_time,
          score: norm.score,
          updated_at: now,
          _key: norm.key
        };
        if (markSynced) updates.synced_at = now;
        writeObjectRow_(sh, map, targetRow, updates);
        updated++; keysProcessed.push(norm.key);
      } else {
        const id = norm.id || Utilities.getUuid();
        const obj = {
          id: id,
          nik: norm.nik,
          report_date: norm.report_date,
          send_date: norm.send_date,
          send_time: norm.send_time,
          score: norm.score,
          synced_at: markSynced ? now : '',
          created_at: now,
          updated_at: now,
          _key: norm.key
        };
        appendObjectRow_(sh, map, obj);
        inserted++; keysProcessed.push(norm.key);
      }
    } catch(e){
      skipped++;
      errors.push({ index:i, key:'', nik:String(r.nik||''), report_date:String(r.report_date||''), reason:'exception', message:String(e) });
    }
  });

  return asJSON_({ ok:true, inserted:inserted, updated:updated, skipped:skipped, keys:keysProcessed, errors:errors });
}

function handleDeleteReport_(payload){
  const p = payload || {};
  const id = String(p.id||'').trim();
  const nik = String(p.nik||'').trim();
  const report_date = parseDateFlexibleToDDMMYYYY_(p.report_date);

  const sh = ensureSheet_(CFG.SHEETS.REPORTS, CFG.HEADERS.REPORTS);
  const map = headerMap_(sh);

  let rowIdx = -1;
  let by = '';

  if (id) {
    const idCol = map['id'];
    if (!idCol) return asJSON_({ok:false, error:'Column id not found'});
    rowIdx = findRowByKey_(sh, idCol, id);
    by = 'id';
  } else if (nik && report_date) {
    const keyCol = map['_key'];
    if (!keyCol) return asJSON_({ok:false, error:'Column _key not found'});
    rowIdx = findRowByKey_(sh, keyCol, buildKey_(nik, report_date));
    by = 'key';
  } else {
    return asJSON_({ok:false, error:'Provide id or (nik + report_date)'});
  }

  if (rowIdx < 0) {
    return asJSON_({ok:true, deleted:false, by, reason:'not found'});
  }

  sh.deleteRow(rowIdx);
  return asJSON_({ok:true, deleted:true, by});
}

function appendObjectRow_(sh, map, obj) {
  const cols = Object.keys(map).length;
  const row = new Array(cols).fill('');
  Object.keys(map).forEach(h=>{
    row[map[h]-1] = h in obj ? obj[h] : '';
  });
  sh.appendRow(row);
}

function writeObjectRow_(sh, map, rowNumber, obj) {
  const cols = Object.keys(map).length;
  const row = sh.getRange(rowNumber,1,1,cols).getValues()[0];
  Object.keys(obj).forEach(h=>{
    if (map[h]) row[map[h]-1] = obj[h];
  });
  sh.getRange(rowNumber,1,1,cols).setValues([row]);
}

function handleImportParticipants_(payload){
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  if (!rows.length) return asJSON_({ok:false, error:'no rows'});
  const sh = ensureSheet_(CFG.SHEETS.PARTICIPANTS, CFG.HEADERS.PARTICIPANTS);
  const map = headerMap_(sh);

  let inserted = 0, updated = 0, skipped = 0;
  rows.forEach(r=>{
    const nik = String(r.nik||'').trim();
    if (!nik) { skipped++; return; }
    const keyColIdx = map['nik'];
    const rowIdx = findRowByKey_(sh, keyColIdx, nik);

    const obj = {
      nik,
      nama: r.nama||'',
      program: r.program||'',
      divisi: r.divisi||'',
      unit: r.unit||'',
      region: r.region||'',
      group: r.group||'',
      is_active: (String(r.is_active||'').toLowerCase()==='true' || String(r.is_active||'').toLowerCase()==='1') ? true : false
    };

    if (rowIdx > -1) {
      writeObjectRow_(sh, map, rowIdx, obj);
      updated++;
    } else {
      appendObjectRow_(sh, map, obj);
      inserted++;
    }
  });

  return asJSON_({ok:true, inserted, updated, skipped});
}

function handleListUsers_(){
  const sh = ensureSheet_(CFG.SHEETS.USERS, CFG.HEADERS.USERS);
  const map = headerMap_(sh);
  const last = sh.getLastRow();
  const out = [];
  if (last >= 2){
    const data = sh.getRange(2,1,last-1, sh.getLastColumn()).getValues();
    data.forEach(row=>{
      const o = {};
      Object.keys(map).forEach(h=> o[h] = row[map[h]-1]);
      out.push(o);
    });
  }
  return asJSON_({ok:true, data:out});
}

function handleUpsertUser_(payload){
  const username = String(payload.username||'').trim();
  if (!username) return asJSON_({ok:false, error:'username required'});
  const sh = ensureSheet_(CFG.SHEETS.USERS, CFG.HEADERS.USERS);
  const map = headerMap_(sh);
  const colIdx = map['username'];
  const rowIdx = findRowByKey_(sh, colIdx, username);
  const now = nowWIB_();

  const obj = {
    username,
    nama: payload.nama||'',
    role: payload.role||'User',
    status: payload.status||'Aktif',
    updated_at: now
  };

  if (rowIdx > -1) {
    writeObjectRow_(sh, map, rowIdx, obj);
    return asJSON_({ok:true, updated:true});
  } else {
    obj.created_at = now;
    appendObjectRow_(sh, map, obj);
    return asJSON_({ok:true, created:true});
  }
}

function parseBody_(e){
  if (e && e.parameter && e.parameter.payload) {
    try { return JSON.parse(e.parameter.payload); } catch(_) {}
  }
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch(_) {}
  }
  return {};
}
