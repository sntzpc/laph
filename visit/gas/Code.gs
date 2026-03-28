const SPREADSHEET_ID = '1czxW4yH9OZv0ag9GnO4E-EqY8aNuU_nD6KxMmhs2zZg';
const PARTICIPANTS_SHEET = 'participants';
const REPORTS_SHEET = 'reports';

const PARTICIPANTS_HEADERS = [
  'nik','nama','jenis_pelatihan','tahun','lokasi_ojt','unit','region','group',
  'createdAt','updatedAt','sourceDevice'
];

const REPORTS_HEADERS = [
  'id','visitDate','location','mentees_json','mentee_names','unit_list','summary','activityReview',
  'technicalUnderstanding','fieldObservation','microTeaching','livingCondition','motivationFuture',
  'resultsObtained','followUp','specialNotes','createdAt','updatedAt','sourceDevice'
];

function doGet(e) {
  const params = e.parameter || {};
  const callback = params.callback;
  const result = handleRequest_(params, null);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const params = Object.assign({}, e.parameter || {});
  const result = handleRequest_(params, e);
  const opId = params.opId || '';
  const html = HtmlService.createHtmlOutput(
    '<script>' +
    'window.top.postMessage(' + JSON.stringify({ __visitSync: true, opId: opId, payload: result }) + ', "*");' +
    '</script>'
  );
  return html;
}

function handleRequest_(params, e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureSheets_(ss);
    const action = String(params.action || '').trim();

    if (action === 'init') {
      return {
        ok: true,
        message: 'Spreadsheet siap dipakai.',
        stats: {
          participants: countDataRows_(ss.getSheetByName(PARTICIPANTS_SHEET)),
          reports: countDataRows_(ss.getSheetByName(REPORTS_SHEET))
        }
      };
    }

    if (action === 'pull') {
      return {
        ok: true,
        data: {
          participants: getSheetObjects_(ss.getSheetByName(PARTICIPANTS_SHEET)),
          reports: getSheetObjects_(ss.getSheetByName(REPORTS_SHEET))
        }
      };
    }

    if (action === 'getOne') {
      const entityType = String(params.entityType || '').trim();
      const entityId = String(params.entityId || '').trim();
      if (!entityType || !entityId) throw new Error('entityType dan entityId wajib diisi.');
      const row = entityType === 'participant'
        ? findByKey_(ss.getSheetByName(PARTICIPANTS_SHEET), PARTICIPANTS_HEADERS, 'nik', entityId)
        : findByKey_(ss.getSheetByName(REPORTS_SHEET), REPORTS_HEADERS, 'id', entityId);
      return { ok: true, data: { exists: !!row, row: row } };
    }

    if (action === 'upsert') {
      const entityType = String(params.entityType || '').trim();
      const payload = parsePayload_(params.payload);
      if (entityType === 'participant') {
        upsertParticipant_(ss, payload);
        return { ok: true, message: 'Participant berhasil di-upsert.', id: payload.nik };
      }
      if (entityType === 'report') {
        upsertReport_(ss, payload);
        return { ok: true, message: 'Report berhasil di-upsert.', id: payload.id };
      }
      throw new Error('entityType tidak dikenal.');
    }

    throw new Error('Action tidak dikenal.');
  } catch (err) {
    return { ok: false, message: err.message || String(err) };
  }
}

function ensureSheets_(ss) {
  let participants = ss.getSheetByName(PARTICIPANTS_SHEET);
  if (!participants) participants = ss.insertSheet(PARTICIPANTS_SHEET);
  ensureHeaders_(participants, PARTICIPANTS_HEADERS);

  let reports = ss.getSheetByName(REPORTS_SHEET);
  if (!reports) reports = ss.insertSheet(REPORTS_SHEET);
  ensureHeaders_(reports, REPORTS_HEADERS);
}

function ensureHeaders_(sheet, headers) {
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const same = headers.every((h, i) => String(current[i] || '') === h);
  if (!same) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function upsertParticipant_(ss, row) {
  const sheet = ss.getSheetByName(PARTICIPANTS_SHEET);
  const clean = {
    nik: str_(row.nik),
    nama: str_(row.nama),
    jenis_pelatihan: str_(row.jenis_pelatihan),
    tahun: str_(row.tahun),
    lokasi_ojt: str_(row.lokasi_ojt),
    unit: str_(row.unit),
    region: str_(row.region),
    group: str_(row.group),
    createdAt: str_(row.createdAt || row.updatedAt || new Date().toISOString()),
    updatedAt: str_(row.updatedAt || new Date().toISOString()),
    sourceDevice: str_(row.sourceDevice)
  };
  upsertByKey_(sheet, PARTICIPANTS_HEADERS, 'nik', clean);
}

function upsertReport_(ss, row) {
  const sheet = ss.getSheetByName(REPORTS_SHEET);
  const mentees = Array.isArray(row.mentees) ? row.mentees : [];
  const clean = {
    id: str_(row.id),
    visitDate: str_(row.visitDate),
    location: str_(row.location),
    mentees_json: JSON.stringify(mentees),
    mentee_names: mentees.map(function(m){ return str_(m.nama); }).filter(String).join('; '),
    unit_list: mentees.map(function(m){ return str_(m.unit); }).filter(String).filter(onlyUnique_).join('; '),
    summary: str_(row.summary),
    activityReview: str_(row.activityReview),
    technicalUnderstanding: str_(row.technicalUnderstanding),
    fieldObservation: str_(row.fieldObservation),
    microTeaching: str_(row.microTeaching),
    livingCondition: str_(row.livingCondition),
    motivationFuture: str_(row.motivationFuture),
    resultsObtained: str_(row.resultsObtained),
    followUp: str_(row.followUp),
    specialNotes: str_(row.specialNotes),
    createdAt: str_(row.createdAt || row.updatedAt || new Date().toISOString()),
    updatedAt: str_(row.updatedAt || new Date().toISOString()),
    sourceDevice: str_(row.sourceDevice)
  };
  upsertByKey_(sheet, REPORTS_HEADERS, 'id', clean);
}

function findByKey_(sheet, headers, keyHeader, keyValue) {
  const keyIndex = headers.indexOf(keyHeader) + 1;
  if (keyIndex < 1) throw new Error('Header key tidak ditemukan.');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const idx = values.findIndex(function(row){ return String(row[keyIndex - 1]) === String(keyValue); });
  if (idx < 0) return null;
  var obj = {};
  headers.forEach(function(h, i){ obj[h] = values[idx][i]; });
  return obj;
}

function upsertByKey_(sheet, headers, keyHeader, rowObj) {
  const keyIndex = headers.indexOf(keyHeader) + 1;
  if (keyIndex < 1) throw new Error('Header key tidak ditemukan.');
  const lastRow = sheet.getLastRow();
  const keys = lastRow > 1 ? sheet.getRange(2, keyIndex, lastRow - 1, 1).getValues().flat() : [];
  const existingIndex = keys.findIndex(function(v){ return String(v) === String(rowObj[keyHeader]); });
  const values = headers.map(function(h){ return rowObj[h] || ''; });

  if (existingIndex >= 0) {
    sheet.getRange(existingIndex + 2, 1, 1, headers.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }
}

function getSheetObjects_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return [];
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values.shift();
  return values
    .filter(function(row){ return row.some(function(cell){ return String(cell) !== ''; }); })
    .map(function(row){
      var obj = {};
      headers.forEach(function(header, i){ obj[header] = row[i]; });
      return obj;
    });
}

function countDataRows_(sheet) {
  return Math.max(0, sheet.getLastRow() - 1);
}

function parsePayload_(raw) {
  if (!raw) throw new Error('payload kosong.');
  return JSON.parse(raw);
}

function onlyUnique_(value, index, array) {
  return array.indexOf(value) === index;
}

function str_(value) {
  return value == null ? '' : String(value);
}
