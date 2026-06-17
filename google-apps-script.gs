const SHEET_ID = '1Qgo5MPMtL04fCDDxt58VrAWRoOAELoNjw5Op-h-Z3Ko';
const PROGRESS_SHEET_NAME = 'flashcard_progress';

function doGet(e) {
  const learnerId = String(e?.parameter?.learnerId || '').trim();
  if (!learnerId) {
    return jsonResponse_({ ok: false, error: 'Missing learnerId', progress: {}, updatedAt: 0 });
  }

  const sheet = getProgressSheet_();
  const rowIndex = findLearnerRow_(sheet, learnerId);
  if (!rowIndex) {
    return jsonResponse_({ ok: true, learnerId, progress: {}, updatedAt: 0 });
  }

  const row = sheet.getRange(rowIndex, 1, 1, 3).getValues()[0];
  const payload = parsePayload_(row[2]);

  return jsonResponse_({
    ok: true,
    learnerId,
    updatedAt: Number(row[1] || 0) || 0,
    progress: payload
  });
}

function doPost(e) {
  const body = parseBody_(e);
  const learnerId = String(body.learnerId || '').trim();
  if (!learnerId) {
    return jsonResponse_({ ok: false, error: 'Missing learnerId' });
  }

  const sheet = getProgressSheet_();
  const rowIndex = findLearnerRow_(sheet, learnerId);
  const updatedAt = Number(body.updatedAt || Date.now()) || Date.now();
  const payloadText = JSON.stringify(body.progress || {});

  if (rowIndex) {
    sheet.getRange(rowIndex, 1, 1, 3).setValues([[learnerId, updatedAt, payloadText]]);
  } else {
    sheet.appendRow([learnerId, updatedAt, payloadText]);
  }

  return jsonResponse_({ ok: true, learnerId, updatedAt });
}

function parseBody_(e) {
  try {
    return JSON.parse(e?.postData?.contents || '{}');
  } catch (error) {
    return {};
  }
}

function parsePayload_(value) {
  try {
    return JSON.parse(value || '{}');
  } catch (error) {
    return {};
  }
}

function getProgressSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const existingSheet = spreadsheet.getSheetByName(PROGRESS_SHEET_NAME);
  if (existingSheet) return existingSheet;

  const newSheet = spreadsheet.insertSheet(PROGRESS_SHEET_NAME);
  newSheet.getRange(1, 1, 1, 3).setValues([['learnerId', 'updatedAt', 'progressJson']]);
  return newSheet;
}

function findLearnerRow_(sheet, learnerId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let index = 0; index < values.length; index += 1) {
    if (String(values[index][0]).trim() === learnerId) {
      return index + 2;
    }
  }

  return 0;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
