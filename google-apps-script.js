/*
Google Apps Script connector for the Race Tracker web interface.

How to use:
1. Open your Google Sheet.
2. Extensions > Apps Script.
3. Paste this whole file into Code.gs.
4. Save.
5. Deploy > New deployment > Web app.
6. Execute as: Me.
7. Who has access: Anyone.
8. Copy the Web App URL.
9. Paste that URL into the web app Settings page.
*/

const SHEET_NAMES = {
  CHECKPOINTS: 'Checkpoint Log',
  FINISH_TIMER: 'Finish Times',
  FINISH_RECORDER: 'Finish Recorder'
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    ensureSheets_(ss);

    if (payload.type === 'checkpoint') {
      ss.getSheetByName(SHEET_NAMES.CHECKPOINTS).appendRow([
        new Date(),
        payload.clockTime || '',
        payload.location || '',
        payload.bib || '',
        payload.timestamp || ''
      ]);
    }

    if (payload.type === 'finish-timer') {
      ss.getSheetByName(SHEET_NAMES.FINISH_TIMER).appendRow([
        new Date(),
        payload.clockTime || '',
        payload.timestamp || ''
      ]);
    }

    if (payload.type === 'finish-recorder') {
      ss.getSheetByName(SHEET_NAMES.FINISH_RECORDER).appendRow([
        new Date(),
        payload.clockTime || '',
        payload.bib || '',
        payload.timestamp || ''
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService
    .createTextOutput('Race Tracker endpoint is live')
    .setMimeType(ContentService.MimeType.TEXT);
}

function ensureSheets_(ss) {
  createSheetIfMissing_(ss, SHEET_NAMES.CHECKPOINTS, [
    'Received At', 'Clock Time', 'Checkpoint', 'Bib Number', 'Device Timestamp'
  ]);

  createSheetIfMissing_(ss, SHEET_NAMES.FINISH_TIMER, [
    'Received At', 'Clock Time', 'Device Timestamp'
  ]);

  createSheetIfMissing_(ss, SHEET_NAMES.FINISH_RECORDER, [
    'Received At', 'Clock Time', 'Bib Number', 'Device Timestamp'
  ]);
}

function createSheetIfMissing_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
}
