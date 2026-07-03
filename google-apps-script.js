/*****
Race Tracker Latest - Google Apps Script Backend

Use this script with the latest Race Tracker workbook layout:
- Entrants
- Start List
- Settings
- Stanage
- Hathersage
- Finish Times
- Dashboard
- Results

Deploy as Web App:
  Execute as: Me
  Who has access: Anyone

The Vercel web app sends queued entries here.
Checkpoint/finish timestamps use the phone tap/input timestamp, not upload time.
*****/

const SHEETS = {
  ENTRANTS: 'Entrants',
  START: 'Start List',
  SETTINGS: 'Settings',
  STANAGE: 'Stanage',
  HATHERSAGE: 'Hathersage',
  FINISH: 'Finish Times',
  DASHBOARD: 'Dashboard',
  RESULTS: 'Results'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Race Tracker')
    .addItem('Refresh layout / formulas', 'setupWorkbook')
    .addItem('Set official start time to now', 'setStartTimeToNow')
    .addToUi();
}

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = String(params.action || 'ping').toLowerCase();
  try {
    if (action === 'setup') {
      setupWorkbook();
      return json_({ ok: true, message: 'Workbook layout refreshed' });
    }
    if (action === 'start_now') {
      setStartTimeToNow();
      return json_({ ok: true, message: 'Official start time set to now' });
    }
    return json_({ ok: true, message: 'Race Tracker backend is working' });
  } catch (err) {
    return json_({ ok: false, message: errorMessage_(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const payload = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : '{}');

    if (payload.type === 'checkpoint') {
      return json_(logCheckpoint(payload));
    }
    if (payload.type === 'finish-timer') {
      return json_(logFinishTime(payload));
    }
    if (payload.type === 'finish-recorder') {
      return json_(assignFinishBib(payload));
    }

    return json_({ ok: false, message: 'Unknown payload type: ' + String(payload.type || '') });
  } catch (err) {
    return json_({ ok: false, message: errorMessage_(err) });
  } finally {
    lock.releaseLock();
  }
}

function setupWorkbook() {
  const entrants = getSheet_(SHEETS.ENTRANTS);
  const start = getSheet_(SHEETS.START);
  const settings = getSheet_(SHEETS.SETTINGS);
  const stanage = getSheet_(SHEETS.STANAGE);
  const hathersage = getSheet_(SHEETS.HATHERSAGE);
  const finish = getSheet_(SHEETS.FINISH);
  const dash = getSheet_(SHEETS.DASHBOARD);
  const results = getSheet_(SHEETS.RESULTS);

  setHeaders_(entrants, ['Bib','First Name','Last Name','Category','Gender','Club','Checked In']);
  setHeaders_(start, ['Bib','First Name','Last Name','Category','Gender','Club']);
  setHeaders_(stanage, ['Submitted At','Race Time','Bib','Name','Category','Gender','Status','Message']);
  setHeaders_(hathersage, ['Submitted At','Race Time','Bib','Name','Category','Gender','Status','Message']);
  setHeaders_(finish, ['Position','Submitted At','Race Time','Bib','Name','Category','Gender','Status','Message']);

  if (!settings.getRange('A1').getValue()) {
    settings.getRange('A1:B8').setValues([
      ['Race Tracker Settings',''],
      ['Race Name','Exterminator Test'],
      ['Official Start Time','10:00:00'],
      ['Checkpoint 1','Stanage'],
      ['Checkpoint 2','Hathersage'],
      ['Timezone','Europe/London'],
      ['Notes','Edit B3 for official start time. Use format hh:mm:ss.'],
      ['Apps Script Web App URL','Paste deployed /exec URL here for reference only']
    ]);
  }

  start.getRange('A2').setFormula('=IFERROR(FILTER(Entrants!A2:F1000,UPPER(Entrants!G2:G1000)="YES",Entrants!A2:A1000<>""),"")');

  buildDashboard_();
  buildResults_();
  formatBase_();
}

function setStartTimeToNow() {
  const settings = getSheet_(SHEETS.SETTINGS);
  settings.getRange('B3').setValue(new Date()).setNumberFormat('hh:mm:ss');
}

function logCheckpoint(payload) {
  const bib = cleanBib_(payload.bib);
  const checkpointSheetName = normaliseCheckpoint_(payload.location || payload.checkpoint);

  if (!bib) return { ok: false, message: 'No bib number supplied' };
  if (!checkpointSheetName) return { ok: false, message: 'Unknown checkpoint location: ' + String(payload.location || '') };

  const sheet = getSheet_(checkpointSheetName);
  ensureCheckpointSheet_(sheet);

  const runner = findRunner_(bib);
  const submittedAt = parseSubmittedAt_(payload.timestamp);
  const raceTimeNumber = raceTimeNumber_(submittedAt);
  const duplicate = isDuplicateCheckpoint_(sheet, bib);

  let status = 'OK';
  let message = 'Logged';

  if (!runner) {
    status = 'UNKNOWN BIB';
    message = 'Bib not found on Start List';
  } else if (duplicate) {
    status = 'DUPLICATE';
    message = 'Runner already logged at this checkpoint';
  }

  sheet.appendRow([
    submittedAt,
    raceTimeNumber,
    bib,
    runner ? runner.name : 'UNKNOWN',
    runner ? runner.category : '',
    runner ? runner.gender : '',
    status,
    message
  ]);

  const row = sheet.getLastRow();
  sheet.getRange(row, 1).setNumberFormat('hh:mm:ss');
  sheet.getRange(row, 2).setNumberFormat('[h]:mm:ss');

  return {
    ok: status === 'OK',
    status,
    message,
    bib,
    checkpoint: checkpointSheetName,
    name: runner ? runner.name : 'UNKNOWN',
    raceTime: sheet.getRange(row, 2).getDisplayValue()
  };
}

function logFinishTime(payload) {
  const finish = getSheet_(SHEETS.FINISH);
  ensureFinishSheet_(finish);

  const submittedAt = parseSubmittedAt_(payload.timestamp);
  const raceTimeNumber = raceTimeNumber_(submittedAt);
  const position = nextFinishPosition_(finish);

  finish.appendRow([
    position,
    submittedAt,
    raceTimeNumber,
    '',
    '',
    '',
    '',
    'UNASSIGNED',
    'Finish time recorded'
  ]);

  const row = finish.getLastRow();
  finish.getRange(row, 2).setNumberFormat('hh:mm:ss');
  finish.getRange(row, 3).setNumberFormat('[h]:mm:ss');

  return {
    ok: true,
    message: 'Finish time recorded',
    position,
    raceTime: finish.getRange(row, 3).getDisplayValue()
  };
}

function assignFinishBib(payload) {
  const bib = cleanBib_(payload.bib);
  if (!bib) return { ok: false, message: 'No bib number supplied' };

  const finish = getSheet_(SHEETS.FINISH);
  ensureFinishSheet_(finish);

  const runner = findRunner_(bib);
  const lastRow = finish.getLastRow();
  if (lastRow < 2) return { ok: false, message: 'No finish times exist yet' };

  const assignedBibs = finish.getRange(2, 4, lastRow - 1, 1).getValues().flat();
  const alreadyAssigned = assignedBibs.some(v => Number(v) === Number(bib));

  const values = finish.getRange(2, 1, lastRow - 1, 9).getValues();
  let targetRow = -1;
  for (let i = 0; i < values.length; i++) {
    if (!values[i][3]) {
      targetRow = i + 2;
      break;
    }
  }
  if (targetRow === -1) return { ok: false, message: 'No unassigned finish time available' };

  let status = 'OK';
  let message = 'Runner assigned to finish time';

  if (!runner) {
    status = 'UNKNOWN BIB';
    message = 'Bib not found on Start List';
  } else if (alreadyAssigned) {
    status = 'DUPLICATE FINISH';
    message = 'Bib already assigned to a finish time';
  }

  finish.getRange(targetRow, 4, 1, 6).setValues([[
    bib,
    runner ? runner.name : 'UNKNOWN',
    runner ? runner.category : '',
    runner ? runner.gender : '',
    status,
    message
  ]]);

  return {
    ok: status === 'OK',
    status,
    message,
    bib,
    name: runner ? runner.name : 'UNKNOWN',
    position: finish.getRange(targetRow, 1).getValue(),
    raceTime: finish.getRange(targetRow, 3).getDisplayValue()
  };
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function setHeaders_(sheet, headers) {
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeaders = existing.join('') === '' || existing[0] !== headers[0];
  if (needsHeaders) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#111827')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

function ensureCheckpointSheet_(sheet) {
  setHeaders_(sheet, ['Submitted At','Race Time','Bib','Name','Category','Gender','Status','Message']);
}

function ensureFinishSheet_(sheet) {
  setHeaders_(sheet, ['Position','Submitted At','Race Time','Bib','Name','Category','Gender','Status','Message']);
}

function formatBase_() {
  [SHEETS.STANAGE, SHEETS.HATHERSAGE].forEach(name => {
    const sh = getSheet_(name);
    sh.getRange('A:A').setNumberFormat('hh:mm:ss');
    sh.getRange('B:B').setNumberFormat('[h]:mm:ss');
    sh.autoResizeColumns(1, 8);
  });

  const finish = getSheet_(SHEETS.FINISH);
  finish.getRange('B:B').setNumberFormat('hh:mm:ss');
  finish.getRange('C:C').setNumberFormat('[h]:mm:ss');
  finish.autoResizeColumns(1, 9);

  getSheet_(SHEETS.SETTINGS).getRange('B3').setNumberFormat('hh:mm:ss');
  getSheet_(SHEETS.RESULTS).getRange('G:G').setNumberFormat('[h]:mm:ss');
  getSheet_(SHEETS.RESULTS).getRange('N:N').setNumberFormat('[h]:mm:ss');
}

function cleanBib_(bib) {
  const cleaned = String(bib || '').replace(/[^0-9]/g, '');
  return cleaned ? Number(cleaned) : '';
}

function normaliseCheckpoint_(location) {
  const raw = String(location || '').trim().toLowerCase();
  if (raw === 'stanage') return SHEETS.STANAGE;
  if (raw === 'hathersage') return SHEETS.HATHERSAGE;
  return '';
}

function findRunner_(bib) {
  SpreadsheetApp.flush();
  const start = getSheet_(SHEETS.START);
  const lastRow = start.getLastRow();
  if (lastRow < 2) return null;

  const values = start.getRange(2, 1, lastRow - 1, 6).getValues();
  for (const row of values) {
    if (Number(row[0]) === Number(bib)) {
      const first = String(row[1] || '').trim();
      const last = String(row[2] || '').trim();
      return {
        bib: Number(row[0]),
        name: (first + ' ' + last).trim(),
        category: row[3],
        gender: row[4],
        club: row[5]
      };
    }
  }
  return null;
}

function isDuplicateCheckpoint_(sheet, bib) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const bibValues = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
  return bibValues.some(v => Number(v) === Number(bib));
}

function nextFinishPosition_(finish) {
  const lastRow = finish.getLastRow();
  if (lastRow < 2) return 1;
  const positions = finish.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(v => v !== '' && !isNaN(Number(v)));
  return positions.length ? Math.max.apply(null, positions.map(Number)) + 1 : 1;
}

function parseSubmittedAt_(timestamp) {
  if (timestamp) {
    const d = new Date(timestamp);
    if (!isNaN(d.getTime())) return d;
  }
  return new Date();
}

function raceTimeNumber_(submittedAt) {
  const settings = getSheet_(SHEETS.SETTINGS);
  const startValue = settings.getRange('B3').getValue();
  const start = new Date(submittedAt);

  if (startValue instanceof Date) {
    start.setHours(startValue.getHours(), startValue.getMinutes(), startValue.getSeconds(), 0);
  } else {
    const parts = String(startValue || '00:00:00').split(':').map(Number);
    start.setHours(parts[0] || 0, parts[1] || 0, parts[2] || 0, 0);
  }

  let diffSeconds = Math.floor((submittedAt.getTime() - start.getTime()) / 1000);
  if (diffSeconds < 0) diffSeconds += 24 * 3600;
  return diffSeconds / 86400;
}

function buildDashboard_() {
  const dash = getSheet_(SHEETS.DASHBOARD);
  dash.clear();

  dash.getRange('A1:H1').merge().setValue('Race Dashboard')
    .setFontSize(18).setFontWeight('bold')
    .setBackground('#111827').setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  dash.getRange('A3:B8').setValues([
    ['Race Name','=Settings!B2'],
    ['Start Time','=Settings!B3'],
    ['Started','=COUNTA(\'Start List\'!A2:A)'],
    ['Logged Stanage','=COUNTIF(Stanage!G:G,"OK")'],
    ['Logged Hathersage','=COUNTIF(Hathersage!G:G,"OK")'],
    ['Finished','=COUNTIF(\'Finish Times\'!H:H,"OK")']
  ]);

  dash.getRange('A10:H10').setValues([['Bib','Name','Category','Gender','Stanage','Hathersage','Finish','Status']])
    .setFontWeight('bold').setBackground('#374151').setFontColor('#ffffff');

  for (let r = 11; r <= 210; r++) {
    const src = r - 9;
    dash.getRange(r, 1, 1, 8).setFormulas([[
      `=IFERROR(IF('Start List'!A${src}="","",'Start List'!A${src}),"")`,
      `=IFERROR(IF(A${r}="","",'Start List'!B${src}&" "&'Start List'!C${src}),"")`,
      `=IFERROR(IF(A${r}="","",'Start List'!D${src}),"")`,
      `=IFERROR(IF(A${r}="","",'Start List'!E${src}),"")`,
      `=IFERROR(IF(A${r}="","",IF(COUNTIF(Stanage!C:C,A${r})>0,"YES","")),"")`,
      `=IFERROR(IF(A${r}="","",IF(COUNTIF(Hathersage!C:C,A${r})>0,"YES","")),"")`,
      `=IFERROR(IF(A${r}="","",IF(COUNTIF('Finish Times'!D:D,A${r})>0,"YES","")),"")`,
      `=IFERROR(IF(A${r}="","",IF(G${r}="YES","Finished",IF(F${r}="YES","Passed Hathersage",IF(E${r}="YES","Passed Stanage","On Course")))),"")`
    ]]);
  }
  dash.setFrozenRows(10);
  dash.autoResizeColumns(1, 8);
}

function buildResults_() {
  const results = getSheet_(SHEETS.RESULTS);
  results.clear();

  results.getRange('A1:G1').setValues([['Overall Pos','Category Pos','Bib','Name','Category','Gender','Finish Time']])
    .setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff');

  results.getRange('A2:A1000').setFormulaR1C1('=IF(RC[2]="","",ROW(RC[2])-1)');
  results.getRange('B2:B1000').setFormulaR1C1('=IF(RC[1]="","",COUNTIFS(R2C5:RC5,RC5,R2C6:RC6,RC6))');
  results.getRange('C2').setFormula('=IFERROR(SORT(FILTER(HSTACK(\'Finish Times\'!D2:D1000,\'Finish Times\'!E2:E1000,\'Finish Times\'!F2:F1000,\'Finish Times\'!G2:G1000,\'Finish Times\'!C2:C1000),\'Finish Times\'!D2:D1000<>""),5,TRUE),"")');

  results.getRange('J1:N1').merge().setValue('Top 3 by Category + Gender')
    .setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  results.getRange('J3:N3').setValues([['Category/Gender','Cat Pos','Bib','Name','Finish Time']])
    .setFontWeight('bold').setBackground('#374151').setFontColor('#ffffff');
  results.getRange('J4').setFormula('=IFERROR(SORT(FILTER(HSTACK(E2:E1000&" "&F2:F1000,B2:B1000,C2:C1000,D2:D1000,G2:G1000),C2:C1000<>"",B2:B1000<=3),1,TRUE,2,TRUE),"")');

  results.getRange('G:G').setNumberFormat('[h]:mm:ss');
  results.getRange('N:N').setNumberFormat('[h]:mm:ss');
  results.setFrozenRows(1);
  results.autoResizeColumns(1, 14);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorMessage_(err) {
  return String(err && err.message ? err.message : err);
}
