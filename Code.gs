const SHEET_NAME = 'LunchBoardData';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange('A1:B3').setValues([
      ['key','json'],
      ['roster','[]'],
      ['schedule','{}']
    ]);
  }
  return sh;
}

function readData_() {
  const sh = getSheet_();
  const values = sh.getRange(2,1,2,2).getValues();
  const out = {};
  values.forEach(r => {
    try { out[r[0]] = JSON.parse(r[1] || (r[0] === 'roster' ? '[]' : '{}')); }
    catch(e) { out[r[0]] = r[0] === 'roster' ? [] : {}; }
  });
  return { roster: out.roster || [], schedule: out.schedule || {} };
}

function writeKey_(key, obj) {
  const sh = getSheet_();
  const data = sh.getRange(1,1,sh.getLastRow(),2).getValues();
  let row = data.findIndex(r => r[0] === key) + 1;
  if (!row) row = sh.getLastRow() + 1;
  sh.getRange(row,1,1,2).setValues([[key, JSON.stringify(obj)]]);
}

function doGet(e) {
  const data = readData_();
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  if (body.action === 'saveRoster') writeKey_('roster', body.roster || []);
  if (body.action === 'saveSchedule') writeKey_('schedule', body.schedule || {});
  return ContentService
    .createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}
