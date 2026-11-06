/*
 * Baby Log: stores the shared log in this Google Sheet. The app's screen is hosted on GitHub Pages.
 * Paste this whole file into Extensions > Apps Script, then deploy it as a web app.
 * The "Log" tab holds every entry. You can read it, or fix entries in it directly.
 */

var LOG_SHEET = 'Log';
var HEADERS = ['ID', 'Type', 'Started', 'Ended', 'Side', 'Logged by'];
var TYPES = { feed: 1, sleep: 1, pee: 1, poop: 1 };
var SIDE_PATTERN = /^(Left|Right)(, (Left|Right)){0,19}$/;
var ID_PATTERN = /^[A-Za-z0-9_-]{1,40}$/;
var HISTORY_DAYS = 60;

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (params.action === 'getData') return json_(getData());
  return ContentService.createTextOutput('This is the Baby Log data link. Open the app from your GitHub Pages link.');
}

/* A failure inside these calls returns an error page, and the app keeps the entry and retries. */
function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'That request could not be read.' });
  }
  if (body.action === 'saveEntry') return json_(saveEntry(body.arg));
  if (body.action === 'deleteEntry') return json_(deleteEntry(body.arg));
  if (body.action === 'setBabyName') return json_(setBabyName(body.arg));
  return json_({ ok: false, error: 'That request could not be read.' });
}

function logSheet_() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = book.insertSheet(LOG_SHEET);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange('C:D').setNumberFormat('yyyy-mm-dd hh:mm');
    sheet.getRange('A:A').setNumberFormat('@');
  }
  return sheet;
}

function toMillis_(value) {
  if (value instanceof Date) return value.getTime();
  if (value === '' || value === null || value === undefined) return null;
  var parsed = new Date(value).getTime();
  return isNaN(parsed) ? null : parsed;
}

function getData() {
  var sheet = logSheet_();
  var lastRow = sheet.getLastRow();
  var entries = [];
  var cutoff = Date.now() - HISTORY_DAYS * 86400000;
  if (lastRow > 1) {
    var rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var start = toMillis_(row[2]);
      var end = toMillis_(row[3]);
      if (!row[0] || !TYPES[row[1]] || start === null) continue;
      if (start < cutoff && end !== null) continue;
      entries.push({
        id: String(row[0]),
        type: String(row[1]),
        start: start,
        end: end === null ? 0 : end,
        side: String(row[4] || ''),
        by: String(row[5] || '')
      });
    }
  }
  return {
    entries: entries,
    baby: PropertiesService.getScriptProperties().getProperty('babyName') || 'Baby'
  };
}

function cleanEntry_(entry) {
  if (!entry || !ID_PATTERN.test(String(entry.id))) return null;
  if (!TYPES[entry.type]) return null;
  var start = Number(entry.start);
  /* A running timer has no end time. The screen sends that as 0. */
  var end = entry.end ? Number(entry.end) : null;
  if (!isFinite(start) || start <= 0) return null;
  if (end !== null && (!isFinite(end) || end < start)) return null;
  var side = '';
  if (entry.type === 'feed') {
    side = String(entry.side || '');
    if (!SIDE_PATTERN.test(side)) return null;
  }
  return {
    id: String(entry.id),
    type: String(entry.type),
    start: start,
    end: end,
    side: side,
    by: cleanText_(entry.by, 20)
  };
}

/* Keeps a typed name from being treated as a Sheets formula. */
function cleanText_(value, maxLength) {
  return String(value || '')
    .replace(/[\u0000-\u001f]/g, '')
    .replace(/^[=+\-@'\s]+/, '')
    .trim()
    .slice(0, maxLength);
}

function findRow_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) return i + 2;
  }
  return 0;
}

function saveEntry(entry) {
  var clean = cleanEntry_(entry);
  if (!clean) return { ok: false, error: 'That entry could not be saved. Check its times and try again.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = logSheet_();
    var values = [
      clean.id,
      clean.type,
      new Date(clean.start),
      clean.end === null ? '' : new Date(clean.end),
      clean.side,
      clean.by
    ];
    var rowNumber = findRow_(sheet, clean.id);
    if (rowNumber) {
      sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function deleteEntry(id) {
  if (!ID_PATTERN.test(String(id))) return { ok: true };
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = logSheet_();
    var rowNumber = findRow_(sheet, String(id));
    if (rowNumber) sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function setBabyName(name) {
  var clean = cleanText_(name, 30) || 'Baby';
  PropertiesService.getScriptProperties().setProperty('babyName', clean);
  return { ok: true };
}
