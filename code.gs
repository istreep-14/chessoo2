/**
 * Google Apps Script Web App
 * - Appends session rows to Sheet1 (with headers)
 * - Logs each problem to Problems (with headers)
 * - Ensures a row exists per Game Key in GameModes (you fill settings manually) (with headers)
 * - Tracks Sitdowns (10-minute idle gap) (with headers)
 * - Maintains DailyStats (single row per date) (with headers)
 *
 * Deploy > Manage deployments > New deployment > Web app
 *   Execute as: Me
 *   Access: Anyone with the link
 */

var SPREADSHEET_ID        = '1krmzTncZE4cTqTIn4HRJLXWKp6_9-sSuZx-t0ie2ZIQ';
var SESSIONS_SHEET_NAME   = 'Sheet1';
var PROBLEMS_SHEET_NAME   = 'Problems';
var GAMEMODES_SHEET_NAME  = 'GameModes';
var SITDOWN_SHEET_NAME    = 'Sitdowns';
var DAILYSTATS_SHEET_NAME = 'DailyStats';

var SITDOWN_IDLE_MS = 10 * 60 * 1000; // 10 minutes

function doPost(e) {
  try {
    var body = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var data = {};
    try { data = JSON.parse(body); } catch (_){ }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    var sessions = getOrCreateSheet_(ss, SESSIONS_SHEET_NAME);
    var sessionsHeaders = [
      'Timestamp','User ID','Score','Page URL','Game Key','Duration Seconds','Problems JSON'
    ];
    ensureHeader_(sessions, sessionsHeaders);

    var timestamp       = String(data.timestamp || new Date().toISOString());
    var userId          = String(data.userId || '');
    var score           = Number(data.score || 0);
    var pageUrl         = String(data.pageUrl || '');
    var gameKey         = extractGameKey_(pageUrl);
    var durationSeconds = (data.durationSeconds != null && data.durationSeconds !== '') ? Number(data.durationSeconds) : '';
    var problems        = (data.problems && data.problems.length) ? data.problems : [];

    sessions.appendRow([
      timestamp,
      userId,
      score,
      pageUrl,
      gameKey,
      durationSeconds,
      JSON.stringify(problems)
    ]);

    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput(String(err)).setMimeType(ContentService.MimeType.TEXT);
  }
}

/* ===== Helpers ===== */

function getOrCreateSheet_(ss, name) {
  var s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  return s;
}

function ensureHeader_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }
  var range = sheet.getRange(1, 1, 1, headers.length);
  var existing = range.getValues()[0];
  if (existing.join('|') !== headers.join('|')) {
    range.setValues([headers]);
  }
}

function extractGameKey_(url) {
  try {
    url = String(url || '');
    var match = url.match(/[?&]key=([^&#]+)/);
    return match ? decodeURIComponent(match[1]) : '';
  } catch (_){
    return '';
  }
}

function computeLocalFields_(isoTs) {
  var d = new Date(isoTs);
  var tz = Session.getScriptTimeZone();
  var localDate = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  var hour = Number(Utilities.formatDate(d, tz, 'H')); // 0-23
  var bucket = (hour<6) ? 'Night' : (hour<12) ? 'Morning' : (hour<18) ? 'Midday' : 'Evening';
  return { localDate: localDate, localHour: hour, bucket: bucket };
}

/* ===== Sitdowns ===== */
/* Sitdowns sheet: [User ID, Last Timestamp, Sitdown ID, Attempt #] */

function upsertSitdown_(ss, userId, isoTs) {
  var s = getOrCreateSheet_(ss, SITDOWN_SHEET_NAME);
  var headers = ['User ID','Last Timestamp','Sitdown ID','Attempt #'];
  ensureHeader_(s, headers);

  var row = findRowByKey_(s, 1, userId);
  var now = new Date(isoTs).getTime();

  if (row === -1) {
    var sidNew = userId + '-' + now;
    s.appendRow([userId, isoTs, sidNew, 1]);
    return { id: sidNew, attempt: 1 };
  }

  var vals = s.getRange(row, 1, 1, 4).getValues()[0];
  var lastIso = vals[1] || '';
  var sidOld  = String(vals[2] || '');
  var attOld  = Number(vals[3] || 0);

  var lastMs = lastIso ? new Date(lastIso).getTime() : 0;
  var gap = now - lastMs;
  var sid = sidOld;
  var attempt = attOld;

  if (!lastMs || gap > SITDOWN_IDLE_MS) {
    sid = userId + '-' + now; // new sitdown
    attempt = 1;
  } else {
    attempt = attOld + 1;     // same sitdown
  }

  s.getRange(row, 2, 1, 3).setValues([[isoTs, sid, attempt]]);
  return { id: sid, attempt: attempt };
}

function findRowByKey_(sheet, keyColIndex1Based, key) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var rng = sheet.getRange(2, keyColIndex1Based, last - 1, 1).getValues();
  for (var i = 0; i < rng.length; i++) {
    if (String(rng[i][0]) === String(key)) return i + 2;
  }
  return -1;
}

/* ===== GameModes ===== */
/* GameModes columns (you fill manually):
   'Game Key','Duration Seconds',
   'Add A Min','Add A Max','Add B Min','Add B Max',
   'Mul A Min','Mul A Max','Mul B Min','Mul B Max',
   'Op Addition','Op Subtraction','Op Multiplication','Op Division'
*/

function ensureGameModeRow_(ss, gameKey) {
  if (!gameKey) return;

  var headers = [
    'Game Key','Duration Seconds',
    'Add A Min','Add A Max','Add B Min','Add B Max',
    'Mul A Min','Mul A Max','Mul B Min','Mul B Max',
    'Op Addition','Op Subtraction','Op Multiplication','Op Division'
  ];

  var gm = getOrCreateSheet_(ss, GAMEMODES_SHEET_NAME);
  ensureHeader_(gm, headers);

  var last = gm.getLastRow();
  if (last > 1) {
    var keys = gm.getRange(2, 1, last - 1, 1).getValues().map(function(r){ return r[0]; });
    if (keys.indexOf(gameKey) !== -1) return;
  }

  gm.appendRow([
    gameKey, '',
    '','','','',
    '','','','',
    false, false, false, false
  ]);
}

/* ===== DailyStats (single row per date) ===== */
/* DailyStats columns:
   'Date','Sessions','Total Duration Seconds','Std*Dur Sum','Weighted Avg Std Score'
*/

function upsertDailyStats_(ss, localDate, durationSeconds, standardizedScore) {
  var sh = getOrCreateSheet_(ss, DAILYSTATS_SHEET_NAME);
  var headers = ['Date','Sessions','Total Duration Seconds','Std*Dur Sum','Weighted Avg Std Score'];
  ensureHeader_(sh, headers);

  // normalize display for matching
  sh.getRange(2, 1, Math.max(1, sh.getMaxRows()-1), 1).setNumberFormat('yyyy-mm-dd');

  var row = findRowByDisplayKey_(sh, 1, localDate);
  var addDur = (durationSeconds && durationSeconds > 0) ? durationSeconds : 0;
  var addStdDur = (addDur > 0 && standardizedScore !== '' && !isNaN(standardizedScore)) ? (standardizedScore * addDur) : 0;

  if (row === -1) {
    var sessionsCount = 1;
    var totalDur = addDur;
    var sumStdDur = addStdDur;
    var weightedAvg = (totalDur > 0) ? (sumStdDur / totalDur) : '';
    sh.appendRow([localDate, sessionsCount, totalDur, sumStdDur, weightedAvg]);
  } else {
    var vals = sh.getRange(row, 1, 1, headers.length).getValues()[0];
    var sessionsCount2 = Number(vals[1] || 0) + 1;
    var totalDur2 = Number(vals[2] || 0) + addDur;
    var sumStdDur2 = Number(vals[3] || 0) + addStdDur;
    var weightedAvg2 = (totalDur2 > 0) ? (sumStdDur2 / totalDur2) : '';
    sh.getRange(row, 2, 1, 4).setValues([[sessionsCount2, totalDur2, sumStdDur2, weightedAvg2]]);
  }
}

function findRowByDisplayKey_(sheet, keyColIndex1Based, key) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var disp = sheet.getRange(2, keyColIndex1Based, last - 1, 1).getDisplayValues();
  for (var i = 0; i < disp.length; i++) {
    if (String(disp[i][0]).trim() === String(key)) return i + 2;
  }
  return -1;
}
