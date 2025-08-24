/**
 * Dashboard.gs
 * Builds a no-formula Dashboard sheet from existing Sheets data.
 * Reads: Sheet1, Problems, DailyStats, ByOperator, TrendStdScore, HardestFacts, PacingThirds
 * Writes: Dashboard (KPIs, per-operator table, hardest facts, pacing/trend snapshots)
 *
 * Menu: Analytics > Rebuild Dashboard
 */

var SPREADSHEET_ID        = '1krmzTncZE4cTqTIn4HRJLXWKp6_9-sSuZx-t0ie2ZIQ';
var SESSIONS_SHEET_NAME   = 'Sheet1';
var PROBLEMS_SHEET_NAME   = 'Problems';
var DAILYSTATS_SHEET_NAME = 'DailyStats';
var BYOP_SHEET_NAME       = 'ByOperator';
var TREND_SHEET_NAME      = 'TrendStdScore';
var HARDEST_SHEET_NAME    = 'HardestFacts';
var PACING3_SHEET_NAME    = 'PacingThirds';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Analytics')
    .addItem('Rebuild Dashboard', 'rebuildDashboard')
    .addToUi();
}

function rebuildDashboard() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dash = getOrCreate_(ss, 'Dashboard');
  dash.clear();

  // Read sources (tolerant if missing)
  var sessions = readSheet1_(ss);
  var problems = readProblems_(ss);
  var daily    = readDaily_(ss);
  var byOp     = readTable_(ss, BYOP_SHEET_NAME);
  var trend    = readTable_(ss, TREND_SHEET_NAME);
  var hardest  = readTable_(ss, HARDEST_SHEET_NAME);
  var pacing3  = readTable_(ss, PACING3_SHEET_NAME);

  // Compute KPIs
  var totalSessions = sessions.length;
  var totalDur = sum_(sessions.map(function(s){ return num(s[10]); }));
  var totalProblems = problems.length;
  var avgStd = avgNonZero_(sessions.map(function(s){ return num(s[12]); }));
  var bestStd = max_(sessions.map(function(s){ return num(s[12]); }));
  var bestScore = max_(sessions.map(function(s){ return num(s[7]); }));
  var ppsOverall = totalDur > 0 ? totalProblems / totalDur : 0;

  var last7AvgStd = lastNAvg_(trend, 7, 2);    // column index of “Rolling 7-day Avg” on TrendStdScore = 2
  var last14AvgStd = lastNAvg_(trend, 14, 2);

  // Layout
  var r = 1;
  write2D_(dash, r, 1, [[ 'Dashboard (computed)', '' ]]); r += 2;

  write2D_(dash, r, 1, [[ 'Key metrics', '' ]]); r++;
  write2D_(dash, r, 1, [
    [ 'Total Sessions', totalSessions ],
    [ 'Total Duration (s)', totalDur ],
    [ 'Total Problems', totalProblems ],
    [ 'Problems / sec (overall)', ppsOverall ],
    [ 'Avg Std Score (all-time)', avgStd ],
    [ 'Best Std Score (session)', bestStd ],
    [ 'Best Raw Score (session)', bestScore ],
    [ 'Rolling 7-day Avg Std Score', last7AvgStd ],
    [ 'Rolling 14-day Avg Std Score', last14AvgStd ]
  ]); r += 10;

  write2D_(dash, r, 1, [[ 'By operator', '' ]]); r++;
  var opTable = [['Operator','Count','Share','Avg Latency Ms','Median Ms','p90 Ms','p95 Ms','Problems/sec','Std Score Contribution']];
  opTable = opTable.concat(byOp);
  write2D_(dash, r, 1, opTable); r += Math.max(2, byOp.length+2);

  write2D_(dash, r, 1, [[ 'Hardest facts (top N by avg latency)', '' ]]); r++;
  var hardestHeaders = [['Operator','A','B','Avg Latency Ms','Count']];
  var hardRows = (hardest.length && hardest[0][0] === 'Operator') ? hardest.slice(1) : hardest;
  write2D_(dash, r, 1, hardestHeaders.concat(hardRows)); r += Math.max(2, hardRows.length+2);

  write2D_(dash, r, 1, [[ 'Pacing (Thirds)', '' ]]); r++;
  var pacingTable = [['Third','Avg Ms','Median Ms','p90 Ms','p95 Ms']];
  var pacRows = (pacing3.length && pacing3[0][0] === 'Third') ? pacing3.slice(1) : pacing3;
  write2D_(dash, r, 1, pacingTable.concat(pacRows)); r += Math.max(2, pacRows.length+2);

  write2D_(dash, r, 1, [[ 'Trend (Std Score)', '' ]]); r++;
  var trendHeaders = [['Date','Avg Std Score','Rolling 7-day Avg']];
  var trRows = (trend.length && trend[0][0] === 'Date') ? trend.slice(1) : trend;
  write2D_(dash, r, 1, trendHeaders.concat(trRows)); r += Math.max(2, trRows.length+2);

  // Basic formatting
  dash.autoResizeColumns(1, 6);
  setHeaderStyle_(dash, 1, 1);
}

function getOrCreate_(ss, name) {
  var s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  return s;
}

function write2D_(sh, row, col, values2D) {
  if (!values2D || !values2D.length) return;
  sh.getRange(row, col, values2D.length, values2D[0].length).setValues(values2D);
}

function setHeaderStyle_(sh, startRow, startCol) {
  sh.getRange(startRow, startCol, 1, 2).setFontWeight('bold');
}

/* Readers for base sheets */

function readSheet1_(ss) {
  var sh = ss.getSheetByName(SESSIONS_SHEET_NAME);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last-1, sh.getLastColumn()).getValues();
}

function readProblems_(ss) {
  var sh = ss.getSheetByName(PROBLEMS_SHEET_NAME);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last-1, sh.getLastColumn()).getValues();
}

function readDaily_(ss) {
  var sh = ss.getSheetByName(DAILYSTATS_SHEET_NAME);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last-1, sh.getLastColumn()).getValues();
}

function readTable_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last-1, sh.getLastColumn()).getValues();
}

/* Helpers */

function num(v) { return (v === '' || v == null) ? 0 : Number(v); }

function sum_(arr) {
  var s = 0; for (var i = 0; i < arr.length; i++) s += num(arr[i]);
  return s;
}

function avgNonZero_(arr) {
  var vals = [];
  for (var i = 0; i < arr.length; i++) {
    var x = num(arr[i]); if (x) vals.push(x);
  }
  if (!vals.length) return 0;
  return sum_(vals) / vals.length;
}

function max_(arr) {
  var m = null;
  for (var i = 0; i < arr.length; i++) {
    var x = num(arr[i]);
    if (m == null || x > m) m = x;
  }
  return m == null ? 0 : m;
}

function lastNAvg_(rows, n, colIdx) {
  if (!rows || !rows.length) return 0;
  var vals = [];
  var take = Math.min(n, rows.length);
  for (var i = rows.length - take; i < rows.length; i++) {
    var v = num(rows[i][colIdx]);
    if (v) vals.push(v);
  }
  if (!vals.length) return 0;
  return sum_(vals) / vals.length;
}
