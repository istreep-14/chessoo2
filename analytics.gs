/**
 * Analytics.gs (new file)
 * Run-on-demand analytics WITHOUT spreadsheet formulas.
 *
 * Provides a custom menu: Analytics > Recompute All Analytics
 * Reads from:
 *  - Sheet1 (sessions)
 *  - Problems (per-problem rows)
 *  - DailyStats (one row per date)
 * Writes to:
 *  - ByOperator
 *  - ByOperandRanges
 *  - HeatmapMul
 *  - HeatmapDiv
 *  - HardestFacts
 *  - PacingThirds
 *  - PacingDeciles
 *  - TrendStdScore
 *  - Consistency
 *  - Throughput
 *  - SessionAggregates
 *  - ByGameKey
 *  - WeeklyStats
 *
 * Assumed schemas:
 * Sheet1 headers:
 *   Timestamp, Local Date, Local Hour, Time-of-Day, User ID, Sitdown ID, Attempt #,
 *   Score, Page URL, Game Key, Duration Seconds, Score/Second, Standardized Score,
 *   Problems JSON, Problem Count
 *
 * Problems headers:
 *   Timestamp, User ID, Game Key, Duration Seconds,
 *   Problem #, Operator, A, B, Latency Ms,
 *   Cum Ms, Third, Decile,
 *   Correct Answer, Final Answer, Wrong Full-Length Attempts
 *
 * DailyStats headers:
 *   Date, Sessions, Total Duration Seconds, Std*Dur Sum, Weighted Avg Std Score
 */

var SPREADSHEET_ID        = '1krmzTncZE4cTqTIn4HRJLXWKp6_9-sSuZx-t0ie2ZIQ';
var SESSIONS_SHEET_NAME   = 'Sheet1';
var PROBLEMS_SHEET_NAME   = 'Problems';
var DAILYSTATS_SHEET_NAME = 'DailyStats';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Analytics')
    .addItem('Recompute All Analytics', 'recomputeAllAnalytics')
    .addToUi();
}

function recomputeAllAnalytics() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sessions = readSessions_(ss);
  var problems = readProblems_(ss);
  var daily = readDaily_(ss);

  recomputeByOperator_(ss, sessions, problems);
  recomputeOperandBuckets_(ss, problems);
  recomputeHeatmaps_(ss, problems);
  recomputeHardestFacts_(ss, problems, 20);
  recomputePacing_(ss, problems);
  recomputeTrendStdScore_(ss, sessions);
  recomputeConsistency_(ss, problems);
  recomputeThroughput_(ss, sessions, problems);
  recomputeSessionAggregates_(ss, sessions, problems);
  recomputeByGameKey_(ss, sessions);
  recomputeWeeklyStats_(ss, daily);
}

/* ========== Readers ========== */

function readSessions_(ss) {
  var sh = ss.getSheetByName(SESSIONS_SHEET_NAME);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var width = sh.getLastColumn();
  var values = sh.getRange(2, 1, last - 1, width).getValues();
  var res = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    res.push({
      timestamp: String(r[0]),
      localDate: String(r[1] || ''),
      localHour: Number(r[2] || 0),
      timeOfDay: String(r[3] || ''),
      userId: String(r[4] || ''),
      sitdownId: String(r[5] || ''),
      attempt: Number(r[6] || 0),
      score: Number(r[7] || 0),
      pageUrl: String(r[8] || ''),
      gameKey: String(r[9] || ''),
      duration: Number(r[10] || 0),
      scorePerSec: Number(r[11] || 0),
      stdScore: Number(r[12] || 0),
      problemCount: Number(r[14] || 0)
    });
  }
  return res;
}

function readProblems_(ss) {
  var sh = ss.getSheetByName(PROBLEMS_SHEET_NAME);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var width = sh.getLastColumn();
  var values = sh.getRange(2, 1, last - 1, width).getValues();
  var res = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    res.push({
      timestamp: String(r[0]),
      userId: String(r[1] || ''),
      gameKey: String(r[2] || ''),
      duration: Number(r[3] || 0),
      index: Number(r[4] || 0),
      op: String(r[5] || ''),
      A: Number(r[6] || 0),
      B: Number(r[7] || 0),
      latency: Number(r[8] || 0),
      cumMs: Number(r[9] || 0),
      third: Number(r[10] || 0),
      decile: Number(r[11] || 0),
      correct: r[12],
      final: r[13],
      wrongFull: (r.length > 14 && r[14] !== '' ? Number(r[14]) : null)
    });
  }
  return res;
}

function readDaily_(ss) {
  var sh = ss.getSheetByName(DAILYSTATS_SHEET_NAME);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, 5).getValues();
  var res = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    res.push({
      date: String(r[0]),
      sessions: Number(r[1] || 0),
      totalDur: Number(r[2] || 0),
      stdDurSum: Number(r[3] || 0),
      weightedAvg: Number(r[4] || 0)
    });
  }
  return res;
}

/* ========== Writers ========== */

function writeTable_(ss, sheetName, headers, rows) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  sh.clear();
  if (!headers || !headers.length) return;
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

/* ========== Stats helpers ========== */

function avg_(arr) {
  if (!arr.length) return 0;
  var s = 0; for (var i = 0; i < arr.length; i++) s += arr[i];
  return s / arr.length;
}
function stddev_(arr) {
  if (arr.length < 2) return 0;
  var m = avg_(arr);
  var s = 0; for (var i = 0; i < arr.length; i++) { var d = arr[i] - m; s += d * d; }
  return Math.sqrt(s / (arr.length - 1));
}
function median_(arr) { return percentileInc_(arr, 0.5); }
function percentileInc_(arr, p) {
  if (!arr.length) return 0;
  var a = arr.slice().sort(function(x,y){return x-y;});
  var n = a.length;
  var rank = 1 + (n - 1) * p;
  if (rank <= 1) return a[0];
  if (rank >= n) return a[n-1];
  var lo = Math.floor(rank) - 1;
  var hi = Math.ceil(rank) - 1;
  var f = rank - Math.floor(rank);
  return a[lo] + f * (a[hi] - a[lo]);
}
function mad_(arr) {
  if (!arr.length) return 0;
  var m = median_(arr);
  var dev = [];
  for (var i = 0; i < arr.length; i++) dev.push(Math.abs(arr[i] - m));
  return median_(dev);
}

/* ========== Analytics builders ========== */

function recomputeByOperator_(ss, sessions, problems) {
  var ops = ['+','-','*','/'];
  var totalProblems = problems.length;

  // Sum durations per operator across unique sessions in which that operator appears
  var sessionDurByOp = {'+':0,'-':0,'*':0,'/':0};
  var sessionDurById = {}; // timestamp -> duration
  for (var s = 0; s < sessions.length; s++) {
    sessionDurById[sessions[s].timestamp] = Number(sessions[s].duration || 0);
  }
  var opToSessionSet = {'+':{},'-':{},'*':{},'/':{}};
  for (var i = 0; i < problems.length; i++) {
    var p = problems[i];
    if (!opToSessionSet[p.op]) opToSessionSet[p.op] = {};
    opToSessionSet[p.op][p.timestamp] = true;
  }
  for (var o = 0; o < ops.length; o++) {
    var op = ops[o];
    var set = opToSessionSet[op] || {};
    var sumDur = 0;
    for (var ts in set) { sumDur += (sessionDurById[ts] || 0); }
    sessionDurByOp[op] = sumDur;
  }

  var groups = {'+':[], '-':[], '*':[], '/':[]};
  for (var k = 0; k < problems.length; k++) {
    var pr = problems[k];
    if (!groups[pr.op]) groups[pr.op] = [];
    groups[pr.op].push(pr.latency);
  }

  var rows = [];
  for (var j = 0; j < ops.length; j++) {
    var op = ops[j];
    var lat = groups[op] || [];
    var count = lat.length;
    var share = totalProblems ? (count / totalProblems) : 0;
    var avg = avg_(lat);
    var med = median_(lat);
    var p90 = percentileInc_(lat, 0.9);
    var p95 = percentileInc_(lat, 0.95);
    var dur = sessionDurByOp[op] || 0;
    var pps = dur > 0 ? (count / dur) : 0; // problems/sec
    var stdContrib = pps * 120;

    rows.push([op, count, share, avg, med, p90, p95, pps, stdContrib]);
  }

  writeTable_(ss, 'ByOperator',
    ['Operator','Count','Share','Avg Latency Ms','Median Ms','p90 Ms','p95 Ms','Problems/sec','Std Score Contribution'],
    rows
  );
}

function bucket_(x) {
  if (x === '' || x == null) return '';
  if (x <= 10) return '2–10';
  if (x <= 20) return '11–20';
  if (x <= 50) return '21–50';
  if (x <= 100) return '51–100';
  return '>100';
}

function recomputeOperandBuckets_(ss, problems) {
  // Build per-operator A and B bucket stats
  var map = {}; // op|A_bucket|B_bucket -> {count, latencies[]}
  for (var i = 0; i < problems.length; i++) {
    var p = problems[i];
    var ab = bucket_(p.A);
    var bb = bucket_(p.B);
    var key = [p.op, ab, bb].join('|');
    if (!map[key]) map[key] = {count:0, lat:[]};
    map[key].count++;
    map[key].lat.push(p.latency);
  }
  var rows = [];
  for (var key in map) {
    var parts = key.split('|');
    var op = parts[0], aB = parts[1], bB = parts[2];
    var lat = map[key].lat;
    rows.push([op, aB, bB, map[key].count, avg_(lat), median_(lat), percentileInc_(lat,0.9), percentileInc_(lat,0.95)]);
  }
  writeTable_(ss, 'ByOperandRanges',
    ['Operator','A Bucket','B Bucket','Count','Avg Ms','Median Ms','p90 Ms','p95 Ms'],
    rows
  );
}

function recomputeHeatmaps_(ss, problems) {
  var ops = {'*':'HeatmapMul','/':'HeatmapDiv'};
  var buckets = ['2–10','11–20','21–50','51–100','>100'];

  for (var op in ops) {
    var grid = {};
    for (var r = 0; r < buckets.length; r++) {
      grid[buckets[r]] = {};
      for (var c = 0; c < buckets.length; c++) {
        grid[buckets[r]][buckets[c]] = [];
      }
    }
    for (var i = 0; i < problems.length; i++) {
      var p = problems[i];
      if (p.op !== op) continue;
      var ab = bucket_(p.A), bb = bucket_(p.B);
      if (!ab || !bb) continue;
      grid[ab][bb].push(p.latency);
    }
    var headers = ['A\\B'].concat(buckets);
    var rows = [];
    for (var r2 = 0; r2 < buckets.length; r2++) {
      var aB = buckets[r2];
      var row = [aB];
      for (var c2 = 0; c2 < buckets.length; c2++) {
        var bB = buckets[c2];
        var lat = grid[aB][bB];
        row.push(lat.length ? avg_(lat) : '');
      }
      rows.push(row);
    }
    writeTable_(ss, ops[op], headers, rows);
  }
}

function recomputeHardestFacts_(ss, problems, topN) {
  var map = {}; // op|A|B -> latencies
  for (var i = 0; i < problems.length; i++) {
    var p = problems[i];
    var key = [p.op,p.A,p.B].join('|');
    if (!map[key]) map[key] = [];
    map[key].push(p.latency);
  }
  var items = [];
  for (var k in map) {
    var parts = k.split('|');
    var lat = map[k];
    items.push({op:parts[0], A:Number(parts[1]), B:Number(parts[2]), avg:avg_(lat), count:lat.length});
  }
  items.sort(function(a,b){
    if (b.avg !== a.avg) return b.avg - a.avg;
    return b.count - a.count;
  });
  var rows = [];
  var n = Math.min(topN || 20, items.length);
  for (var i2 = 0; i2 < n; i2++) {
    rows.push([items[i2].op, items[i2].A, items[i2].B, items[i2].avg, items[i2].count]);
  }
  writeTable_(ss, 'HardestFacts', ['Operator','A','B','Avg Latency Ms','Count'], rows);
}

function recomputePacing_(ss, problems) {
  // Thirds
  var byT = {}; // third -> latencies
  for (var i = 0; i < problems.length; i++) {
    var p = problems[i];
    if (!p.third) continue;
    if (!byT[p.third]) byT[p.third] = [];
    byT[p.third].push(p.latency);
  }
  var rowsT = [];
  for (var k in byT) {
    var lat = byT[k];
    rowsT.push([Number(k), avg_(lat), median_(lat), percentileInc_(lat,0.9), percentileInc_(lat,0.95)]);
  }
  rowsT.sort(function(a,b){return a[0]-b[0];});
  writeTable_(ss, 'PacingThirds', ['Third','Avg Ms','Median Ms','p90 Ms','p95 Ms'], rowsT);

  // Deciles
  var byD = {}; // decile -> latencies
  for (var j = 0; j < problems.length; j++) {
    var p2 = problems[j];
    if (!p2.decile) continue;
    if (!byD[p2.decile]) byD[p2.decile] = [];
    byD[p2.decile].push(p2.latency);
  }
  var rowsD = [];
  for (var k2 in byD) {
    var lat2 = byD[k2];
    rowsD.push([Number(k2), avg_(lat2), median_(lat2), percentileInc_(lat2,0.9), percentileInc_(lat2,0.95)]);
  }
  rowsD.sort(function(a,b){return a[0]-b[0];});
  writeTable_(ss, 'PacingDeciles', ['Decile','Avg Ms','Median Ms','p90 Ms','p95 Ms'], rowsD);
}

function recomputeTrendStdScore_(ss, sessions) {
  // Group by Local Date average standardized score
  var map = {}; // date -> [stdScores...]
  var dates = [];
  for (var i = 0; i < sessions.length; i++) {
    var s = sessions[i];
    if (!s.localDate) continue;
    if (!map[s.localDate]) { map[s.localDate] = []; dates.push(s.localDate); }
    map[s.localDate].push(Number(s.stdScore || 0));
  }
  dates.sort();
  var rows = [];
  var window = []; // last 7 days values
  for (var d = 0; d < dates.length; d++) {
    var day = dates[d];
    var vals = map[day];
    var avgDay = vals.length ? avg_(vals) : 0;

    window.push(avgDay);
    if (window.length > 7) window.shift();
    var roll7 = avg_(window);

    rows.push([day, avgDay, roll7]);
  }
  writeTable_(ss, 'TrendStdScore', ['Date','Avg Std Score','Rolling 7-day Avg'], rows);
}

function recomputeConsistency_(ss, problems) {
  var ops = ['+','-','*','/'];
  var rows = [];
  for (var i = 0; i < ops.length; i++) {
    var op = ops[i];
    var lat = [];
    for (var j = 0; j < problems.length; j++) if (problems[j].op === op) lat.push(problems[j].latency);
    var avgL = avg_(lat);
    var sd = stddev_(lat);
    var madv = mad_(lat);
    var cov = avgL ? (sd / avgL) : 0;
    rows.push([op, avgL, sd, madv, cov]);
  }
  writeTable_(ss, 'Consistency', ['Operator','Avg Ms','StdDev Ms','MAD Ms','CoV'], rows);
}

function recomputeThroughput_(ss, sessions, problems) {
  // Overall problems/sec
  var totalProblems = problems.length;
  var totalDur = 0;
  var used = {};
  for (var s = 0; s < sessions.length; s++) {
    if (used[sessions[s].timestamp]) continue;
    used[sessions[s].timestamp] = true;
    totalDur += Number(sessions[s].duration || 0);
  }
  var overall = totalDur > 0 ? (totalProblems / totalDur) : 0;

  // Per operator
  var ops = ['+','-','*','/'];
  var sessionDurById = {};
  for (var t = 0; t < sessions.length; t++) sessionDurById[sessions[t].timestamp] = Number(sessions[t].duration || 0);
  var opToSession = {'+':{},'-':{},'*':{},'/':{}};
  var countByOp = {'+':0,'-':0,'*':0,'/':0};
  for (var i = 0; i < problems.length; i++) {
    var p = problems[i];
    countByOp[p.op] = (countByOp[p.op] || 0) + 1;
    opToSession[p.op][p.timestamp] = true;
  }
  var rows = [];
  rows.push(['OVERALL','', totalProblems, totalDur, overall, overall*120]);
  for (var o = 0; o < ops.length; o++) {
    var op = ops[o];
    var dur = 0; var set = opToSession[op] || {};
    for (var ts in set) dur += (sessionDurById[ts] || 0);
    var pps = dur > 0 ? (countByOp[op] / dur) : 0;
    rows.push([op, countByOp[op], countByOp[op], dur, pps, pps*120]);
  }
  writeTable_(ss, 'Throughput', ['Scope/Operator','Count','Problems','Total Duration (s)','Problems/sec','Std Score Contribution'], rows);
}

function recomputeSessionAggregates_(ss, sessions, problems) {
  // Build per-session latency stats from Problems
  var map = {}; // ts -> latencies[]
  for (var i = 0; i < problems.length; i++) {
    var p = problems[i];
    if (!map[p.timestamp]) map[p.timestamp] = [];
    map[p.timestamp].push(p.latency);
  }

  var rows = [];
  for (var s = 0; s < sessions.length; s++) {
    var x = sessions[s];
    var lat = map[x.timestamp] || [];
    var avgL = avg_(lat);
    var p90 = percentileInc_(lat, 0.9);
    var p95 = percentileInc_(lat, 0.95);
    rows.push([
      x.timestamp, x.userId, x.gameKey, x.duration, x.score, x.scorePerSec, x.stdScore,
      avgL, p90, p95, lat.length
    ]);
  }

  writeTable_(ss, 'SessionAggregates',
    ['Timestamp','User ID','Game Key','Duration (s)','Score','Score/sec','Std Score','Avg Ms','p90 Ms','p95 Ms','Problem Count'],
    rows
  );
}

function recomputeByGameKey_(ss, sessions) {
  var map = {}; // gameKey -> {sessions,count, sumSps, sumStd, n}
  for (var i = 0; i < sessions.length; i++) {
    var s = sessions[i];
    if (!map[s.gameKey]) map[s.gameKey] = {sessions:0, sumSps:0, sumStd:0, n:0};
    map[s.gameKey].sessions += 1;
    map[s.gameKey].sumSps += Number(s.scorePerSec || 0);
    map[s.gameKey].sumStd += Number(s.stdScore || 0);
    map[s.gameKey].n += 1;
  }
  var rows = [];
  for (var k in map) {
    var g = map[k];
    rows.push([k, g.sessions, g.n ? (g.sumSps/g.n) : 0, g.n ? (g.sumStd/g.n) : 0]);
  }
  writeTable_(ss, 'ByGameKey', ['Game Key','Sessions','Avg Score/sec','Avg Std Score'], rows);
}

function recomputeWeeklyStats_(ss, daily) {
  // weekly by yyyy-ww, sum durations, sum stdDurSum, sessions, then weighted avg
  var map = {}; // week -> {sessions,totalDur,stdDurSum}
  for (var i = 0; i < daily.length; i++) {
    var d = daily[i];
    var w = toIsoWeekKey_(d.date);
    if (!map[w]) map[w] = {sessions:0,totalDur:0,stdDurSum:0};
    map[w].sessions += Number(d.sessions || 0);
    map[w].totalDur += Number(d.totalDur || 0);
    map[w].stdDurSum += Number(d.stdDurSum || 0);
  }
  var weeks = Object.keys(map).sort();
  var rows = [];
  for (var j = 0; j < weeks.length; j++) {
    var wk = weeks[j], v = map[wk];
    var weighted = v.totalDur > 0 ? (v.stdDurSum / v.totalDur) : 0;
    rows.push([wk, v.sessions, v.totalDur, v.stdDurSum, weighted]);
  }
  writeTable_(ss, 'WeeklyStats',
    ['Week','Sessions','Total Duration Seconds','Std*Dur Sum','Weighted Avg Std Score'],
    rows
  );
}

function toIsoWeekKey_(dateStr) {
  // dateStr 'yyyy-mm-dd' -> 'yyyy-ww'
  var parts = String(dateStr).split('-');
  if (parts.length !== 3) return dateStr;
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  var dayNr = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - dayNr + 3);
  var firstThursday = new Date(d.getFullYear(),0,4);
  dayNr = (firstThursday.getDay() + 6) % 7;
  firstThursday.setDate(firstThursday.getDate() - dayNr + 3);
  var week = 1 + Math.round((d - firstThursday) / (7*24*3600*1000));
  var y = d.getFullYear();
  return y + '-' + ('0'+week).slice(-2);
}
