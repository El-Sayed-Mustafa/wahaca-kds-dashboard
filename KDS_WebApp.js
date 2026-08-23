/**
 * KDS Performance Dashboard — Web App (Google Apps Script + HtmlService)
 * ---------------------------------------------------------------------
 * Speed model: getPeriodBundle() sends ONE period's data to the browser (cached
 * per period). Location/Screen filtering then happens client-side = instant,
 * no server calls. Only changing Period or hitting Refresh talks to the server.
 *
 * Heatmap: run buildHourlySummary() once (+ a daily trigger) to populate
 * Hourly_Summary. It reuses the SAME parsing/cleaning as the main pipeline.
 *
 * >>> VERIFY the CONFIG block against your sheets before deploying. <<<
 */

var CONFIG = {
  TIMES:  { id: '10A4Kdwuguirbge13msNLm96NQte39wWUrpKGR3iN8PU', sheet: 'Times_Summary' },
  RED:    { id: '1-n1ZJe9Y8Hl3Tvpu7D3LlaWMWcEnR73PVMCWT1pXM2Q', sheet: 'Red_Ticket_Summary' },
  RANGE:  { id: '1IGxt4K9QqjpbkYci_aM-Bl0Kc81D_Of-gti8GqAtes4', sheet: 'Range_Summary' },
  MANUAL: { id: '1gMGz9E7IkEXGk8qwyYdAECgoZ1ckSphaYCutWOpG5mE', sheet: 'Leaderboard_Manual' },
  HOURLY: { id: '1gMGz9E7IkEXGk8qwyYdAECgoZ1ckSphaYCutWOpG5mE', sheet: 'Hourly_Summary' },
  // Site performance status on the Overview — same red-ticket tolerance as the leaderboard bands
  // below, so a site can't read green on one page and red on the other.
  STATUS_GOOD_MAX: 0.10,  // Good  = under 10%
  STATUS_WATCH_MAX: 0.15, // Watch = 10–15%, Act = 15%+
  // Leaderboard colour bands — ABSOLUTE tolerances, not rank position. green = below .green,
  // amber = .green up to .amber, red = .amber and above. Change these to re-band the whole board.
  LB_BANDS: { red: { green: 0.10, amber: 0.15 }, pass: { green: 0.15, amber: 0.20 } },
  CACHE_TTL: 1800, // seconds cached period bundles live (30 min); Refresh forces a rebuild
  CACHE_VER: 'v15', // bump this whenever the bundle shape changes — invalidates all cached bundles
  WEIGHTS_WITH_DELIVEROO:    { del: 0.10, food: 0.50, red: 0.20, pass: 0.20 },
  WEIGHTS_WITHOUT_DELIVEROO: { del: 0.00, food: 0.50, red: 0.25, pass: 0.25 }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('KDS Performance Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Hourly_Summary can live in its own spreadsheet (keeps the main workbook light).
// Set once via setupHourlyFile(); falls back to CONFIG.HOURLY.id if not set.
function hourlyId_() { try { return PropertiesService.getScriptProperties().getProperty('HOURLY_SS_ID') || CONFIG.HOURLY.id; } catch (e) { return CONFIG.HOURLY.id; } }
function setupHourlyFile() {
  var ss = SpreadsheetApp.create('KDS Hourly Summary');
  PropertiesService.getScriptProperties().setProperty('HOURLY_SS_ID', ss.getId());
  buildHourlySummary(); // writes Hourly_Summary into the new file
  return { id: ss.getId(), url: ss.getUrl(), note: 'Hourly now lives here. You can delete the old Hourly_Summary tab from the main workbook.' };
}

// ---- cache + helpers -----------------------------------------------------
function cacheGet_(key) { try { var v = CacheService.getScriptCache().get(key); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
function cachePut_(key, obj) { try { CacheService.getScriptCache().put(key, JSON.stringify(obj), CONFIG.CACHE_TTL); } catch (e) {} }
function lbVersion_() { var p = PropertiesService.getScriptProperties(); return p.getProperty('LB_VER') || '0'; }
function bumpLbVersion_() { var p = PropertiesService.getScriptProperties(); p.setProperty('LB_VER', String((parseInt(lbVersion_(), 10) || 0) + 1)); }
// When someone edits a score tab directly in the sheet (not via the dashboard Save), invalidate the scores
// cache immediately so the leaderboard reflects it on the next load instead of waiting up to CACHE_TTL.
var SCORE_TABS_ = { 'Leaderboard_Scores': 1, 'Leaderboard_Period_Food': 1, 'Leaderboard_Period_Deliveroo': 1 };
function onScoreEdit(e) {
  try { var name = (e && e.range && e.range.getSheet) ? e.range.getSheet().getName() : (e && e.source ? e.source.getActiveSheet().getName() : ''); if (SCORE_TABS_[name]) bumpLbVersion_(); } catch (err) {}
}
// Run once to install the trigger (needed if this script isn't container-bound to the scores workbook).
function installScoreEditTrigger() {
  var ssId = CONFIG.MANUAL.id, existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) { if (existing[i].getHandlerFunction() === 'onScoreEdit') ScriptApp.deleteTrigger(existing[i]); }
  ScriptApp.newTrigger('onScoreEdit').forSpreadsheet(SpreadsheetApp.openById(ssId)).onEdit().create();
  return 'Score-edit trigger installed on the scores workbook.';
}

function readSheet_(id, sheetName) {
  var sh = SpreadsheetApp.openById(id).getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var out = [];
  for (var i = 1; i < values.length; i++) { var o = {}; for (var c = 0; c < headers.length; c++) o[headers[c]] = values[i][c]; out.push(o); }
  return out;
}
function norm_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function pick_(row, names) { for (var k in row) { var kn = norm_(k); for (var j = 0; j < names.length; j++) if (kn === norm_(names[j])) return row[k]; } return undefined; }
function num_(v) { if (v === '' || v == null) return null; var n = parseFloat(v); return isNaN(n) ? null : n; }
function toDate_(v) { if (v instanceof Date) return v; var d = new Date(v); return isNaN(d.getTime()) ? null : d; }
function wavg_(pairs) { var w = 0, s = 0; pairs.forEach(function (p) { if (p[0] != null && p[1] != null) { s += p[0] * p[1]; w += p[1]; } }); return w ? s / w : null; }

var C = {
  loc: ['Location'], screen: ['Screen'], date: ['Date'], pw: ['Period_Week'], pwOrd: ['Period_Week_Order'],
  dur: ['Average_Duration'], tk: ['Total_Tickets', 'Tickets'],
  red: ['Red ticket %', 'Red_Ticket_%'], range: ['Red Range', 'Red_Range'],
  redOver: ['Over_Tolerance_Tickets'], redTot: ['Total_Tickets'],
  rngOver: ['Over_5_Min_Orders', 'Over_8_Min_Orders', 'Over_Range_Orders', 'Over_Tolerance_Orders'], rngTot: ['Total_Orders'],
  dayName: ['Day_Name'], hour: ['Hour'], over: ['Over_Tolerance_Tickets', 'Over_5_Min_Orders']
};
var MON = { Monday: 0, Tuesday: 1, Wednesday: 2, Thursday: 3, Friday: 4, Saturday: 5, Sunday: 6 };

function distinctPeriods_(rows) {
  var seen = {}, list = [];
  rows.forEach(function (r) { var v = pick_(r, C.pw); if (v === '' || v == null) return; if (!(v in seen)) { seen[v] = true; list.push({ value: String(v), order: num_(pick_(r, C.pwOrd)) || 0 }); } });
  list.sort(function (a, b) { return a.order - b.order; });
  return list;
}
function filterRows_(rows, f) {
  return rows.filter(function (r) {
    if (f.period && String(pick_(r, C.pw)) !== String(f.period)) return false;
    if (f.location && f.location !== 'All') { var lv = pick_(r, C.loc); if (lv !== undefined && String(lv) !== String(f.location)) return false; }
    if (f.screen && f.screen !== 'All') { var sv = pick_(r, C.screen); if (sv !== undefined && String(sv) !== String(f.screen)) return false; }
    return true;
  });
}

// =========================================================================
//  ONE call per period -> browser filters Location/Screen instantly
// =========================================================================
// FY derived from date (Option B). FY26 P1W1 = Mon 30 Jun 2025; each FY = 52 weeks (364 days).
var FY_BASE_MS = new Date(2025, 5, 30).getTime();
function fyOf_(d) { var dt = toDate_(d); if (!dt) return null; return 26 + Math.floor((dt.getTime() - FY_BASE_MS) / (364 * 86400000)); }
function pwSortKey_(pw) { var m = /P0*(\d+)W0*(\d+)/i.exec(String(pw)); return m ? (parseInt(m[1], 10) * 10 + parseInt(m[2], 10)) : 9999; }
function distinctPeriodsFY_(rows) {
  var seen = {}, list = [];
  rows.forEach(function (r) { var pw = pick_(r, C.pw); if (pw == null || pw === '') return; var fy = fyOf_(pick_(r, C.date)); if (fy == null) return; var k = fy + '|' + pw; if (seen[k]) return; seen[k] = 1; var ord = num_(pick_(r, C.pwOrd)); if (ord == null || ord === 0) ord = pwSortKey_(pw); list.push({ fy: fy, pw: String(pw), order: ord, g: fy * 10000 + ord }); });
  list.sort(function (a, b) { return a.g - b.g; });
  return list;
}
function periodOfLbl_(pw) { return String(pw).split(/w/i)[0]; }
function quarterPeriods_(q) { var m = { Q1: ['P1', 'P2', 'P3'], Q2: ['P4', 'P5', 'P6', 'P7'], Q3: ['P8', 'P9', 'P10'], Q4: ['P11', 'P12', 'P13'] }; return m[String(q).toUpperCase()] || []; }

function getPeriodBundle(period, force) {
  var T = readSheet_(CONFIG.TIMES.id, CONFIG.TIMES.sheet);
  var periodsFY = distinctPeriodsFY_(T);
  if (!period) { var last = periodsFY[periodsFY.length - 1]; period = last ? (last.fy + '|' + last.pw) : ''; }
  var key = 'bundle|' + CONFIG.CACHE_VER + '|' + period;
  if (!force) { var c = cacheGet_(key); if (c) return c; }

  var R = readSheet_(CONFIG.RED.id, CONFIG.RED.sheet);
  var G = readSheet_(CONFIG.RANGE.id, CONFIG.RANGE.sheet);
  var H = readSheet_(hourlyId_(), CONFIG.HOURLY.sheet);
  var tz = Session.getScriptTimeZone();

  // parse key: "FY|pkey" (e.g. 26|P9W1, 26|period:P9, 26|__ALL__) or "__ALL__" (every FY)
  var fy = null, pkey = period;
  if (period === '__ALL__') { pkey = '__ALL__'; }
  else { var mm = /^(\d{2})\|(.*)$/.exec(String(period)); if (mm) { fy = +mm[1]; pkey = mm[2]; } }
  var mode, wholeLabel = null, weekLabel = null, qSet = null;
  if (pkey === '__ALL__') mode = 'all';
  else if (pkey.indexOf('quarter:') === 0) { mode = 'quarter'; qSet = {}; quarterPeriods_(pkey.slice(8)).forEach(function (p) { qSet[p] = 1; }); }
  else if (pkey.indexOf('period:') === 0) { mode = 'period'; wholeLabel = pkey.slice(7); }
  else { mode = 'week'; weekLabel = pkey; }
  function fyMatch(r) { return fy == null || fyOf_(pick_(r, C.date)) === fy; }
  var keepRow = (mode === 'week') ? function (r) { return String(pick_(r, C.pw)) === String(weekLabel) && fyMatch(r); }
    : (mode === 'period') ? function (r) { return periodOfLbl_(pick_(r, C.pw)) === wholeLabel && fyMatch(r); }
      : (mode === 'quarter') ? function (r) { return qSet[periodOfLbl_(pick_(r, C.pw))] && fyMatch(r); }
        : function (r) { return fyMatch(r); };

  var locIndex = {}, locations = [], scrIndex = {}, screens = [];
  function li(l) { if (l == null || l === '') return -1; if (!(l in locIndex)) { locIndex[l] = locations.length; locations.push(String(l)); } return locIndex[l]; }
  function si(s) { if (s == null || s === '') return -1; if (!(s in scrIndex)) { scrIndex[s] = screens.length; screens.push(String(s)); } return scrIndex[s]; }
  function iso(d) { var dt = toDate_(d); return dt ? Utilities.formatDate(dt, tz, 'yyyy-MM-dd') : ''; }
  function padPw(lbl) { var m = /P0*(\d+)W(\d+)/i.exec(String(lbl)); return m ? ('P' + ('0' + m[1]).slice(-2) + 'W' + m[2]) : String(lbl); }

  // per-date packing (single week)
  function packT(kr) { var o = []; T.forEach(function (r) { if (!kr(r)) return; o.push([li(pick_(r, C.loc)), si(pick_(r, C.screen)), iso(pick_(r, C.date)), num_(pick_(r, C.tk)) || 0, num_(pick_(r, C.dur))]); }); return o; }
  function packR(kr) { var o = []; R.forEach(function (r) { if (!kr(r)) return; o.push([li(pick_(r, C.loc)), si(pick_(r, C.screen)), iso(pick_(r, C.date)), num_(pick_(r, C.redOver)) || 0, num_(pick_(r, C.redTot)) || 0]); }); return o; }
  function packG(kr) { var o = []; G.forEach(function (r) { if (!kr(r)) return; o.push([li(pick_(r, C.loc)), iso(pick_(r, C.date)), num_(pick_(r, C.rngOver)) || 0, num_(pick_(r, C.rngTot)) || 0]); }); return o; }
  function packH(kr) { var o = []; H.forEach(function (r) { if (!kr(r)) return; var dn = pick_(r, C.dayName), di = MON[dn]; if (di == null) return; o.push([li(pick_(r, C.loc)), di, num_(pick_(r, C.hour)), num_(pick_(r, C.tk)) || 0, num_(pick_(r, C.over)) || 0]); }); return o; }
  // weekly-aggregated packing (whole period / whole FY / all)
  function packTAll(kr) { var g = {}; T.forEach(function (r) { if (!kr(r)) return; var l = li(pick_(r, C.loc)), s = si(pick_(r, C.screen)), pw = padPw(pick_(r, C.pw)); var tk = num_(pick_(r, C.tk)) || 0, du = num_(pick_(r, C.dur)); var k = l + '|' + s + '|' + pw; var a = g[k] || (g[k] = { l: l, s: s, pw: pw, tk: 0, dw: 0, dt: 0 }); a.tk += tk; if (du != null) { a.dw += du * tk; a.dt += tk; } }); return Object.keys(g).map(function (k) { var a = g[k]; return [a.l, a.s, a.pw, a.tk, a.dt ? +(a.dw / a.dt).toFixed(2) : null]; }); }
  function packRAll(kr) { var g = {}; R.forEach(function (r) { if (!kr(r)) return; var l = li(pick_(r, C.loc)), s = si(pick_(r, C.screen)), pw = padPw(pick_(r, C.pw)); var o = num_(pick_(r, C.redOver)) || 0, t = num_(pick_(r, C.redTot)) || 0; var k = l + '|' + s + '|' + pw; var a = g[k] || (g[k] = { l: l, s: s, pw: pw, o: 0, t: 0 }); a.o += o; a.t += t; }); return Object.keys(g).map(function (k) { var a = g[k]; return [a.l, a.s, a.pw, a.o, a.t]; }); }
  function packGAll(kr) { var g = {}; G.forEach(function (r) { if (!kr(r)) return; var l = li(pick_(r, C.loc)), pw = padPw(pick_(r, C.pw)); var o = num_(pick_(r, C.rngOver)) || 0, t = num_(pick_(r, C.rngTot)) || 0; var k = l + '|' + pw; var a = g[k] || (g[k] = { l: l, pw: pw, o: 0, t: 0 }); a.o += o; a.t += t; }); return Object.keys(g).map(function (k) { var a = g[k]; return [a.l, a.pw, a.o, a.t]; }); }
  function packHAll(kr) { var g = {}; H.forEach(function (r) { if (!kr(r)) return; var dn = pick_(r, C.dayName), di = MON[dn]; if (di == null) return; var l = li(pick_(r, C.loc)), hr = num_(pick_(r, C.hour)); var tot = num_(pick_(r, C.tk)) || 0, ov = num_(pick_(r, C.over)) || 0; var k = l + '|' + di + '|' + hr; var a = g[k] || (g[k] = { l: l, d: di, h: hr, t: 0, o: 0 }); a.t += tot; a.o += ov; }); return Object.keys(g).map(function (k) { var a = g[k]; return [a.l, a.d, a.h, a.t, a.o]; }); }

  var keepH = (mode === 'week') ? function (r) { return String(pick_(r, C.pw)) === String(weekLabel); }
    : (mode === 'period') ? function (r) { return periodOfLbl_(pick_(r, C.pw)) === wholeLabel; }
      : (mode === 'quarter') ? function (r) { return qSet[periodOfLbl_(pick_(r, C.pw))]; }
        : function () { return true; };

  var cur = (mode === 'week')
    ? { t: packT(keepRow), r: packR(keepRow), g: packG(keepRow), h: packH(keepH) }
    : { t: packTAll(keepRow), r: packRAll(keepRow), g: packGAll(keepRow), h: packHAll(keepH) };

  // previous period rows (crosses FY boundary via global order)
  var prevRows = null, prevKeyStr = null;
  if (mode === 'week') {
    var idx = -1; for (var i = 0; i < periodsFY.length; i++) if (periodsFY[i].fy === fy && periodsFY[i].pw === weekLabel) { idx = i; break; }
    if (idx > 0) { var pv = periodsFY[idx - 1], kp = function (r) { return String(pick_(r, C.pw)) === pv.pw && fyOf_(pick_(r, C.date)) === pv.fy; }; prevRows = { t: packT(kp), r: packR(kp), g: packG(kp) }; prevKeyStr = pv.fy + '|' + pv.pw; }
  } else if (mode === 'period') {
    var plist = [], pseen = {}; periodsFY.forEach(function (p) { var pl = periodOfLbl_(p.pw), k2 = p.fy + '|' + pl; if (!pseen[k2]) { pseen[k2] = 1; plist.push({ fy: p.fy, p: pl }); } });
    var pidx = -1; for (var j = 0; j < plist.length; j++) if (plist[j].fy === fy && plist[j].p === wholeLabel) { pidx = j; break; }
    if (pidx > 0) { var pp = plist[pidx - 1], kpp = function (r) { return periodOfLbl_(pick_(r, C.pw)) === pp.p && fyOf_(pick_(r, C.date)) === pp.fy; }; prevRows = { t: packTAll(kpp), r: packRAll(kpp), g: packGAll(kpp) }; prevKeyStr = pp.fy + '|period:' + pp.p; }
  }

  // per-site Food/Deliveroo scores for this selection + the previous one (for the KPI tiles)
  var curScores = getManualScores(period);
  var prevScores = prevKeyStr ? getManualScores(prevKeyStr) : null;

  var fys = []; periodsFY.forEach(function (p) { if (fys.indexOf(p.fy) < 0) fys.push(p.fy); });
  var out = {
    updated: Utilities.formatDate(new Date(), tz, 'd MMM yyyy, HH:mm'),
    period: period, fys: fys, periodsFY: periodsFY.map(function (p) { return { fy: p.fy, pw: p.pw }; }),
    locations: locations, screens: screens,
    thresholds: { good: CONFIG.STATUS_GOOD_MAX, watch: CONFIG.STATUS_WATCH_MAX, bands: CONFIG.LB_BANDS },
    cur: cur, prev: prevRows, scores: { cur: curScores, prev: prevScores }, heatmapReady: H.length > 0
  };
  cachePut_(key, out);
  return out;
}

// server-side aggregate (used only for the previous-period KPI baseline)
function aggregate_(T, R, G) {
  var orders = 0, durW = [];
  T.forEach(function (r) { var tk = num_(pick_(r, C.tk)) || 0, dur = num_(pick_(r, C.dur)); orders += tk; durW.push([dur, tk]); });
  var prep = wavg_(durW);
  var ro = 0, rt = 0; R.forEach(function (r) { ro += num_(pick_(r, C.redOver)) || 0; rt += num_(pick_(r, C.redTot)) || 0; });
  var go = 0, gt = 0; G.forEach(function (r) { go += num_(pick_(r, C.rngOver)) || 0; gt += num_(pick_(r, C.rngTot)) || 0; });
  var delayedPct = rt ? ro / rt : null, rangePct = gt ? go / gt : null;
  return { orders: orders, prep: prep == null ? null : +prep.toFixed(1), onTime: delayedPct == null ? null : 1 - delayedPct, delayedPct: delayedPct, delayedCount: delayedPct == null ? null : Math.round(delayedPct * orders), rangePct: rangePct };
}
function groupWeighted_(rows, valueNames) {
  var acc = {}, allW = [];
  rows.forEach(function (r) { var v = num_(pick_(r, valueNames)); if (v == null) return; var tk = num_(pick_(r, C.tk)); if (tk == null) tk = 1; var lc = pick_(r, C.loc) || '__none'; acc[lc] = acc[lc] || []; acc[lc].push([v, tk]); allW.push([v, tk]); });
  var out = { __all: wavg_(allW) }; Object.keys(acc).forEach(function (k) { out[k] = wavg_(acc[k]); }); return out;
}

// ---- leaderboard scores live ONLY in 'Leaderboard_Scores' (per Period_Week) --
function sitesFromTimes_() {
  var s = {}; readSheet_(CONFIG.TIMES.id, CONFIG.TIMES.sheet).forEach(function (r) { var l = pick_(r, C.loc); if (l) s[l] = true; });
  return Object.keys(s).sort();
}
// ---- historical scores (per Period_Week) in 'Leaderboard_Scores' ----------
function normLc_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function parseYesNo_(v) { var s = String(v == null ? '' : v).toLowerCase(); return (s === 'yes' || s === 'true' || s === 'y' || s === '1'); }
function metricOf_(h) { var s = String(h == null ? '' : h).toLowerCase(); if (s.indexOf('food') >= 0) return 'food'; if (s.indexOf('del') >= 0) return 'del'; return null; }
function periodOfPw_(pw) { return String(pw).split(/w/i)[0]; }

function ensureScoresTab_() {
  var ss = SpreadsheetApp.openById(CONFIG.MANUAL.id);
  var sh = ss.getSheetByName('Leaderboard_Scores');
  if (!sh) sh = ss.insertSheet('Leaderboard_Scores');
  return sh;
}
// Reads BOTH the new horizontal matrix and the old vertical layout -> flat list
function readScores_() {
  var sh = ensureScoresTab_();
  if (sh.getLastRow() < 1) return [];
  var values = sh.getDataRange().getValues();
  if (!values.length) return [];
  var horizontal = values.length > 1 && String(values[1][0]).trim() === 'Site';
  var out = [];
  if (!horizontal) { // old vertical: Period_Week, Site, On_Deliveroo, Deliveroo, Food
    if (normLc_(values[0][0]) !== 'periodweek') return [];
    for (var r = 1; r < values.length; r++) { var row = values[r], pw = row[0], site = row[1]; if (!site || pw == null || pw === '') continue; out.push({ pw: String(pw), site: String(site), fy: '*', onDeliveroo: parseYesNo_(row[2]), deliveroo: num_(row[3]), food: num_(row[4]) }); }
    return out;
  }
  var h1 = values[1], hasFY = normLc_(h1[1]) === 'fy', onCol = hasFY ? 2 : 1, weekStart = hasFY ? 3 : 2;
  var h0 = values[0], colMap = [], curPw = null;
  for (var c = weekStart; c < h1.length; c++) { if (h0[c] != null && String(h0[c]).trim()) curPw = String(h0[c]).trim(); var m = metricOf_(h1[c]); if (curPw && m) colMap[c] = { pw: curPw, metric: m }; }
  var rec = {};
  for (var rr = 2; rr < values.length; rr++) {
    var row2 = values[rr], site2 = row2[0]; if (!site2) continue;
    var fyKey = hasFY ? (String(row2[1] || '').replace(/[^0-9]/g, '') || '*') : '*', onDel = parseYesNo_(row2[onCol]);
    for (var cc = weekStart; cc < row2.length; cc++) { var cm = colMap[cc]; if (!cm) continue; var v = num_(row2[cc]); if (v == null) continue; var k = fyKey + '|' + site2 + '|' + cm.pw; var a = rec[k] || (rec[k] = { fy: fyKey, pw: cm.pw, site: String(site2), onDeliveroo: onDel, deliveroo: null, food: null }); a.onDeliveroo = onDel; if (cm.metric === 'food') a.food = v; else a.deliveroo = v; }
  }
  Object.keys(rec).forEach(function (k) { out.push(rec[k]); });
  return out;
}

// Build/refresh the weekly matrix: Site | FY | On Deliveroo | [P1W1 Food/Del ...] — preserves existing values
function seedScoresTab() {
  var existing = readScores_(), fys = fysFromTimes_();
  var byKey = {}, onByKey = {};
  existing.forEach(function (r) {
    var fk = (r.fy && r.fy !== '*') ? r.fy : String(fys[0]); // migrate legacy (no-FY) into earliest FY
    byKey[fk + '|' + normSite_(r.site) + '|' + r.pw] = { del: r.deliveroo, food: r.food };
    var ok = fk + '|' + normSite_(r.site); if (r.onDeliveroo) onByKey[ok] = true; else if (!(ok in onByKey)) onByKey[ok] = false;
  });
  var sites = sitesFromTimes_();
  existing.forEach(function (r) { if (sites.map(normSite_).indexOf(normSite_(r.site)) < 0) sites.push(r.site); });
  sites.sort();
  var pws = distinctPeriods_(readSheet_(CONFIG.TIMES.id, CONFIG.TIMES.sheet)).map(function (p) { return p.value; });
  var pwSeen = {}; pws.forEach(function (p) { pwSeen[String(p).toUpperCase()] = 1; });
  existing.forEach(function (r) { if (!pwSeen[r.pw.toUpperCase()]) { pws.push(r.pw); pwSeen[r.pw.toUpperCase()] = 1; } });
  pws.sort(function (a, b) { return pwSortKey_(a) - pwSortKey_(b); }); // P1W1..P13W4 regardless of stored order / FY

  var nCols = 3 + pws.length * 2;
  function blank() { var a = []; for (var i = 0; i < nCols; i++) a.push(''); return a; }
  var row0 = blank(), row1 = blank(); row1[0] = 'Site'; row1[1] = 'FY'; row1[2] = 'On Deliveroo';
  pws.forEach(function (pw, i) { var c = 3 + i * 2; row0[c] = pw; row1[c] = 'Food'; row1[c + 1] = 'Del'; });
  var data = [row0, row1];
  fys.forEach(function (fy) {
    sites.forEach(function (s) {
      var row = blank(); row[0] = s; row[1] = 'FY' + fy; row[2] = (onByKey[fy + '|' + normSite_(s)] === false ? 'No' : 'Yes');
      pws.forEach(function (pw, i) { var c = 3 + i * 2, e = byKey[fy + '|' + normSite_(s) + '|' + pw]; if (e) { row[c] = (e.food == null ? '' : e.food); row[c + 1] = (e.del == null ? '' : e.del); } });
      data.push(row);
    });
  });

  var ss = SpreadsheetApp.openById(CONFIG.MANUAL.id);
  var sh = ss.getSheetByName('Leaderboard_Scores') || ss.insertSheet('Leaderboard_Scores');
  sh.clear(); try { sh.setConditionalFormatRules([]); } catch (e) {}
  try { sh.getRange(1, 1, sh.getMaxRows(), Math.max(sh.getMaxColumns(), 1)).clearDataValidations(); } catch (e) {} // old Yes/No rule sat on a now-different column
  sh.getRange(1, 1, data.length, nCols).setValues(data);
  applyScoresFormatting_(sh, pws, data.length - 1, nCols);
  bumpLbVersion_();
  return { sites: sites.length, fys: fys, periodWeeks: pws.length };
}
function applyScoresFormatting_(sh, pws, nRows, nCols) {
  var lastRow = 2 + nRows;
  sh.setFrozenRows(2); sh.setFrozenColumns(3);
  pws.forEach(function (pw, i) { try { sh.getRange(1, 4 + i * 2, 1, 2).merge(); } catch (e) {} });
  sh.getRange(1, 1, 1, nCols).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(1, 1, 1, 3).setBackground('#0b57d0');
  sh.getRange(2, 1, 1, nCols).setBackground('#e8f0fe').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, 1, 1, 2).setHorizontalAlignment('left');
  sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 60); sh.setColumnWidth(3, 100);
  for (var c = 4; c <= nCols; c++) sh.setColumnWidth(c, 58);
  if (lastRow >= 2) sh.getRange(1, 1, lastRow, nCols).setBorder(true, true, true, true, true, true, '#d0d7de', SpreadsheetApp.BorderStyle.SOLID);
  if (nRows > 0) {
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build();
    sh.getRange(3, 3, nRows, 1).setDataValidation(rule);
    var foodRanges = [], delRanges = [];
    pws.forEach(function (pw, i) { foodRanges.push(sh.getRange(3, 4 + i * 2, nRows, 1)); delRanges.push(sh.getRange(3, 5 + i * 2, nRows, 1)); });
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(7).setBackground('#d9ead3').setRanges(foodRanges).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(7).setBackground('#f4cccc').setRanges(foodRanges).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(4.65).setBackground('#d9ead3').setRanges(delRanges).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(4.65).setBackground('#f4cccc').setRanges(delRanges).build()
    ]);
  }
}
function fysFromTimes_() { var T = readSheet_(CONFIG.TIMES.id, CONFIG.TIMES.sheet), s = {}; T.forEach(function (r) { var fy = fyOf_(pick_(r, C.date)); if (fy != null) s[fy] = 1; }); var a = Object.keys(s).map(Number).sort(function (x, y) { return x - y; }); return a.length ? a : [26]; }
// Run ONCE to (re)build the period score tabs as: Site | FY | P1..P13 (one row per site per fiscal year).
function seedPeriodTabs() {
  var ss = SpreadsheetApp.openById(CONFIG.MANUAL.id);
  var sites = sitesFromTimes_(), fys = fysFromTimes_();
  var periods = []; for (var p = 1; p <= 13; p++) periods.push('P' + p);
  ['Leaderboard_Period_Food', 'Leaderboard_Period_Deliveroo'].forEach(function (name) {
    var existing = readPeriodMatrix_(name); // { ns -> { fyKey('26'|'*') -> { Px: v } } }
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clear(); try { sh.setConditionalFormatRules([]); } catch (e) {}
    try { sh.getRange(1, 1, sh.getMaxRows(), Math.max(sh.getMaxColumns(), 1)).clearDataValidations(); } catch (e) {}
    var header = ['Site', 'FY'].concat(periods), data = [header];
    fys.forEach(function (fy) {
      sites.forEach(function (s) {
        var ns = normSite_(s), byFy = existing[ns] || {};
        var rec = byFy[String(fy)] || (fy === fys[0] ? byFy['*'] : null) || {}; // migrate old (no-FY) rows into the earliest FY
        var row = [s, 'FY' + fy]; periods.forEach(function (pl) { var e = rec[pl]; row.push(e == null ? '' : e); });
        data.push(row);
      });
    });
    sh.getRange(1, 1, data.length, header.length).setValues(data);
    sh.setFrozenRows(1); sh.setFrozenColumns(2);
    sh.getRange(1, 1, 1, header.length).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(1, 1, 1, 2).setHorizontalAlignment('left');
    sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 60); for (var cw = 3; cw <= header.length; cw++) sh.setColumnWidth(cw, 58);
    var nRows = data.length - 1;
    if (nRows) {
      var rng = [sh.getRange(2, 3, nRows, periods.length)];
      rng[0].setBorder(true, true, true, true, true, true, '#d0d7de', SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment('center');
      var thr = (name.indexOf('Food') >= 0) ? 7 : 4.65;
      sh.setConditionalFormatRules([
        SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(thr).setBackground('#d9ead3').setRanges(rng).build(),
        SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(thr).setBackground('#f4cccc').setRanges(rng).build()
      ]);
    }
  });
  bumpLbVersion_();
  return { sites: sites.length, fys: fys, tabs: ['Leaderboard_Period_Food', 'Leaderboard_Period_Deliveroo'] };
}
function normSite_(s) { return String(s == null ? '' : s).toLowerCase().replace(/wahaca/g, '').replace(/[^a-z0-9]/g, ''); }
// Reads Site | [FY] | P1..P13 -> { normSite -> { fyKey -> { 'P9': v } } }. fyKey is '26'/'27' or '*' (no FY col). 0=blank.
function readPeriodMatrix_(sheetName) {
  var sh = SpreadsheetApp.openById(CONFIG.MANUAL.id).getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return {};
  var values = sh.getDataRange().getValues(), header = values[0];
  var hasFY = normLc_(header[1]) === 'fy', start = hasFY ? 2 : 1, cols = [];
  for (var c = start; c < header.length; c++) { var m = /P0*(\d+)/i.exec(String(header[c]).trim()); if (m) cols[c] = 'P' + m[1]; }
  var map = {};
  for (var r = 1; r < values.length; r++) {
    var row = values[r], site = row[0]; if (!site) continue; var ns = normSite_(site);
    var fyKey = hasFY ? (String(row[1] || '').replace(/[^0-9]/g, '') || '*') : '*';
    var byFy = map[ns] || (map[ns] = {}), rec = byFy[fyKey] || (byFy[fyKey] = {});
    for (var cc = start; cc < row.length; cc++) { if (!cols[cc]) continue; var v = num_(row[cc]); if (v == null || v === 0) continue; rec[cols[cc]] = v; }
  }
  return map;
}
// Period / all scores from the period tabs (FY-aware). Returns null if the tabs are empty (caller falls back to weekly).
function periodScoresForKey_(fy, pkey) {
  var food = readPeriodMatrix_('Leaderboard_Period_Food'), del = readPeriodMatrix_('Leaderboard_Period_Deliveroo');
  if (!Object.keys(food).length && !Object.keys(del).length) return null;
  var isPeriod = String(pkey).indexOf('period:') === 0, wholeLabel = isPeriod ? ('P' + String(pkey).slice(7).replace(/^P/i, '')) : null;
  var isQuarter = String(pkey).indexOf('quarter:') === 0, qPeriods = isQuarter ? quarterPeriods_(pkey.slice(8)) : null;
  function pickFy(byFy) {
    if (!byFy) return {};
    if (fy != null) return byFy[String(fy)] || byFy['*'] || {};
    var merged = {}; Object.keys(byFy).forEach(function (fk) { var rec = byFy[fk]; Object.keys(rec).forEach(function (p) { (merged[p] = merged[p] || []).push(rec[p]); }); });
    var out = {}; Object.keys(merged).forEach(function (p) { var a = merged[p]; out[p] = a.reduce(function (x, y) { return x + y; }, 0) / a.length; }); return out;
  }
  function avgOver(rec, periods) { var s = 0, n = 0; periods.forEach(function (p) { if (rec[p] != null) { s += rec[p]; n++; } }); return n ? s / n : null; }
  return sitesFromTimes_().map(function (s) {
    var ns = normSite_(s), fr = pickFy(food[ns]), dr = pickFy(del[ns]), f, d;
    if (isPeriod) { f = fr[wholeLabel]; d = dr[wholeLabel]; }
    else if (isQuarter) { f = avgOver(fr, qPeriods); d = avgOver(dr, qPeriods); }
    else { var fs = 0, fn = 0, ds = 0, dn = 0; Object.keys(fr).forEach(function (p) { fs += fr[p]; fn++; }); Object.keys(dr).forEach(function (p) { ds += dr[p]; dn++; }); f = fn ? fs / fn : null; d = dn ? ds / dn : null; }
    return { site: s, food: (f == null ? null : +f.toFixed(1)), deliveroo: (d == null ? null : +d.toFixed(2)), onDeliveroo: (d != null) };
  });
}
function weeklyScoresForKey_(fy, key) {
  var rows = readScores_();
  var isWeek = /^P\d+W\d+$/i.test(String(key));
  var isPeriod = String(key).indexOf('period:') === 0, wholeLabel = isPeriod ? String(key).slice(7) : null;
  var acc = {};
  rows.forEach(function (r) {
    if (fy != null && r.fy !== '*' && +r.fy !== fy) return; // FY filter ('*' = legacy, always allowed)
    var include = isWeek ? (r.pw.toUpperCase() === String(key).toUpperCase())
      : isPeriod ? (periodOfPw_(r.pw).toUpperCase() === String(wholeLabel).toUpperCase()) : true;
    if (!include) return;
    var ns = normSite_(r.site), a = acc[ns] || (acc[ns] = { onDel: false, dS: 0, dN: 0, fS: 0, fN: 0 });
    if (r.onDeliveroo) a.onDel = true;
    if (r.deliveroo != null) { a.dS += r.deliveroo; a.dN++; }
    if (r.food != null) { a.fS += r.food; a.fN++; }
  });
  return sitesFromTimes_().map(function (s) { var a = acc[normSite_(s)]; return { site: s, onDeliveroo: a ? a.onDel : true, deliveroo: (a && a.dN) ? +(a.dS / a.dN).toFixed(2) : null, food: (a && a.fN) ? +(a.fS / a.fN).toFixed(1) : null }; });
}
function getManualScores(key) {
  key = key || '';
  var vkey = 'scores|' + lbVersion_() + '|' + key;
  var c = cacheGet_(vkey); if (c) return c;
  var fy = null, pkey = key;
  if (key === '__ALL__') { pkey = '__ALL__'; }
  else { var m = /^(\d{2})\|(.*)$/.exec(String(key)); if (m) { fy = +m[1]; pkey = m[2]; } }
  var isWeek = /^P\d+W\d+$/i.test(String(pkey));
  var out = null;
  if (!isWeek) out = periodScoresForKey_(fy, pkey); // period / whole-FY / all-years -> period tabs
  if (!out) out = weeklyScoresForKey_(fy, pkey);     // single week, or fallback if period tabs empty
  cachePut_(vkey, out);
  return out;
}
function saveManualScores(list, fullKey) {
  fullKey = String(fullKey || '');
  var fy = null, weekKey = fullKey, m = /^(\d{2})\|(.*)$/.exec(fullKey); if (m) { fy = +m[1]; weekKey = m[2]; }
  if (!/^P\d+W\d+$/i.test(weekKey)) return { ok: false, errors: ['Pick a single week to save scores (whole-period / all-periods views are read-only).'] };
  var ss = SpreadsheetApp.openById(CONFIG.MANUAL.id);
  var sh = ss.getSheetByName('Leaderboard_Scores');
  var hasFYNow = sh && sh.getLastRow() > 1 && String(sh.getRange(2, 1).getValue()).trim() === 'Site' && normLc_(String(sh.getRange(2, 2).getValue())) === 'fy';
  if (!hasFYNow) { seedScoresTab(); sh = ss.getSheetByName('Leaderboard_Scores'); }
  var values = sh.getDataRange().getValues(), h0 = values[0], h1 = values[1];
  var foodCol = -1, delCol = -1, curPw = null;
  for (var c = 3; c < h1.length; c++) { if (h0[c] != null && String(h0[c]).trim()) curPw = String(h0[c]).trim(); if (curPw && curPw.toUpperCase() === weekKey.toUpperCase()) { var mm = metricOf_(h1[c]); if (mm === 'food') foodCol = c; else if (mm === 'del') delCol = c; } }
  if (foodCol < 0) { var nc = values[0].length; sh.getRange(1, nc + 1).setValue(weekKey); sh.getRange(1, nc + 1, 1, 2).merge(); sh.getRange(2, nc + 1).setValue('Food'); sh.getRange(2, nc + 2).setValue('Del'); foodCol = nc; delCol = nc + 1; }
  var fyNum = (fy == null ? String(fysFromTimes_().slice(-1)[0]) : String(fy)), fyLabel = 'FY' + fyNum;
  var rowOf = {}; for (var r = 2; r < values.length; r++) { var s0 = values[r][0], f0 = String(values[r][1] || '').replace(/[^0-9]/g, ''); if (s0) rowOf[s0 + '|' + f0] = r + 1; }
  var errors = [];
  (list || []).forEach(function (item) {
    var site = item.site; if (!site) return;
    var del = item.deliveroo === '' || item.deliveroo == null ? null : parseFloat(item.deliveroo);
    var food = item.food === '' || item.food == null ? null : parseFloat(item.food);
    if (del != null && (isNaN(del) || del < 1 || del > 5)) { errors.push(site + ': Deliveroo 1-5'); return; }
    if (food != null && (isNaN(food) || food < 1 || food > 10)) { errors.push(site + ': Food 1-10'); return; }
    var rowNum = rowOf[site + '|' + fyNum];
    if (!rowNum) { rowNum = sh.getLastRow() + 1; sh.getRange(rowNum, 1).setValue(site); sh.getRange(rowNum, 2).setValue(fyLabel); rowOf[site + '|' + fyNum] = rowNum; }
    sh.getRange(rowNum, 3).setValue(item.onDeliveroo ? 'Yes' : 'No');
    sh.getRange(rowNum, foodCol + 1).setValue(food == null ? '' : +food.toFixed(1));
    sh.getRange(rowNum, delCol + 1).setValue(del == null ? '' : +del.toFixed(2));
  });
  bumpLbVersion_();
  return { ok: errors.length === 0, errors: errors };
}
function rankColumn_(items, valueFn, higherIsBetter) {
  var arr = items.map(function (it) { return { id: it.site, v: valueFn(it) }; }).filter(function (x) { return x.v != null && !isNaN(x.v); });
  arr.sort(function (a, b) { return higherIsBetter ? b.v - a.v : a.v - b.v; });
  var ranks = {}, lastV = null, lastRank = 0;
  for (var i = 0; i < arr.length; i++) { var rank; if (lastV !== null && arr[i].v === lastV) rank = lastRank; else { rank = i + 1; lastRank = rank; lastV = arr[i].v; } ranks[arr[i].id] = rank; }
  return ranks;
}

/**
 * buildHourlySummary — populates Hourly_Summary for the heatmap.
 * Reuses the SAME read-only helpers as the main pipeline. Delay = duration >
 * tolerance-for-screen (identical to the Red-ticket rule). Run once, then add a
 * daily trigger. Keep the cleaning below in sync with processCsvFile_.
 */
function buildHourlySummary() {
  var settings = getSettings_(), periodMap = buildPeriodMap_(), stationMap = getStationMap_();
  var folder = DriveApp.getFolderById(RAW_FOLDER_ID), it = folder.getFiles();
  var acc = {}, t0 = (typeof EXEC_T0 !== 'undefined' ? EXEC_T0 : Date.now()), files = 0, stopped = false;
  while (it.hasNext()) {
    if (Date.now() - t0 > MAX_RUNTIME_MS) { stopped = true; break; }
    var file = it.next(); if (!isCsvFile_(file)) continue;
    var text = file.getBlob().getDataAsString('UTF-8'); if (text && text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    var data = Utilities.parseCsv(text); if (!data || data.length < 2) continue;
    var idx = buildHeaderIndex_(data[0]);
    var cLoc = getCol_(idx, 'Location'), cStation = getCol_(idx, 'Station'), cFired = getCol_(idx, 'Fired Date'), cFulf = getCol_(idx, 'Fulfilled Date');
    var cTable = getCol_(idx, 'Table');
    if (cLoc < 0 || cStation < 0 || cFired < 0 || cFulf < 0) continue;
    for (var r = 1; r < data.length; r++) {
      var row = data[r]; if (!row || !row.length) continue;
      var location = (row[cLoc] || '').toString().trim(); if (!location) continue;
      if (location.toLowerCase() === 'wahaca [office]') continue;
      // Same Excluded_Tables filter as processCsvFile_ — the staff table sits open all day.
      if (cTable >= 0 && isExcludedTable_(row[cTable], settings.excludedTables)) continue;
      var stationRaw = (row[cStation] || '').toString().trim();
      if (stationRaw.toLowerCase() === 'no print') continue;
      if (stationRaw && /^[-\u2013\u2014\s]+$/.test(stationRaw)) continue;
      var fired = parseRawDateTime_(row[cFired]); if (!fired) continue;
      if (minutesOfDay_(fired) < settings.minStartMinutes) continue;
      var fulfilled = parseRawDateTime_(row[cFulf]); if (!fulfilled) continue;
      var screen = mapStation_(stationRaw, stationMap); if (!screen) continue;
      var duration = (fulfilled.getTime() - fired.getTime()) / 60000;
      var dateOnly = new Date(fired.getFullYear(), fired.getMonth(), fired.getDate());
      var pm = periodMap[dateKey_(dateOnly)] || MISSING_PW;
      var dayName = DAY_NAMES[dateOnly.getDay()], dayOrder = DAY_ORDER[dayName], hour = fired.getHours();
      var tol = settings.toleranceByScreen[screen]; if (tol == null) tol = settings.toleranceByScreen.Mains;
      var key = location + '|' + pm.pw + '|' + dayOrder + '|' + hour;
      var a = acc[key] || (acc[key] = { loc: location, pw: pm.pw, pwOrder: pm.order, dayName: dayName, dayOrder: dayOrder, hour: hour, total: 0, over: 0 });
      a.total++; if (duration > tol) a.over++;
    }
    files++;
  }
  var ss = SpreadsheetApp.openById(hourlyId_());
  var sh = ss.getSheetByName(CONFIG.HOURLY.sheet) || ss.insertSheet(CONFIG.HOURLY.sheet);
  var hadData = sh.getLastRow() > 1;
  // If we ran out of time mid-pass AND a good summary already exists, keep it rather than overwrite with a partial one.
  if (stopped && hadData) { return { files: files, rows: null, stopped: true, complete: false }; }
  sh.clear();
  var header = ['Location', 'Period_Week', 'Period_Week_Order', 'Day_Name', 'Day_Order', 'Hour', 'Total_Tickets', 'Over_Tolerance_Tickets'];
  var rows = [header];
  Object.keys(acc).forEach(function (k) { var a = acc[k]; rows.push([a.loc, a.pw, a.pwOrder, a.dayName, a.dayOrder, a.hour, a.total, a.over]); });
  sh.getRange(1, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
  return { files: files, rows: rows.length - 1, stopped: stopped, complete: !stopped };
}