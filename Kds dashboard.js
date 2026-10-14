/*************************************************************************************************
 * KDS DASHBOARDS  (file 2 of 2)  -  isolated from the processing code
 * -----------------------------------------------------------------------------------------------
 * Two clearly separated jobs:
 *
 *   buildKdsDashboards()    FULL build: (re)creates each dashboard tab, pivot, slicers, colour
 *                           control, fonts, widths, frozen panes and conditional formatting.
 *                           Run this once, or whenever you want to reset the look.
 *
 *   refreshKdsDashboards()  DATA ONLY: rewrites the hidden snapshot tabs and re-points the pivots.
 *                           Formatting, titles, colours, column widths and slicers are left alone,
 *                           so it is fast and never resets your layout. This is what runs
 *                           automatically after processing.
 *
 * Colours are driven by FORMULAS that read the threshold cells (B2 green, B4 red) live - edit the
 * cells and the colours update instantly, no button needed.
 *************************************************************************************************/

var DASH_MAIN_ID = '1gMGz9E7IkEXGk8qwyYdAECgoZ1ckSphaYCutWOpG5mE';

var DASH_PAD_PERIOD     = true;   // true = pad to "P05W2" in the dashboard so P10-P13 sort after P2-P9 (needed now you have 13 periods)
var DASH_DAY_LABELLED   = true;   // true = days shown as "1 Mon".."7 Sun" so columns order Monday -> Sunday
var DASH_HIDE_DATA_TABS = true;   // hide the copied snapshot tabs
var DASH_DATA_FONT      = 11;     // base font size on the dashboards
var DASH_CHUNK = 5000;

var DASH_GREEN_BG  = '#B7E1CD';
var DASH_YELLOW_BG = '#FCE8B2';
var DASH_RED_BG    = '#F4C7C3';
var DASH_HEADER_BG = '#E8EAED';
// Separator rules between blocks of rows. Period total is the darkest (outermost), site total a
// step lighter, so the hierarchy still reads top-down while both act as hard dividers.
var DASH_TOTAL_BG      = '#5F6368';  // period / grand total
var DASH_SITE_TOTAL_BG = '#9AA0A6';  // site total — the line between one site and the next
var DASH_HEADER_FG = '#202124';   // dark header font so the day/shift labels are readable
var DASH_TITLE_FG  = '#1a73e8';

var DASH_DAY_IDX = { Monday: 1, Tuesday: 2, Wednesday: 3, Thursday: 4, Friday: 5, Saturday: 6, Sunday: 7 };

var DASH_DEFS = [
  {
    title: 'Passmaster Range Dashboard', oldTitles: ['Range Dashboard'], dataTab: '_RangeData', caption: 'Passmaster range % by day & shift (volume-weighted)',
    sourceId: '1IGxt4K9QqjpbkYci_aM-Bl0Kc81D_Of-gti8GqAtes4', sourceSheet: 'Range_Summary',
    headers: ['Location','Date','Shift','Period_Week','Period_Week_Order','Period_Week_Label','Day_Name','Day_Order','Shift_Order','Total_Orders','Over_5_Min_Orders','Red Range','Source File ID','Source File Name'],
    rows: ['Period_Week','Location'], cols: ['Day_Name','Shift'], value: 'Red Range', num: 'Over_5_Min_Orders', den: 'Total_Orders',
    slicers: ['Location','Period_Week'],
    percent: true, greenDef: 0.10, yellowDef: 0.20, redDef: 0.25
  },
  {
    title: 'Red Ticket Dashboard', dataTab: '_RedData', caption: 'Red ticket % by day & shift (volume-weighted)',
    sourceId: '1-n1ZJe9Y8Hl3Tvpu7D3LlaWMWcEnR73PVMCWT1pXM2Q', sourceSheet: 'Red_Ticket_Summary',
    headers: ['Location','Screen','Date','Shift','Period_Week','Period_Week_Order','Period_Week_Label','Day_Name','Day_Order','Shift_Order','Total_Tickets','Over_Tolerance_Tickets','Red ticket %','Source File ID','Source File Name'],
    rows: ['Period_Week','Location','Screen'], cols: ['Day_Name','Shift'], value: 'Red ticket %', num: 'Over_Tolerance_Tickets', den: 'Total_Tickets',
    slicers: ['Screen','Location','Date','Period_Week'],
    percent: true, greenDef: 0.10, yellowDef: 0.20, redDef: 0.30
  },
  {
    title: 'Times Dashboard', dataTab: '_TimesData', caption: 'Average duration (mins) by day & shift',
    sourceId: '10A4Kdwuguirbge13msNLm96NQte39wWUrpKGR3iN8PU', sourceSheet: 'Times_Summary',
    headers: ['Location','Screen','Date','Shift','Period_Week','Period_Week_Order','Period_Week_Label','Day_Name','Day_Order','Shift_Order','Average_Duration','Total_Tickets','Source File ID','Source File Name'],
    // Same shape as the other dashboards: Period_Week first in the rows, weekday across the top (Mon-Sun).
    rows: ['Period_Week','Location','Screen'], cols: ['Day_Name','Shift'], value: 'Average_Duration', noTotals: true,
    slicers: ['Location','Screen','Period_Week'],
    percent: false, bands: true, greenDef: 8, yellowDef: 10, redDef: 12, // thresholds in MINUTES
    numFmt: '0' // whole minutes, no decimal point
  }
];

/* ============================== ENTRY POINTS ============================== */

function buildKdsDashboards() {
  var out = DASH_DEFS.map(function (d) { return dbBuildOne_(d); }).filter(Boolean);
  var msg = 'Dashboards built: ' + (out.length ? out.join(', ') : 'none (no data?)');
  dbToast_(msg); dbLog_(msg);
}

function refreshKdsDashboards() {
  var ss = dbMain_();
  var out = DASH_DEFS.map(function (d) { return dbRefreshOne_(ss, d); }).filter(Boolean);
  var msg = 'Dashboard data refreshed: ' + (out.length ? out.join(', ') : 'none (no data?)');
  dbToast_(msg); dbLog_(msg);
}

// Re-establish the formula colour rules (rarely needed - colours already update live from the cells).
function applyDashboardColours() {
  var ss = dbMain_(), done = [];
  DASH_DEFS.forEach(function (def) {
    if (!def.percent && !def.bands) return;
    var sh = ss.getSheetByName(def.title);
    if (!sh) return;
    dbApplyBandsFormula_(sh, def);
    done.push(def.title);
  });
  var msg = done.length ? 'Colour rules refreshed: ' + done.join(', ') : 'No % dashboards found - build them first.';
  dbToast_(msg); dbLog_(msg);
}

/* ============================== FULL BUILD (structure + format) ============================== */

function dbBuildOne_(def) {
  try {
    var ss = dbMain_();
    var count = dbWriteSnapshot_(ss, def);
    if (!count) { dbLog_('No data in ' + def.sourceSheet + '; skipping ' + def.title + '.'); return null; }
    var ds = ss.getSheetByName(def.dataTab);

    // Fresh dashboard tab.
    (def.oldTitles || []).forEach(function (t) { var o = ss.getSheetByName(t); if (o) ss.deleteSheet(o); }); // remove stale tabs after a rename
    var old = ss.getSheetByName(def.title);
    if (old) ss.deleteSheet(old);
    var dash = ss.insertSheet(def.title);

    if (def.percent || def.bands) dbWriteColourControl_(dash, def);

    var src = ds.getRange(1, 1, ds.getLastRow(), def.headers.length);
    dbMakePivot_(dash, ds, def);
    dbAddSlicers_(dash, src, def);
    dbStyleFull_(dash, def);
    dbApplyValueFormat_(dash, def);
    if (def.percent || def.bands) dbApplyBandsFormula_(dash, def);
    dbSeparateTotals_(dash, def);

    dbLog_('Built ' + def.title + ' (' + count + ' rows).');
    return def.title;
  } catch (e) { dbLog_('Failed to build ' + def.title + ': ' + e.message); return null; }
}

/* ============================== DATA-ONLY REFRESH (fast) ============================== */

function dbRefreshOne_(ss, def) {
  try {
    var dash = ss.getSheetByName(def.title), ds = ss.getSheetByName(def.dataTab);
    if (!dash || !ds) return dbBuildOne_(def); // not built yet -> full build

    var count = dbWriteSnapshot_(ss, def);
    if (!count) { dbLog_('No data in ' + def.sourceSheet + '; skipping ' + def.title + '.'); return null; }
    ds = ss.getSheetByName(def.dataTab);
    var src = ds.getRange(1, 1, ds.getLastRow(), def.headers.length);

    // Re-point the pivot to the (possibly resized) data, keeping the tab's formatting/slicers.
    dbDropPivots_(dash);
    dbMakePivot_(dash, ds, def);
    dash.getSlicers().forEach(function (s) { try { s.setRange(src); } catch (e) {} });

    // Cheap, data-dependent bits only (so new rows pick up the look). No widths/freeze/title/slicer rebuild.
    dbStyleHeaderBand_(dash, def);
    dbApplyValueFormat_(dash, def);
    if (def.percent || def.bands) dbApplyBandsFormula_(dash, def);
    dbSeparateTotals_(dash, def);

    dbLog_('Refreshed ' + def.title + ' (' + count + ' rows).');
    return def.title;
  } catch (e) { dbLog_('Failed to refresh ' + def.title + ': ' + e.message); return null; }
}

/* ============================== SNAPSHOT (data tab sized exactly to the data) ============================== */

function dbWriteSnapshot_(ss, def) {
  var srcSS = SpreadsheetApp.openById(def.sourceId);
  var srcSheet = srcSS.getSheetByName(def.sourceSheet);
  if (!srcSheet || srcSheet.getLastRow() < 2) return 0;

  var nCols = def.headers.length;
  var values = srcSheet.getRange(1, 1, srcSheet.getLastRow(), nCols).getValues();
  var pwCol = def.headers.indexOf('Period_Week'), dayCol = def.headers.indexOf('Day_Name'), dateCol = def.headers.indexOf('Date');
  for (var i = 1; i < values.length; i++) {
    if (DASH_DAY_LABELLED && dayCol >= 0) values[i][dayCol] = dbDayLabel_(values[i][dayCol]);
    if (pwCol >= 0) {
      var lbl = DASH_PAD_PERIOD ? dbPadPeriod_(values[i][pwCol]) : String(values[i][pwCol]);
      var fy = (dateCol >= 0) ? dbFyOf_(values[i][dateCol]) : null; // year-aware label: "FY26 P08W1" sorts before "FY27 P01W1" and never merges across years
      values[i][pwCol] = (fy != null) ? ('FY' + fy + ' ' + lbl) : lbl;
    }
  }
  var nRows = values.length;

  var ds = ss.getSheetByName(def.dataTab) || ss.insertSheet(def.dataTab);
  var maxR = ds.getMaxRows();
  if (maxR > nRows) ds.deleteRows(nRows + 1, maxR - nRows);
  else if (maxR < nRows) ds.insertRowsAfter(maxR, nRows - maxR);
  if (ds.getMaxColumns() < nCols) ds.insertColumnsAfter(ds.getMaxColumns(), nCols - ds.getMaxColumns());

  ds.getRange(1, 1, nRows, nCols).clearContent();
  for (var r = 0; r < nRows; r += DASH_CHUNK) {
    var b = values.slice(r, r + DASH_CHUNK);
    ds.getRange(r + 1, 1, b.length, nCols).setValues(b);
  }
  if (DASH_HIDE_DATA_TABS) { try { ds.hideSheet(); } catch (e) {} }
  return nRows - 1;
}

/* ============================== PIVOT ============================== */

function dbMakePivot_(dash, ds, def) {
  var src = ds.getRange(1, 1, ds.getLastRow(), def.headers.length);
  var pivot = dash.getRange(8, 1).createPivotTable(src);
  def.rows.forEach(function (f) { pivot.addRowGroup(dbCol_(def, f)).sortAscending(); });
  def.cols.forEach(function (f, i) {
    var g = pivot.addColumnGroup(dbCol_(def, f));
    if (f === 'Shift' || f === 'Date') g.sortDescending(); else g.sortAscending(); // Date desc = newest first; Shift desc = Lunch before Dinner
    g.showTotals(def.noTotals ? false : i === 0); // Grand Total on outer group unless noTotals
  });
  if (def.num && def.den) {
    // Volume-weighted rate: SUM(numerator) / SUM(denominator) per cell (and in subtotals/grand total),
    // matching the web dashboard. A plain AVERAGE of the per-row % over-weights tiny cells.
    pivot.addCalculatedPivotValue(def.value, "='" + def.num + "' / '" + def.den + "'");
  } else {
    pivot.addPivotValue(dbCol_(def, def.value), SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE);
  }
  return pivot;
}

function dbDropPivots_(dash) {
  var pts = dash.getPivotTables();
  for (var i = 0; i < pts.length; i++) { try { pts[i].remove(); } catch (e) {} }
}

function dbAddSlicers_(dash, src, def) {
  // Work out the newest Period_Week so its slicer can open pre-filtered to the latest week.
  // Labels are year-aware ("FY26 P08W1"), so a plain text max is the chronological latest.
  var pwVals = [], latest = null, pwIdx = def.headers.indexOf('Period_Week');
  if (pwIdx >= 0 && src.getNumRows() > 1) {
    var colVals = src.offset(1, pwIdx, src.getNumRows() - 1, 1).getValues(), seen = {};
    colVals.forEach(function (r) { var v = r[0]; if (v !== '' && v != null && !seen[v]) { seen[v] = 1; pwVals.push(String(v)); if (latest === null || String(v) > latest) latest = String(v); } });
  }
  var col = 4;
  def.slicers.forEach(function (f) {
    try {
      var sl = dash.insertSlicer(src, 1, col);
      if (f === 'Period_Week' && latest && pwVals.length > 1) {
        var hidden = pwVals.filter(function (v) { return v !== latest; }); // show only the latest week; users can tick others
        sl.setColumnFilterCriteria(dbCol_(def, f), SpreadsheetApp.newFilterCriteria().setHiddenValues(hidden).build());
      } else {
        sl.setColumnFilterCriteria(dbCol_(def, f), SpreadsheetApp.newFilterCriteria().build());
      }
      try { sl.setTitle(f === 'Period_Week' ? 'Period_Week (opens on latest)' : f); } catch (e) {}
      col += 4;
    } catch (e) { dbLog_('Slicer ' + f + ' on ' + def.title + ': ' + e.message); }
  });
}

/* ============================== COLOUR CONTROL + FORMULA BANDS ============================== */

function dbWriteColourControl_(dash, def) {
  var unit = def.percent ? '' : ' (mins)';
  dash.getRange('A1').setValue('Colours' + unit).setFontWeight('bold');
  dash.getRange('A2').setValue('Green \u2264'); dash.getRange('B2').setValue(def.greenDef);
  dash.getRange('A3').setValue('Yellow');        dash.getRange('B3').setValue(def.yellowDef);
  dash.getRange('A4').setValue('Red \u2265');    dash.getRange('B4').setValue(def.redDef);
  dash.getRange('B2:B4').setNumberFormat(def.percent ? '0%' : (def.numFmt || '0.0'));
  dash.getRange('B2').setBackground(DASH_GREEN_BG);
  dash.getRange('B3').setBackground(DASH_YELLOW_BG);
  dash.getRange('B4').setBackground(DASH_RED_BG);
  dash.getRange('A1:A4').setFontSize(9);
  dash.getRange('A1:B4').setBorder(true, true, true, true, true, true, '#D0D0D0', SpreadsheetApp.BorderStyle.SOLID);
}

// Formula-driven colours: read $B$2 (green) and $B$4 (red) live. Editing the cells recolours instantly.
// The range starts BELOW the pivot header rows - date headers are numbers and would otherwise be coloured.
function dbApplyBandsFormula_(dash, def) {
  if (!def.percent && !def.bands) return;
  var dataStart = 8 + def.cols.length + 1, rowsLen = def.rows.length;
  var a = 'A' + dataStart;
  var area = dash.getRange(dataStart, 1, Math.max(dash.getMaxRows() - dataStart + 1, 1), dash.getMaxColumns());
  var rules = [];
  // Total-row separators FIRST (first match wins) — conditional formatting re-evaluates on every
  // slicer/collapse, so these bands follow the filters instead of drifting like static borders.
  // Top-level (period) total: "…Total" sits in column A. Deeper subtotals: column A blank, "Total" further right.
  // The site-total row is the only thing separating one site's screens from the next, and a pivot
  // can't hold a real blank row — so it's drawn as a solid dark rule instead of a faint tint.
  // Teams were reading a site's rows and the next site's rows as one block.
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=ISNUMBER(SEARCH("Total",$A' + dataStart + '))').setBackground(DASH_TOTAL_BG).setFontColor('#ffffff').setBold(true).setRanges([area]).build());
  if (rowsLen >= 2) rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A' + dataStart + '="",ISNUMBER(SEARCH("Total",$B' + dataStart + ')))').setBackground(DASH_SITE_TOTAL_BG).setFontColor('#ffffff').setBold(true).setRanges([area]).build());
  // Value colour bands.
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(' + a + '),' + a + '<=$B$2)').setBackground(DASH_GREEN_BG).setRanges([area]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(' + a + '),' + a + '>=$B$4)').setBackground(DASH_RED_BG).setRanges([area]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(' + a + '),' + a + '>$B$2,' + a + '<$B$4)').setBackground(DASH_YELLOW_BG).setRanges([area]).build());
  dash.setConditionalFormatRules(rules);
}

/* ============================== STYLING ============================== */

// Draw a separator line under every pivot subtotal/total row (pivots can't hold real blank rows).
function dbSeparateTotals_(dash, def) {
  // Separators are now conditional-format bands (see dbApplyBandsFormula_) so they follow slicers/collapses.
  // Static borders drifted when filtered, so we no longer draw them here.
  return;
}

function dbStyleFull_(dash, def) {
  var rowsLen = def.rows.length;
  var headerEnd = 8 + def.cols.length;

  try { dash.setHiddenGridlines(true); } catch (e) {}

  // Title + subtitle.
  var tr = dash.getRange(6, 1, 1, Math.max(rowsLen + 1, 3));
  try { tr.merge(); } catch (e) {}
  tr.setValue(def.title.replace(' Dashboard', '').toUpperCase())
    .setFontSize(15).setFontWeight('bold').setFontColor(DASH_TITLE_FG).setVerticalAlignment('middle');
  dash.getRange(7, 1).setValue(def.caption || '').setFontSize(10).setFontColor('#5f6368').setFontStyle('italic');

  dbStyleHeaderBand_(dash, def);

  // Column widths. Auto-fit reads nicely, but on very wide pivots (Date columns) it is slow and
  // can time out - so beyond 60 columns use a uniform narrow width for the value columns instead.
  var lastCol = Math.max(dash.getLastColumn(), rowsLen + 1);
  if (lastCol <= 60) {
    try { dash.autoResizeColumns(1, lastCol); } catch (e) {}
  } else {
    try { dash.autoResizeColumns(1, rowsLen); } catch (e) {}
    dash.setColumnWidths(rowsLen + 1, lastCol - rowsLen, 56);
  }
  dash.setColumnWidth(1, 120);

  // Frozen panes (row labels + the whole day/shift header).
  try { dash.setFrozenColumns(rowsLen); } catch (e) {}
  try { if (dash.getMaxRows() > headerEnd + 3) dash.setFrozenRows(headerEnd); } catch (e) {}
}

function dbStyleHeaderBand_(dash, def) {
  var rowsLen = def.rows.length, depth = def.cols.length;
  var lastCol = Math.max(dash.getLastColumn(), rowsLen + 1);
  dash.getRange(8, 1, depth + 1, lastCol)
      .setFontWeight('bold').setBackground(DASH_HEADER_BG).setFontColor(DASH_HEADER_FG)
      .setFontSize(DASH_DATA_FONT).setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.getRange(8, 1, depth + 1, rowsLen).setHorizontalAlignment('left');
}

function dbApplyValueFormat_(dash, def) {
  var depth = def.cols.length;
  var dataStart = 8 + depth + 1; // skip the pivot header rows (field row + one row per column group)
  var area = dash.getRange(dataStart, 1, Math.max(dash.getMaxRows() - dataStart + 1, 1), dash.getMaxColumns());
  area.setNumberFormat(def.numFmt || (def.percent ? '0.0%' : '0.0')); // numFmt overrides (e.g. '0' = no decimals)
  area.setFontSize(DASH_DATA_FONT);
  // If a column group is a real Date, format its header row so it shows m/d/yyyy, not a serial number.
  var dateLevel = def.cols.indexOf('Date');
  if (dateLevel >= 0) dash.getRange(9 + dateLevel, 1, 1, dash.getMaxColumns()).setNumberFormat('m/d/yyyy');
}

/* ============================== HELPERS ============================== */

function dbWriteColourControl_safe_(dash, def) { try { dbWriteColourControl_(dash, def); } catch (e) {} }

var DASH_FY_BASE_MS = new Date(2025, 5, 30).getTime(); // FY26 P1W1 = Mon 30 Jun 2025
function dbFyOf_(d) { var dt = (d instanceof Date) ? d : new Date(d); if (!dt || isNaN(dt.getTime())) return null; return 26 + Math.floor((dt.getTime() - DASH_FY_BASE_MS) / (364 * 86400000)); }
function dbPadPeriod_(v) {
  var s = String(v == null ? '' : v).trim();
  var m = s.match(/^P\s*(\d+)\s*W\s*(\d+)$/i);
  if (!m) return s;
  return 'P' + (parseInt(m[1], 10) < 10 ? '0' : '') + parseInt(m[1], 10) + 'W' + parseInt(m[2], 10);
}
function dbDayLabel_(v) {
  var s = String(v == null ? '' : v).trim();
  if (!DASH_DAY_IDX[s]) return s;
  return DASH_DAY_IDX[s] + ' ' + s.slice(0, 3); // "1 Mon" .. "7 Sun" - sorts Monday->Sunday, stays compact
}
function dbMain_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return (ss && ss.getId() === DASH_MAIN_ID) ? ss : SpreadsheetApp.openById(DASH_MAIN_ID);
}
function dbCol_(def, field) { return def.headers.indexOf(field) + 1; }
function dbLog_(message) {
  try { var sh = dbMain_().getSheetByName('Logs'); if (sh) sh.appendRow([new Date(), 'DASH', message]); } catch (e) {}
  Logger.log('DASH: ' + message);
}
function dbToast_(msg) { try { var ss = SpreadsheetApp.getActiveSpreadsheet(); if (ss) ss.toast(msg, 'KDS Dashboards', 6); } catch (e) {} }