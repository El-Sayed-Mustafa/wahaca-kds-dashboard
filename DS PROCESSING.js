/*************************************************************************************************
 * KDS PROCESSING  (file 1 of 2)
 * -----------------------------------------------------------------------------------------------
 * Reads Toast/KDS CSVs from Drive, processes in memory, writes SUMMARY tables to the external
 * summary spreadsheets. Dashboards live in a SEPARATE file (KDS_Dashboards.gs) and read snapshots
 * into the main spreadsheet.
 *
 * RUN ORDER
 *  1. Setup Workbook            (seeds Settings, StationMap, PeriodWeek; creates summary sheets)
 *  2. Rebuild All               (reprocess everything with the current PeriodWeek calendar)
 *  3. Start Catch-Up            (finish all files)
 *  4. Build Dashboards          (defined in KDS_Dashboards.gs)
 *************************************************************************************************/

/* ============================== CONFIG ============================== */

var RAW_FOLDER_ID   = '1wlNULT0ZGHxpAOEbDDC8hzLax5IItcfo';
var SPREADSHEET_ID  = '1gMGz9E7IkEXGk8qwyYdAECgoZ1ckSphaYCutWOpG5mE';

var OUTPUT_SPREADSHEETS = {
  'Red_Ticket_Summary': '1-n1ZJe9Y8Hl3Tvpu7D3LlaWMWcEnR73PVMCWT1pXM2Q',
  'Times_Summary':      '10A4Kdwuguirbge13msNLm96NQte39wWUrpKGR3iN8PU',
  'Range_Summary':      '1IGxt4K9QqjpbkYci_aM-Bl0Kc81D_Of-gti8GqAtes4',
  'Range_Check_Detail': '',
  'Unmapped_Stations':  ''
};

var PERIOD_LABEL_INCLUDE_PW = true;
var ENABLE_RANGE_DETAIL = false;
var MAX_RUNTIME_MS = 4.5 * 60 * 1000; // 4.5 min per run. Consumer Gmail hard-caps at 6 min; leave headroom so the run
// STOPS ITSELF and schedules a continuation instead of being force-killed. (On Google Workspace you can raise this to ~25.)
var CHAIN_DELAY_MS = 60 * 1000;
var MONITOR_INTERVAL_MIN = 15;
var WRITE_CHUNK = 5000;
var AUTO_BUILD_DASHBOARDS = true; // rebuild dashboards automatically once a refresh finishes and data actually changed

var CHAIN_HANDLER   = 'scheduledRefresh';
var MONITOR_HANDLER = 'monitorRefresh';

var MISSING_PW = { pw: 'Missing Period Mapping', order: 999999, label: 'Missing Period Mapping' };
var _SS_CACHE = {};
var EXEC_T0 = Date.now(); // true start of THIS execution (module loads once per run) - runtime budget is measured from here so setup time in rebuildAll counts too

/* ============================== STATION MAP DEFAULTS ============================== */

var DEFAULT_STATION_MAP = [
  ['Wimb - Bar', 'Bar'], ['WCity -  Bar', 'Bar'], ['Waterloo  - Bar', 'Bar'], ['Stratford - Bar', 'Bar'],
  ['Shor - Bar', 'Bar'], ['SB Bar - Up', 'Bar'], ['SB - Bar Down', 'Bar'], ['Padd - Bar', 'Bar'],
  ['Isl - Bar', 'Bar'], ['GPS-Bar', 'Bar'], ['Edi Bar', 'Bar'], ['CDF - Bar', 'Bar'],
  ['C Wharf - Bar', 'Bar'], ['Bright - Bar', 'Bar'], ['*Covent G - Bar', 'Bar'],
  ['Wimb - Mains', 'Mains'], ['Wimb - Del Mains', 'Mains'], ['WCity Del Mains', 'Mains'], ['WCity - Mains', 'Mains'],
  ['Waterloo Del Mains', 'Mains'], ['Waterloo - Mains', 'Mains'], ['Stratford - Mains', 'Mains'], ['Strat Del Mains', 'Mains'],
  ['Shor - Mains', 'Mains'], ['Shor - Del Mains', 'Mains'], ['SB - Mains', 'Mains'], ['Padd - Mains', 'Mains'],
  ['Isl Del - Mains', 'Mains'], ['Isl - Mains', 'Mains'], ['GPS-Mains', 'Mains'], ['GPS - Del Mains', 'Mains'],
  ['Edi Del - Mains', 'Mains'], ['Edi - Mains', 'Mains'], ['CDF - Mains', 'Mains'], ['Cardiff Del Mains', 'Mains'],
  ['C Wharf Del - Mains', 'Mains'], ['C Wharf - Mains', 'Mains'], ['Bright- Mains', 'Mains'], ['Bright Del Mains', 'Mains'],
  ['*Covent G - Mains', 'Mains'],
  ['Passmaster', 'Passmaster'],
  ['Wimb - Sal&Dess', 'Salad'], ['Wimb - Del Sal&Des', 'Salad'], ['WCity Del Sal&Des', 'Salad'], ['WCity - Sal&Des', 'Salad'],
  ['Waterloo Del Sal&Des', 'Salad'], ['Waterloo - Sal&Des', 'Salad'], ['Stratford - Sal&Dess', 'Salad'], ['Strat Del Sal&Des', 'Salad'],
  ['Shor - Sal&Dess', 'Salad'], ['Shor - Del Sal&Dess', 'Salad'], ['SB - Sal&Dess', 'Salad'], ['Padd - Sal&Dess', 'Salad'],
  ['GPS-Salad/Dessert', 'Salad'], ['GPS - Del Sal&Dess', 'Salad'], ['Edi Del - Sal&Dess', 'Salad'], ['Edi - Sal&Dess', 'Salad'],
  ['CDF - Sal&Dessert', 'Salad'], ['Cardiff Del Sal&Dess', 'Salad'], ['C Wharf Del - Sal&Dess', 'Salad'], ['C Wharf - Sal&Dess', 'Salad'],
  ['Bright -Sal&Dess', 'Salad'], ['Bright Del Sal&Dess', 'Salad'], ['*Covent G - Sal&Dess', 'Salad'],
  ['Wimb - Street', 'Street'], ['Wimb - Del Street', 'Street'], ['WCity Del Street', 'Street'], ['WCity - Street', 'Street'],
  ['Waterloo Del Street', 'Street'], ['Waterloo - Street', 'Street'], ['Stratford - Street', 'Street'], ['Strat Del Street', 'Street'],
  ['Shor - Street', 'Street'], ['Shor - Del Street', 'Street'], ['SB - Street', 'Street'], ['Padd - Street', 'Street'],
  ['Isl Del - Street', 'Street'], ['Isl - Street', 'Street'], ['GPS-Street', 'Street'], ['GPS - Del Street', 'Street'],
  ['Edi Del - Street', 'Street'], ['Edi - Street', 'Street'], ['CDF - Street', 'Street'], ['Cardiff Del Street', 'Street'],
  ['C Wharf Del - Street', 'Street'], ['C Wharf - Street', 'Street'], ['Bright Del Street', 'Street'], ['Bright - Street', 'Street'],
  ['*Covent G - Street', 'Street']
];

/* ============================== PERIOD WEEK CALENDAR ============================== *
 * The PeriodWeek sheet is a PER-DATE table: Date | Period_Week | Period_Week_Order.
 * Below are the week anchors (first date of each period-week, US M/D/YYYY) for the current
 * calendar; each is expanded to 7 consecutive dates when seeded. Add more anchors to extend.
 */
var PERIOD_WEEK_ANCHORS = [
  // ===== Last financial year: P1W1 = Mon 30 Jun 2025 .. P13W4 ends Sun 28 Jun 2026 =====
  ['6/30/2025','P1W1',101], ['7/7/2025','P1W2',102], ['7/14/2025','P1W3',103], ['7/21/2025','P1W4',104],
  ['7/28/2025','P2W1',201], ['8/4/2025','P2W2',202], ['8/11/2025','P2W3',203], ['8/18/2025','P2W4',204],
  ['8/25/2025','P3W1',301], ['9/1/2025','P3W2',302], ['9/8/2025','P3W3',303], ['9/15/2025','P3W4',304],
  ['9/22/2025','P4W1',401], ['9/29/2025','P4W2',402], ['10/6/2025','P4W3',403], ['10/13/2025','P4W4',404],
  ['10/20/2025','P5W1',501], ['10/27/2025','P5W2',502], ['11/3/2025','P5W3',503], ['11/10/2025','P5W4',504],
  ['11/17/2025','P6W1',601], ['11/24/2025','P6W2',602], ['12/1/2025','P6W3',603], ['12/8/2025','P6W4',604],
  ['12/15/2025','P7W1',701], ['12/22/2025','P7W2',702], ['12/29/2025','P7W3',703], ['1/5/2026','P7W4',704],
  ['1/12/2026','P8W1',801], ['1/19/2026','P8W2',802], ['1/26/2026','P8W3',803], ['2/2/2026','P8W4',804],
  ['2/9/2026','P9W1',901], ['2/16/2026','P9W2',902], ['2/23/2026','P9W3',903], ['3/2/2026','P9W4',904],
  ['3/9/2026','P10W1',1001], ['3/16/2026','P10W2',1002], ['3/23/2026','P10W3',1003], ['3/30/2026','P10W4',1004],
  ['4/6/2026','P11W1',1101], ['4/13/2026','P11W2',1102], ['4/20/2026','P11W3',1103], ['4/27/2026','P11W4',1104],
  ['5/4/2026','P12W1',1201], ['5/11/2026','P12W2',1202], ['5/18/2026','P12W3',1203], ['5/25/2026','P12W4',1204],
  ['6/1/2026','P13W1',1301], ['6/8/2026','P13W2',1302], ['6/15/2026','P13W3',1303], ['6/22/2026','P13W4',1304],
  // ===== New financial year: P1W1 = Mon 29 Jun 2026 .. P13W4 ends Sun 27 Jun 2027 =====
  ['6/29/2026','P1W1',101], ['7/6/2026','P1W2',102], ['7/13/2026','P1W3',103], ['7/20/2026','P1W4',104],
  ['7/27/2026','P2W1',201], ['8/3/2026','P2W2',202], ['8/10/2026','P2W3',203], ['8/17/2026','P2W4',204],
  ['8/24/2026','P3W1',301], ['8/31/2026','P3W2',302], ['9/7/2026','P3W3',303], ['9/14/2026','P3W4',304],
  ['9/21/2026','P4W1',401], ['9/28/2026','P4W2',402], ['10/5/2026','P4W3',403], ['10/12/2026','P4W4',404],
  ['10/19/2026','P5W1',501], ['10/26/2026','P5W2',502], ['11/2/2026','P5W3',503], ['11/9/2026','P5W4',504],
  ['11/16/2026','P6W1',601], ['11/23/2026','P6W2',602], ['11/30/2026','P6W3',603], ['12/7/2026','P6W4',604],
  ['12/14/2026','P7W1',701], ['12/21/2026','P7W2',702], ['12/28/2026','P7W3',703], ['1/4/2027','P7W4',704],
  ['1/11/2027','P8W1',801], ['1/18/2027','P8W2',802], ['1/25/2027','P8W3',803], ['2/1/2027','P8W4',804],
  ['2/8/2027','P9W1',901], ['2/15/2027','P9W2',902], ['2/22/2027','P9W3',903], ['3/1/2027','P9W4',904],
  ['3/8/2027','P10W1',1001], ['3/15/2027','P10W2',1002], ['3/22/2027','P10W3',1003], ['3/29/2027','P10W4',1004],
  ['4/5/2027','P11W1',1101], ['4/12/2027','P11W2',1102], ['4/19/2027','P11W3',1103], ['4/26/2027','P11W4',1104],
  ['5/3/2027','P12W1',1201], ['5/10/2027','P12W2',1202], ['5/17/2027','P12W3',1203], ['5/24/2027','P12W4',1204],
  ['5/31/2027','P13W1',1301], ['6/7/2027','P13W2',1302], ['6/14/2027','P13W3',1303], ['6/21/2027','P13W4',1304],
  // ===== Next financial year: P1W1 = Mon 28 Jun 2027 .. P13W4 ends Sun 25 Jun 2028 =====
  // Without these, every date past 27 Jun 2027 falls through to "Missing Period Mapping".
  ['6/28/2027','P1W1',101], ['7/5/2027','P1W2',102], ['7/12/2027','P1W3',103], ['7/19/2027','P1W4',104],
  ['7/26/2027','P2W1',201], ['8/2/2027','P2W2',202], ['8/9/2027','P2W3',203], ['8/16/2027','P2W4',204],
  ['8/23/2027','P3W1',301], ['8/30/2027','P3W2',302], ['9/6/2027','P3W3',303], ['9/13/2027','P3W4',304],
  ['9/20/2027','P4W1',401], ['9/27/2027','P4W2',402], ['10/4/2027','P4W3',403], ['10/11/2027','P4W4',404],
  ['10/18/2027','P5W1',501], ['10/25/2027','P5W2',502], ['11/1/2027','P5W3',503], ['11/8/2027','P5W4',504],
  ['11/15/2027','P6W1',601], ['11/22/2027','P6W2',602], ['11/29/2027','P6W3',603], ['12/6/2027','P6W4',604],
  ['12/13/2027','P7W1',701], ['12/20/2027','P7W2',702], ['12/27/2027','P7W3',703], ['1/3/2028','P7W4',704],
  ['1/10/2028','P8W1',801], ['1/17/2028','P8W2',802], ['1/24/2028','P8W3',803], ['1/31/2028','P8W4',804],
  ['2/7/2028','P9W1',901], ['2/14/2028','P9W2',902], ['2/21/2028','P9W3',903], ['2/28/2028','P9W4',904],
  ['3/6/2028','P10W1',1001], ['3/13/2028','P10W2',1002], ['3/20/2028','P10W3',1003], ['3/27/2028','P10W4',1004],
  ['4/3/2028','P11W1',1101], ['4/10/2028','P11W2',1102], ['4/17/2028','P11W3',1103], ['4/24/2028','P11W4',1104],
  ['5/1/2028','P12W1',1201], ['5/8/2028','P12W2',1202], ['5/15/2028','P12W3',1203], ['5/22/2028','P12W4',1204],
  ['5/29/2028','P13W1',1301], ['6/5/2028','P13W2',1302], ['6/12/2028','P13W3',1303], ['6/19/2028','P13W4',1304]
];

/* ============================== SHEET NAMES & HEADERS ============================== */

var SHEETS = {
  DASHBOARD:'Dashboard', SETTINGS:'Settings', PERIOD_WEEK:'PeriodWeek', STATION_MAP:'StationMap',
  PROCESSED_FILES:'ProcessedFiles', RED_SUMMARY:'Red_Ticket_Summary', TIMES_SUMMARY:'Times_Summary',
  RANGE_SUMMARY:'Range_Summary', RANGE_DETAIL:'Range_Check_Detail', UNMAPPED:'Unmapped_Stations', LOGS:'Logs', DAY_INDEX:'DayIndex'
};

var HEADERS = {
  Settings:           ['Setting', 'Value'],
  PeriodWeek:         ['Date', 'Period_Week', 'Period_Week_Order'],
  StationMap:         ['Station', 'Screen'],
  ProcessedFiles:     ['File ID', 'File Name', 'File Size', 'Last Updated', 'Processed At', 'Status', 'Rows Processed', 'Error Message'],
  Red_Ticket_Summary: ['Location', 'Screen', 'Date', 'Shift', 'Period_Week', 'Period_Week_Order', 'Period_Week_Label', 'Day_Name', 'Day_Order', 'Shift_Order', 'Total_Tickets', 'Over_Tolerance_Tickets', 'Red ticket %', 'Source File ID', 'Source File Name'],
  Times_Summary:      ['Location', 'Screen', 'Date', 'Shift', 'Period_Week', 'Period_Week_Order', 'Period_Week_Label', 'Day_Name', 'Day_Order', 'Shift_Order', 'Average_Duration', 'Total_Tickets', 'Source File ID', 'Source File Name'],
  Range_Summary:      ['Location', 'Date', 'Shift', 'Period_Week', 'Period_Week_Order', 'Period_Week_Label', 'Day_Name', 'Day_Order', 'Shift_Order', 'Total_Orders', 'Over_5_Min_Orders', 'Red Range', 'Source File ID', 'Source File Name'],
  Range_Check_Detail: ['Location', 'Date', 'Shift', 'Check #', 'Table', 'Order Group', 'Station', 'Screen', 'Fired Date', 'Fulfilled Date', 'Previous Fired Date', 'Gap Minutes', 'Group_Fired_Date', 'First_Fulfilled', 'Last_Fulfilled', 'Range Duration', 'Over 5 mins', 'Ticket_Count', 'Screen_Count', 'Screens', 'Source File ID', 'Source File Name'],
  Unmapped_Stations:  ['Source File ID', 'Source File Name', 'Location', 'Station', 'Count'],
  DayIndex:           ['Location', 'Date', 'Screen', 'Shift', 'Source File ID'],
  Logs:               ['Timestamp', 'Level', 'Message']
};

var DEFAULT_SETTINGS = [
  ['Dinner_Cutoff_Time', '17:00'], ['Min_Start_Time', '11:30'],
  ['Range_Tolerance_Minutes', 8], ['Order_Group_Gap_Minutes', 1],
  ['Bar_Tolerance_Minutes', 5], ['Mains_Tolerance_Minutes', 15], ['Passmaster_Tolerance_Minutes', 2],
  ['Salad_Tolerance_Minutes', 15], ['Street_Tolerance_Minutes', 15],
  ['Green_Threshold', 0.10], ['Yellow_Threshold', 0.20], ['Red_Threshold', 0.30],
  // Tables to drop before any aggregation. The staff table is left open all day, so its tickets
  // carry huge fired-to-fulfilled gaps that wreck Passmaster range and prep times.
  // Comma-separated; each entry matches a table name that STARTS with it ("Staff", "Staff 2").
  ['Excluded_Tables', 'staff']
];

var REQUIRED_COLUMNS = [
  'Location', 'Server', 'Check #', 'Table', 'Station',
  'Expediter Level', 'Fired Date', 'Fulfilled Date', 'Fulfillment Time', 'Fulfilled By'
];

var FILE_SCOPED_SHEETS = [SHEETS.RED_SUMMARY, SHEETS.TIMES_SUMMARY, SHEETS.RANGE_SUMMARY, SHEETS.UNMAPPED, SHEETS.DAY_INDEX];
if (ENABLE_RANGE_DETAIL) FILE_SCOPED_SHEETS.push(SHEETS.RANGE_DETAIL);

var DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
var DAY_ORDER = { Monday: 1, Tuesday: 2, Wednesday: 3, Thursday: 4, Friday: 5, Saturday: 6, Sunday: 7 };

/* ============================== MENU ============================== */

function onOpen() {
  // getUi() only works when the spreadsheet UI is open. Running onOpen manually from the
  // script editor has no UI context, so guard it. The menu still appears when the sheet opens.
  try {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('KDS Dashboard')
    // --- everyday actions ---
    .addItem('Refresh New Files', 'refreshNewFiles')
    .addItem('Rebuild All', 'rebuildAll')
    .addItem('Refresh Dashboard Data', 'refreshKdsDashboards')
    .addItem('Build Dashboards', 'buildKdsDashboards')
    .addItem('Rebuild Heatmap', 'buildHourlySummary')
    .addSeparator()
    // --- everything else lives here, still one click away ---
    .addSubMenu(ui.createMenu('Advanced')
      .addItem('Setup Workbook', 'setupKDSWorkbook')
      .addItem('Load Default StationMap', 'loadStationMapDefaults')
      .addItem('Load Default PeriodWeek', 'loadPeriodWeekDefaults')
      .addSeparator()
      .addItem('Audit Missing Days', 'auditMissingDays')
      .addItem('Audit Screen Coverage', 'auditScreenCoverage')
      .addItem('Check Cell (site/date/screen)', 'checkCell')
      .addItem('Passmaster Group Extract', 'exportPassmasterGroups')
      .addItem('Apply Dashboard Colours', 'applyDashboardColours')
      .addSeparator()
      .addItem('Start Catch-Up', 'startCatchUp')
      .addItem('Stop Catch-Up', 'stopCatchUp')
      .addItem('Install 15-min Monitor', 'installMonitorTrigger')
      .addItem('Install Score-Edit Trigger', 'installScoreEditTrigger')
      .addItem('Remove 15-min Monitor', 'removeMonitorTrigger')
      .addSeparator()
      .addItem('Clear Logs', 'clearLogs'))
    .addToUi();
  } catch (e) { Logger.log('onOpen (no UI context): ' + e.message); }
}

/* ============================== SCHEDULED CATCH-UP ============================== */

function startCatchUp() { log_('INFO', 'Catch-up started.'); toast_('Catch-up started.'); scheduledRefresh(); }
function stopCatchUp() { clearChainTriggers_(); log_('INFO', 'Catch-up stopped.'); toast_('Catch-up stopped.'); }
function scheduledRefresh() {
  clearChainTriggers_();
  refreshNewFiles();
  var p = PropertiesService.getScriptProperties();
  var dirty = p.getProperty('KDS_DIRTY') === '1' || p.getProperty('HOURLY_DIRTY') === '1';
  if (hasPendingFiles_() || dirty) { scheduleContinuation_(); log_('INFO', 'More work queued (files, dashboard and/or heatmap refresh).'); }
  else { clearChainTriggers_(); log_('INFO', 'Catch-up complete.'); toast_('Catch-up complete.'); }
}
function scheduleContinuation_() { clearChainTriggers_(); ScriptApp.newTrigger(CHAIN_HANDLER).timeBased().after(CHAIN_DELAY_MS).create(); }
function clearChainTriggers_() {
  var ts = ScriptApp.getProjectTriggers();
  for (var i = 0; i < ts.length; i++) if (ts[i].getHandlerFunction() === CHAIN_HANDLER) ScriptApp.deleteTrigger(ts[i]);
}
function hasPendingFiles_() {
  try {
    var folder = DriveApp.getFolderById(RAW_FOLDER_ID);
    var processed = getProcessedFilesMap_();
    var it = folder.getFiles();
    while (it.hasNext()) {
      var f = it.next();
      if (!isCsvFile_(f)) continue;
      var prev = processed[f.getId()];
      var unchanged = prev && prev.status === 'Success' && Number(prev.fileSize) === Number(f.getSize()) && sameTime_(prev.lastUpdated, f.getLastUpdated());
      if (!unchanged) return true;
    }
  } catch (e) { log_('ERROR', 'hasPendingFiles_ failed: ' + e.message); }
  return false;
}

/* ============================== OPTIONAL MONITOR ============================== */

function installMonitorTrigger() {
  removeMonitorTrigger();
  ScriptApp.newTrigger(MONITOR_HANDLER).timeBased().everyMinutes(MONITOR_INTERVAL_MIN).create();
  log_('INFO', 'Monitor installed.'); toast_('Monitor installed.');
}
function removeMonitorTrigger() {
  var ts = ScriptApp.getProjectTriggers(), removed = 0;
  for (var i = 0; i < ts.length; i++) if (ts[i].getHandlerFunction() === MONITOR_HANDLER) { ScriptApp.deleteTrigger(ts[i]); removed++; }
  toast_(removed ? 'Monitor removed.' : 'No monitor trigger found.');
}
function monitorRefresh() { refreshNewFiles(); }

/* ============================== SPREADSHEET RESOLUTION ============================== */

function getSpreadsheet_() { var ss = SpreadsheetApp.getActiveSpreadsheet(); return ss ? ss : SpreadsheetApp.openById(SPREADSHEET_ID); }
function openSpreadsheetById_(id) { if (!_SS_CACHE[id]) _SS_CACHE[id] = SpreadsheetApp.openById(id); return _SS_CACHE[id]; }
function getOutputSpreadsheet_(name) {
  var id = OUTPUT_SPREADSHEETS[name];
  if (id && String(id).trim()) return openSpreadsheetById_(String(id).trim());
  return getSpreadsheet_();
}
function getOutputSheet_(name) {
  var ss = getOutputSpreadsheet_(name);
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() < 1) setHeader_(sh, HEADERS[name]);
  return sh;
}
function resetOutputSheet_(name) {
  var ss = getOutputSpreadsheet_(name);
  var sh = ss.getSheetByName(name);
  if (sh) {
    if (ss.getSheets().length > 1) { ss.deleteSheet(sh); sh = ss.insertSheet(name); }
    else { try { sh.setConditionalFormatRules([]); } catch (e) {} sh.clear(); if (sh.getMaxRows() < 2) sh.insertRowsAfter(1, 999); }
  } else { sh = ss.insertSheet(name); }
  setHeader_(sh, HEADERS[name]);
  return sh;
}

/* ============================== GENERIC SHEET HELPERS ============================== */

function getOrCreateSheet_(ss, name) { var sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name); return sh; }
function setHeader_(sh, headerArr) {
  sh.getRange(1, 1, 1, headerArr.length).setValues([headerArr]);
  sh.getRange(1, 1, 1, headerArr.length).setFontWeight('bold');
  if (sh.getMaxRows() > 1) sh.setFrozenRows(1);
}
function clearMainDataKeepHeader_(name) {
  var sh = getSpreadsheet_().getSheetByName(name);
  if (!sh) return;
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
}
function writeChunks_(sh, startRow, rows) {
  if (!rows || !rows.length) return;
  var cols = rows[0].length;
  for (var i = 0; i < rows.length; i += WRITE_CHUNK) {
    var batch = rows.slice(i, i + WRITE_CHUNK);
    sh.getRange(startRow + i, 1, batch.length, cols).setValues(batch);
  }
}

/* ============================== OUTPUT WRITE / CLEAR ============================== */

function appendRows_(name, rows) { if (!rows || !rows.length) return; var sh = getOutputSheet_(name); writeChunks_(sh, sh.getLastRow() + 1, rows); }
function removeExistingRowsForFile_(name, fileId) {
  var sh = getOutputSheet_(name);
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return;
  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var idx = values[0].indexOf('Source File ID');
  if (idx < 0) return;
  var kept = [];
  for (var i = 1; i < values.length; i++) if (String(values[i][idx]) !== String(fileId)) kept.push(values[i]);
  var dataRows = lastRow - 1;
  if (kept.length === dataRows) return;
  sh.getRange(2, 1, dataRows, lastCol).clearContent();
  if (kept.length) writeChunks_(sh, 2, kept);
  var surplus = dataRows - kept.length;
  if (surplus > 0) { try { sh.deleteRows(2 + kept.length, surplus); } catch (e) {} }
}
function removeAllForFile_(fileId) { FILE_SCOPED_SHEETS.forEach(function (name) { removeExistingRowsForFile_(name, fileId); }); }

/* ============================== SETUP ============================== */

function setupKDSWorkbook() {
  var ss = getSpreadsheet_();

  setHeader_(getOrCreateSheet_(ss, SHEETS.SETTINGS), HEADERS.Settings);
  setHeader_(getOrCreateSheet_(ss, SHEETS.PERIOD_WEEK), HEADERS.PeriodWeek);
  setHeader_(getOrCreateSheet_(ss, SHEETS.STATION_MAP), HEADERS.StationMap);
  setHeader_(getOrCreateSheet_(ss, SHEETS.PROCESSED_FILES), HEADERS.ProcessedFiles);
  setHeader_(getOrCreateSheet_(ss, SHEETS.LOGS), HEADERS.Logs);

  FILE_SCOPED_SHEETS.forEach(function (name) {
    try { getOutputSheet_(name); }
    catch (e) { log_('ERROR', 'Cannot open target spreadsheet for ' + name + ': ' + e.message); }
  });

  var settingsSh = ss.getSheetByName(SHEETS.SETTINGS);
  if (settingsSh.getLastRow() < 2) settingsSh.getRange(2, 1, DEFAULT_SETTINGS.length, 2).setValues(DEFAULT_SETTINGS);
  settingsSh.getRange(2, 2, settingsSh.getMaxRows() - 1, 1).setNumberFormat('@');

  var pwSh = ss.getSheetByName(SHEETS.PERIOD_WEEK);
  pwSh.getRange(2, 2, pwSh.getMaxRows() - 1, 1).setNumberFormat('@'); // Period_Week as text

  loadDefaultStationMap_();
  loadDefaultPeriodWeek_();

  setColFormatOut_(SHEETS.RED_SUMMARY,   'Date',             'dd/mm/yyyy');
  setColFormatOut_(SHEETS.RED_SUMMARY,   'Red ticket %',     '0.0%');
  setColFormatOut_(SHEETS.TIMES_SUMMARY, 'Date',             'dd/mm/yyyy');
  setColFormatOut_(SHEETS.TIMES_SUMMARY, 'Average_Duration', '0.0');
  setColFormatOut_(SHEETS.RANGE_SUMMARY, 'Date',             'dd/mm/yyyy');
  setColFormatOut_(SHEETS.RANGE_SUMMARY, 'Red Range',        '0.0%');

  var settings = getSettings_();
  applyPercentCFOut_(SHEETS.RED_SUMMARY,   'Red ticket %', settings);
  applyPercentCFOut_(SHEETS.RANGE_SUMMARY, 'Red Range',    settings);

  removeUnusedTabs_();

  log_('INFO', 'Setup complete.');
  toast_('Workbook setup complete.');
}

function removeUnusedTabs_() {
  var ss = getSpreadsheet_();
  var kill = ['Dashboard'];
  if (!ENABLE_RANGE_DETAIL) kill.push('Range_Check_Detail');
  kill.forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh && ss.getSheets().length > 1) { try { ss.deleteSheet(sh); } catch (e) {} }
  });
}

function setColFormatOut_(name, headerName, format) {
  try {
    var sh = getOutputSheet_(name);
    var col = HEADERS[name].indexOf(headerName) + 1;
    if (col < 1) return;
    var rows = sh.getMaxRows() - 1;
    if (rows < 1) return;
    sh.getRange(2, col, rows, 1).setNumberFormat(format);
  } catch (e) { log_('WARN', 'Format ' + name + '/' + headerName + ': ' + e.message); }
}
function applyPercentCFOut_(name, headerName, settings) {
  try {
    var sh = getOutputSheet_(name);
    var col = HEADERS[name].indexOf(headerName) + 1;
    if (col < 1) return;
    var rows = sh.getMaxRows() - 1;
    if (rows < 1) return;
    var range = sh.getRange(2, col, rows, 1);
    var green = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThanOrEqualTo(settings.greenThreshold).setBackground('#B7E1CD').setRanges([range]).build();
    var red = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(settings.redThreshold).setBackground('#F4C7C3').setRanges([range]).build();
    var yellow = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(settings.greenThreshold).setBackground('#FCE8B2').setRanges([range]).build();
    sh.setConditionalFormatRules([green, red, yellow]);
  } catch (e) { log_('WARN', 'CF ' + name + ': ' + e.message); }
}

/* ============================== DEFAULTS SEEDING ============================== */

function loadStationMapDefaults() { var n = loadDefaultStationMap_(); toast_('StationMap defaults: ' + n + ' row(s) added.'); }
function loadDefaultStationMap_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEETS.STATION_MAP) || getOrCreateSheet_(ss, SHEETS.STATION_MAP);
  if (sh.getLastRow() < 1) setHeader_(sh, HEADERS.StationMap);
  var existing = {};
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { var k = normStationKey_(r[0]); if (k) existing[k] = true; });
  var toAdd = [];
  DEFAULT_STATION_MAP.forEach(function (pair) { var k = normStationKey_(pair[0]); if (k && !existing[k]) { toAdd.push([pair[0], pair[1]]); existing[k] = true; } });
  if (toAdd.length) sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 2).setValues(toAdd);
  return toAdd.length;
}

function loadPeriodWeekDefaults() { var n = loadDefaultPeriodWeek_(); toast_('PeriodWeek defaults: ' + n + ' date(s) added.'); }
function loadDefaultPeriodWeek_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEETS.PERIOD_WEEK) || getOrCreateSheet_(ss, SHEETS.PERIOD_WEEK);

  // Convert an old-format sheet (PW | From | To) to the new per-date layout.
  if (sh.getLastRow() >= 1 && String(sh.getRange(1, 1).getValue()).trim() !== 'Date') sh.clear();
  if (sh.getLastRow() < 1) setHeader_(sh, HEADERS.PeriodWeek);

  // Gather whatever is already there (re-parsed to real dates so we keep manual edits).
  var byKey = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (v) {
      var d = parseFlexDate_(v[0]); var pw = String(v[1] || '').trim();
      if (!d || !pw) return;
      var ord = v[2]; ord = (ord === '' || ord == null || isNaN(parseFloat(ord))) ? pwOrder_(pw) : parseInt(ord, 10);
      byKey[dateKey_(d)] = [d, pw, ord];
    });
  }

  // Expand each week anchor into 7 consecutive real-date rows (add only missing dates).
  var added = 0;
  PERIOD_WEEK_ANCHORS.forEach(function (a) {
    var start = parseFlexDate_(a[0]); if (!start) return;
    for (var i = 0; i < 7; i++) {
      var d = addDays_(start, i), k = dateKey_(d);
      if (!byKey[k]) { byKey[k] = [d, a[1], a[2]]; added++; }
    }
  });

  // Rewrite the table sorted by date, with real Date values. Skip entirely if nothing new was added
  // AND the sheet already holds the full calendar (avoids a slow clear+rewrite on every setup/run).
  var keys = Object.keys(byKey).sort(); // YYYY-MM-DD keys sort chronologically
  if (added === 0 && sh.getLastRow() - 1 >= keys.length) return added;
  var rows = keys.map(function (k) { return byKey[k]; });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
  sh.getRange(2, 1, Math.max(rows.length, 1), 1).setNumberFormat('m/d/yyyy'); // Date as real dates
  sh.getRange(2, 2, Math.max(rows.length, 1), 1).setNumberFormat('@');        // Period_Week text
  return added;
}

/* ============================== READERS ============================== */

function getSettings_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEETS.SETTINGS);
  var map = {};
  if (sh && sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (v) { if (v[0] !== '' && v[0] != null) map[String(v[0]).trim()] = v[1]; });
  function num(name, def) { var v = map[name]; if (v === undefined || v === null || v === '') return def; var n = parseFloat(v); return isNaN(n) ? def : n; }
  return {
    dinnerCutoffMinutes: parseTimeToMinutes_(map['Dinner_Cutoff_Time'], 17 * 60),
    minStartMinutes:     parseTimeToMinutes_(map['Min_Start_Time'], 11 * 60 + 30),
    rangeToleranceMinutes:  num('Range_Tolerance_Minutes', 8),
    orderGroupGapMinutes:   num('Order_Group_Gap_Minutes', 1),
    toleranceByScreen: { Bar: num('Bar_Tolerance_Minutes', 5), Mains: num('Mains_Tolerance_Minutes', 15), Passmaster: num('Passmaster_Tolerance_Minutes', 2), Salad: num('Salad_Tolerance_Minutes', 15), Street: num('Street_Tolerance_Minutes', 15) },
    greenThreshold: num('Green_Threshold', 0.10), yellowThreshold: num('Yellow_Threshold', 0.20), redThreshold: num('Red_Threshold', 0.30),
    excludedTables: parseExcludedTables_(map['Excluded_Tables'])
  };
}

// 'staff, bar tab' -> ['staff','bar tab']. Falls back to the default when the Settings row is
// missing, so workbooks seeded before this setting existed still filter the staff table.
function parseExcludedTables_(v) {
  var s = (v === undefined || v === null || String(v).trim() === '') ? 'staff' : String(v);
  return s.split(',').map(function (x) { return x.trim().toLowerCase(); }).filter(function (x) { return x; });
}
// Matches a table whose name starts with an excluded entry, on a word boundary:
// 'Staff', 'STAFF 2', 'staff-table' all match; 'Staffordshire' does not.
function isExcludedTable_(tableRaw, list) {
  var t = String(tableRaw == null ? '' : tableRaw).trim().toLowerCase();
  if (!t || !list || !list.length) return false;
  for (var i = 0; i < list.length; i++) {
    var e = list[i];
    if (t === e) return true;
    if (t.length > e.length && t.slice(0, e.length) === e && /[^a-z0-9]/.test(t.charAt(e.length))) return true;
  }
  return false;
}

function buildPeriodMap_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEETS.PERIOD_WEEK);
  var map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();

  // Pass 1: collect each Period_Week's min/max date so we can build a nice label.
  var rows = [], span = {};
  vals.forEach(function (v) {
    var d = parseFlexDate_(v[0]);
    var pw = String(v[1] || '').trim();
    if (!d || !pw) return;
    var ord = v[2]; ord = (ord === '' || ord == null || isNaN(parseFloat(ord))) ? pwOrder_(pw) : parseInt(ord, 10);
    rows.push({ d: d, pw: pw, order: ord });
    if (!span[pw]) span[pw] = { min: d, max: d };
    else { if (d < span[pw].min) span[pw].min = d; if (d > span[pw].max) span[pw].max = d; }
  });

  // Pass 2: build the date -> entry map.
  rows.forEach(function (r) {
    var s = span[r.pw];
    map[dateKey_(r.d)] = { pw: r.pw, order: r.order, label: makePwLabel_(r.pw, s.min, s.max) };
  });
  return map;
}

function getStationMap_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEETS.STATION_MAP);
  var map = {};
  if (sh && sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (v) {
    var stKey = normStationKey_(v[0]); var sc = String(v[1] || '').trim();
    if (stKey && sc) map[stKey] = normalizeScreen_(sc);
  });
  return map;
}

function normalizeScreen_(s) {
  var k = String(s).trim().toLowerCase();
  if (k === 'bar') return 'Bar';
  if (k === 'mains' || k === 'main') return 'Mains';
  if (k === 'passmaster' || k === 'pass') return 'Passmaster';
  if (k === 'salad') return 'Salad';
  if (k === 'street') return 'Street';
  return s;
}
function mapStation_(stationRaw, stationMap) {
  var s = String(stationRaw || '').trim();
  if (!s) return 'Passmaster';
  var key = normStationKey_(s);
  if (stationMap[key]) return stationMap[key];
  if (key.indexOf('hotdrink') >= 0 || key.indexOf('coffee') >= 0 || key.indexOf('drink') >= 0) return 'Bar';
  if (key.indexOf('bar') >= 0) return 'Bar';
  if (key.indexOf('pass') >= 0) return 'Passmaster';
  if (key.indexOf('salad') >= 0) return 'Salad';
  if (key.indexOf('street') >= 0) return 'Street';
  if (key.indexOf('main') >= 0) return 'Mains';
  return null;
}

/* ============================== PROCESSED FILES TRACKING ============================== */

function getProcessedFilesMap_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEETS.PROCESSED_FILES);
  var map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.ProcessedFiles.length).getValues().forEach(function (v) {
    var id = String(v[0] || '').trim(); if (!id) return; map[id] = { fileName: v[1], fileSize: v[2], lastUpdated: v[3], status: v[5] };
  });
  return map;
}
function upsertProcessedFile_(ss, rec) {
  var sh = ss.getSheetByName(SHEETS.PROCESSED_FILES);
  var last = sh.getLastRow(), rowNum = -1;
  if (last >= 2) { var ids = sh.getRange(2, 1, last - 1, 1).getValues(); for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(rec.fileId)) { rowNum = i + 2; break; } }
  var row = [rec.fileId, rec.fileName, rec.fileSize, rec.lastUpdated, rec.processedAt, rec.status, rec.rowsProcessed, rec.errorMessage];
  if (rowNum > 0) sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
  else sh.getRange(last + 1, 1, 1, row.length).setValues([row]);
}

/* ============================== MAIN ENTRY POINTS ============================== */

function refreshNewFiles() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) { log_('WARN', 'Another refresh is already running; skipped.'); toast_('Another refresh is running.'); return; }
  try { _refreshNewFilesCore(); } finally { lock.releaseLock(); }
}
function _refreshNewFilesCore() {
  var t0 = EXEC_T0; // measure the runtime budget from the true execution start (covers rebuildAll's setup/reset time)
  var ss = getSpreadsheet_();
  setupKDSWorkbook();
  log_('INFO', 'Refresh started.');

  var settings, periodMap, stationMap;
  try { settings = getSettings_(); periodMap = buildPeriodMap_(); stationMap = getStationMap_(); }
  catch (e) { log_('ERROR', 'Init failed: ' + e.message); toast_('Refresh failed during init.'); return; }

  var folder;
  try { folder = DriveApp.getFolderById(RAW_FOLDER_ID); }
  catch (e) { log_('ERROR', 'Cannot open raw folder: ' + e.message); toast_('Cannot open raw folder.'); return; }

  var processed = getProcessedFilesMap_();
  var dayIndex = loadDayIndex_();
  var queue = [], it = folder.getFiles();
  while (it.hasNext()) { var f = it.next(); if (isCsvFile_(f)) queue.push(f); }

  var found = queue.length, done = 0, skipped = 0, failed = 0, stopped = false, dupDays = 0;
  for (var qi = 0; qi < queue.length; qi++) {
    if (Date.now() - t0 > MAX_RUNTIME_MS) { stopped = true; log_('WARN', 'Runtime budget reached. ' + (queue.length - qi) + ' left.'); break; }
    var file = queue[qi];
    var fileId = file.getId(), fileName = file.getName(), fileSize = file.getSize(), lastUpd = file.getLastUpdated();
    var prev = processed[fileId];
    if (prev && prev.status === 'Success' && Number(prev.fileSize) === Number(fileSize) && sameTime_(prev.lastUpdated, lastUpd)) { skipped++; continue; }
    try {
      if (prev) { removeAllForFile_(fileId); for (var dk in dayIndex) if (dayIndex[dk] === fileId) delete dayIndex[dk]; } // free this file's day claims so it re-claims
      var res = processCsvFile_(file, settings, periodMap, stationMap, dayIndex, fileId);
      appendRows_(SHEETS.RED_SUMMARY,   res.redSummaryRows);
      appendRows_(SHEETS.TIMES_SUMMARY, res.timesSummaryRows);
      appendRows_(SHEETS.RANGE_SUMMARY, res.rangeSummaryRows);
      if (ENABLE_RANGE_DETAIL) appendRows_(SHEETS.RANGE_DETAIL, res.rangeDetailRows);
      appendRows_(SHEETS.UNMAPPED, res.unmappedRows);
      appendRows_(SHEETS.DAY_INDEX, res.dayClaimRows);
      dupDays += res.duplicateDaysSkipped || 0;
      upsertProcessedFile_(ss, { fileId: fileId, fileName: fileName, fileSize: fileSize, lastUpdated: lastUpd, processedAt: new Date(), status: 'Success', rowsProcessed: res.rowsProcessed, errorMessage: '' });
      done++; log_('INFO', 'Processed "' + fileName + '" (' + res.rowsProcessed + ' rows' + (res.duplicateDaysSkipped ? ', ' + res.duplicateDaysSkipped + ' overlap cell(s) skipped' : '') + (res.excludedTableRows ? ', ' + res.excludedTableRows + ' excluded-table row(s) dropped' : '') + ').');
    } catch (e) {
      failed++;
      upsertProcessedFile_(ss, { fileId: fileId, fileName: fileName, fileSize: fileSize, lastUpdated: lastUpd, processedAt: new Date(), status: 'Failed', rowsProcessed: 0, errorMessage: e.message });
      log_('ERROR', 'Failed "' + fileName + '": ' + e.message);
    }
  }
  log_('INFO', 'Refresh complete. Found ' + found + ', processed ' + done + ', skipped ' + skipped + ', failed ' + failed + (dupDays ? ', ' + dupDays + ' overlap cells skipped' : '') + (stopped ? ' (stopped early)' : '') + '.');
  toast_('Done. Processed ' + done + ', skipped ' + skipped + ', failed ' + failed + '.');

  // Auto-refresh dashboards (DATA ONLY - fast, keeps formatting & slicers). Mark "dirty" whenever
  // files were processed; when the whole queue is finished, refresh once. Survives multi-run catch-ups.
  if (AUTO_BUILD_DASHBOARDS) {
    var props = PropertiesService.getScriptProperties();
    if (done > 0) { props.setProperty('KDS_DIRTY', '1'); props.setProperty('HOURLY_DIRTY', '1'); }
    if (!stopped && props.getProperty('KDS_DIRTY') === '1') {
      if ((Date.now() - EXEC_T0) < MAX_RUNTIME_MS * 0.6) { // enough budget left to refresh dashboards safely
        try { refreshKdsDashboards(); props.deleteProperty('KDS_DIRTY'); log_('INFO', 'Dashboard data refreshed.'); }
        catch (e) { log_('WARN', 'Auto-refresh dashboards failed: ' + e.message); }
      } else { log_('INFO', 'Dashboard refresh deferred to next run (low time budget).'); } // KDS_DIRTY stays set -> scheduledRefresh queues a continuation
    }
    // Heatmap (Hourly_Summary) is a heavier full pass, so only once the KDS data is settled and there's
    // a big budget left (usually a fresh continuation run). buildHourlySummary protects itself: it won't
    // overwrite the last good summary if it runs out of time mid-pass.
    if (!stopped && props.getProperty('HOURLY_DIRTY') === '1' && props.getProperty('KDS_DIRTY') !== '1') {
      if ((Date.now() - EXEC_T0) < MAX_RUNTIME_MS * 0.35) {
        try { var hr = buildHourlySummary(); if (hr && hr.complete) { props.deleteProperty('HOURLY_DIRTY'); log_('INFO', 'Heatmap (Hourly_Summary) refreshed.'); } else { log_('INFO', 'Heatmap partial - will finish next run.'); } }
        catch (e) { log_('WARN', 'Heatmap refresh failed: ' + e.message); }
      } else { log_('INFO', 'Heatmap refresh deferred to next run (low time budget).'); }
    }
  }
}

function rebuildAll() {
  log_('INFO', 'Rebuild All started.');
  // Safety: never wipe the summaries if there's nothing to rebuild from (prevents an empty dashboard).
  var csvCount = 0;
  try { var f = DriveApp.getFolderById(RAW_FOLDER_ID), it = f.getFiles(); while (it.hasNext()) { if (isCsvFile_(it.next())) { csvCount++; if (csvCount >= 1) break; } } }
  catch (e) { log_('ERROR', 'Rebuild aborted — cannot open raw folder: ' + e.message); toast_('Rebuild aborted: cannot open raw folder.'); return; }
  if (csvCount === 0) { log_('ERROR', 'Rebuild aborted — no CSV files in the raw folder, so the summaries were NOT cleared.'); toast_('Rebuild aborted: no CSV files found in the raw folder. Nothing was cleared.'); return; }
  FILE_SCOPED_SHEETS.forEach(function (name) { try { resetOutputSheet_(name); } catch (e) { log_('ERROR', 'Reset failed for ' + name + ': ' + e.message); } });
  try { resetOutputSheet_(SHEETS.PROCESSED_FILES); log_('INFO', 'ProcessedFiles reset for rebuild (' + Object.keys(getProcessedFilesMap_()).length + ' entries remaining).'); }
  catch (e) { log_('ERROR', 'ProcessedFiles reset failed: ' + e.message); }
  // scheduledRefresh -> refreshNewFiles runs setupKDSWorkbook itself, then processes up to the runtime
  // budget; if files remain it schedules a continuation (~CHAIN_DELAY_MS later) and resumes until done.
  toast_('Rebuild started. It will auto-continue in the background if it runs past the time limit.');
  scheduledRefresh();
}
function clearLogs() { clearMainDataKeepHeader_(SHEETS.LOGS); toast_('Logs cleared.'); }

/* ============================== CSV PROCESSING ============================== */

function processCsvFile_(file, settings, periodMap, stationMap, dayIndex, fileId, opts) {
  dayIndex = dayIndex || {};
  fileId = fileId || file.getId();
  var text = file.getBlob().getDataAsString('UTF-8');
  if (text && text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  var data = Utilities.parseCsv(text);
  if (!data || data.length < 1) throw new Error('Empty or unreadable CSV.');

  var idx = buildHeaderIndex_(data[0]);
  var missing = [];
  REQUIRED_COLUMNS.forEach(function (c) { if (getCol_(idx, c) < 0) missing.push(c); });
  if (missing.length) throw new Error('Missing required column(s): ' + missing.join(', '));

  var cLoc = getCol_(idx, 'Location'), cCheck = getCol_(idx, 'Check #'), cTable = getCol_(idx, 'Table'),
      cStation = getCol_(idx, 'Station'), cFired = getCol_(idx, 'Fired Date'), cFulf = getCol_(idx, 'Fulfilled Date');

  var cleaned = [], unmapped = {}, rowsProcessed = 0, excludedTableRows = 0;
  var dayClaimRows = [], claimed = {}, dupDays = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (!row || row.length === 0) continue;
    var location = (row[cLoc] || '').toString().trim();
    if (!location) continue;
    if (location.toLowerCase() === 'wahaca [office]') continue;
    // Staff table (and anything else listed in Excluded_Tables) is dropped before any aggregation:
    // it stays open all day, so its fired-to-fulfilled gaps distort range, prep time and red tickets.
    if (isExcludedTable_(row[cTable], settings.excludedTables)) { excludedTableRows++; continue; }
    var stationRaw = (row[cStation] || '').toString().trim();
    if (stationRaw.toLowerCase() === 'no print') continue;
    if (stationRaw && /^[-\u2013\u2014\s]+$/.test(stationRaw)) continue;
    var fired = parseRawDateTime_(row[cFired]);
    if (!fired) continue;
    if (minutesOfDay_(fired) < settings.minStartMinutes) continue;
    var fulfilled = parseRawDateTime_(row[cFulf]);
    if (!fulfilled) continue;
    var dateOnly = new Date(fired.getFullYear(), fired.getMonth(), fired.getDate());
    var dk = dateKey_(dateOnly);
    var screen = mapStation_(stationRaw, stationMap);
    if (!screen) { var uk = location + '||' + stationRaw; if (!unmapped[uk]) unmapped[uk] = { location: location, station: stationRaw, count: 0 }; unmapped[uk].count++; continue; }
    var duration = (fulfilled.getTime() - fired.getTime()) / 60000;
    var pm = periodMap[dk] || MISSING_PW;
    var dayName = DAY_NAMES[dateOnly.getDay()];
    var shift = minutesOfDay_(fired) >= settings.dinnerCutoffMinutes ? 'Dinner' : 'Lunch';
    // De-dupe at the aggregation cell (Location, Date, Screen, Shift): the first file to cover a cell owns it.
    // A partial file (a few bled-over tickets) can only claim its own cells, so it never blocks a fuller
    // file's other screens/shifts on the same day — which is what caused whole days to vanish before.
    var cellKey = location + '|' + dk + '|' + screen + '|' + shift, owner = dayIndex[cellKey];
    if (owner && owner !== fileId) { dupDays[cellKey] = 1; continue; }
    if (!owner) { dayIndex[cellKey] = fileId; if (!claimed[cellKey]) { claimed[cellKey] = 1; dayClaimRows.push([location, dateOnly, screen, shift, fileId]); } }
    cleaned.push({
      location: location, station: stationRaw, screen: screen,
      check: (row[cCheck] || '').toString().trim(), table: (row[cTable] || '').toString().trim(),
      fired: fired, fulfilled: fulfilled, duration: duration,
      date: dateOnly, dateKey: dk, dayName: dayName, dayOrder: DAY_ORDER[dayName],
      shift: shift, shiftOrder: shift === 'Dinner' ? 2 : 1,
      pw: pm.pw, pwOrder: pm.order, pwLabel: pm.label
    });
    rowsProcessed++;
  }

  var fileName = file.getName();
  var redSummaryRows = buildRedSummary_(cleaned, settings, fileId, fileName);
  var timesSummaryRows = buildTimesSummary_(cleaned, fileId, fileName);
  var rangeOut = buildRangeOutputs_(cleaned, settings, fileId, fileName, opts && opts.forceDetail);
  var unmappedRows = Object.keys(unmapped).map(function (k) { var u = unmapped[k]; return [fileId, fileName, u.location, u.station, u.count]; });

  return { redSummaryRows: redSummaryRows, timesSummaryRows: timesSummaryRows, rangeSummaryRows: rangeOut.summaryRows, rangeDetailRows: rangeOut.detailRows, rangeGroups: rangeOut.groups, unmappedRows: unmappedRows, dayClaimRows: dayClaimRows, duplicateDaysSkipped: Object.keys(dupDays).length, rowsProcessed: rowsProcessed, excludedTableRows: excludedTableRows };
}

function buildRedSummary_(rows, settings, fileId, fileName) {
  var groups = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var key = r.location + '|' + r.screen + '|' + r.dateKey + '|' + r.shift;
    var g = groups[key] || (groups[key] = { ref: r, total: 0, over: 0 });
    g.total++;
    var tol = settings.toleranceByScreen[r.screen]; if (tol == null) tol = settings.toleranceByScreen.Mains;
    if (r.duration > tol) g.over++;
  }
  var out = [];
  for (var k in groups) { var grp = groups[k], ref = grp.ref; var pct = grp.total ? grp.over / grp.total : 0;
    out.push([ref.location, ref.screen, ref.date, ref.shift, ref.pw, ref.pwOrder, ref.pwLabel, ref.dayName, ref.dayOrder, ref.shiftOrder, grp.total, grp.over, pct, fileId, fileName]); }
  return out;
}

function buildTimesSummary_(rows, fileId, fileName) {
  var groups = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r.screen === 'Bar' || r.screen === 'Passmaster') continue;
    var key = r.location + '|' + r.screen + '|' + r.dateKey + '|' + r.shift;
    var g = groups[key] || (groups[key] = { ref: r, sum: 0, count: 0 });
    g.sum += r.duration; g.count++;
  }
  var out = [];
  for (var k in groups) { var grp = groups[k], ref = grp.ref; var avg = grp.count ? grp.sum / grp.count : 0;
    out.push([ref.location, ref.screen, ref.date, ref.shift, ref.pw, ref.pwOrder, ref.pwLabel, ref.dayName, ref.dayOrder, ref.shiftOrder, round1_(avg), grp.count, fileId, fileName]); }
  return out;
}

function buildRangeOutputs_(rows, settings, fileId, fileName, forceDetail) {
  var gap = settings.orderGroupGapMinutes, tol = settings.rangeToleranceMinutes;
  var rr = [];
  for (var i = 0; i < rows.length; i++) if (rows[i].screen !== 'Bar' && rows[i].check) rr.push(rows[i]);
  rr.sort(function (a, b) { return cmp_(a.location, b.location) || cmp_(a.check, b.check) || cmp_(a.dateKey, b.dateKey) || (a.fired - b.fired) || (a.fulfilled - b.fulfilled); });

  var detail = [], lastKey = null, prevFired = null, og = 0;
  for (var j = 0; j < rr.length; j++) {
    var r = rr[j];
    var gk = r.location + '|' + r.check + '|' + r.dateKey;
    var previousFired, gapMin;
    if (gk !== lastKey) { lastKey = gk; og = 1; previousFired = null; gapMin = ''; }
    else { previousFired = prevFired; gapMin = (r.fired.getTime() - prevFired.getTime()) / 60000; if (gapMin > gap) og++; }
    detail.push({ r: r, previousFired: previousFired, gapMinutes: (gapMin === '' ? '' : round1_(gapMin)), orderGroup: og, groupKey: gk + '|' + og });
    prevFired = r.fired;
  }

  var groups = {};
  for (var d = 0; d < detail.length; d++) {
    var dd = detail[d], rw = dd.r, g = groups[dd.groupKey];
    if (!g) g = groups[dd.groupKey] = { loc: rw.location, dateKey: rw.dateKey, date: rw.date, pw: rw.pw, pwOrder: rw.pwOrder, pwLabel: rw.pwLabel, dayName: rw.dayName, dayOrder: rw.dayOrder, minFired: rw.fired, minFulf: rw.fulfilled, maxFulf: rw.fulfilled, count: 0, screens: {} };
    g.count++;
    if (rw.fired < g.minFired) g.minFired = rw.fired;
    if (rw.fulfilled < g.minFulf) g.minFulf = rw.fulfilled;
    if (rw.fulfilled > g.maxFulf) g.maxFulf = rw.fulfilled;
    g.screens[rw.screen] = true;
  }
  for (var gk2 in groups) {
    var grp = groups[gk2];
    grp.rangeDuration = (grp.maxFulf.getTime() - grp.minFulf.getTime()) / 60000;
    grp.shift = minutesOfDay_(grp.minFired) >= settings.dinnerCutoffMinutes ? 'Dinner' : 'Lunch';
    grp.shiftOrder = grp.shift === 'Dinner' ? 2 : 1;
    grp.over5 = grp.rangeDuration >= tol;
    var list = Object.keys(grp.screens).sort(); grp.screenCount = list.length; grp.screensStr = list.join(', ');
  }
  var sgroups = {};
  for (var gk3 in groups) { var gg = groups[gk3]; var sk = gg.loc + '|' + gg.dateKey + '|' + gg.shift; var s = sgroups[sk] || (sgroups[sk] = { ref: gg, total: 0, over: 0 }); s.total++; if (gg.over5) s.over++; }
  var summaryRows = [];
  for (var sk2 in sgroups) { var sgrp = sgroups[sk2], ref = sgrp.ref; var redRange = sgrp.total ? sgrp.over / sgrp.total : 0;
    summaryRows.push([ref.loc, ref.date, ref.shift, ref.pw, ref.pwOrder, ref.pwLabel, ref.dayName, ref.dayOrder, ref.shiftOrder, sgrp.total, sgrp.over, redRange, fileId, fileName]); }

  var detailRows = [];
  if (ENABLE_RANGE_DETAIL || forceDetail) for (var dr = 0; dr < detail.length; dr++) {
    var dx = detail[dr], rx = dx.r, gx = groups[dx.groupKey];
    detailRows.push([rx.location, gx.date, gx.shift, rx.check, rx.table, dx.orderGroup, rx.station, rx.screen, rx.fired, rx.fulfilled, (dx.previousFired || ''), dx.gapMinutes, gx.minFired, gx.minFulf, gx.maxFulf, round1_(gx.rangeDuration), gx.over5, gx.count, gx.screenCount, gx.screensStr, fileId, fileName]);
  }
  // groups is returned so the Passmaster extract can show the grouping the metric actually used,
  // rather than a second implementation of it that could drift.
  return { summaryRows: summaryRows, detailRows: detailRows, groups: groups, detail: detail };
}

/* ============================== LOGGING & UTILITIES ============================== */

function log_(level, message) {
  try { var ss = getSpreadsheet_(); var sh = ss.getSheetByName(SHEETS.LOGS) || getOrCreateSheet_(ss, SHEETS.LOGS); if (sh.getLastRow() < 1) setHeader_(sh, HEADERS.Logs); sh.appendRow([new Date(), level, message]); } catch (e) {}
  Logger.log(level + ': ' + message);
}
/* ============================== SCREEN COVERAGE AUDIT ============================== *
 * For each site+screen: how many of the site's active days that screen actually had tickets.
 * A "regular" screen (present on >=50% of days) that's missing on some days is flagged with the
 * exact dates - a genuine gap to investigate. A low-volume screen is labelled so blanks read as normal.
 * Answers "is that blank cell a real gap or just a screen with no tickets?" at a glance.
 */
function auditScreenCoverage() {
  var red = getOutputSheet_(SHEETS.RED_SUMMARY).getDataRange().getValues();
  if (red.length < 2) { toast_('No data to audit.'); return; }
  var h = red[0], iLoc = h.indexOf('Location'), iScr = h.indexOf('Screen'), iDate = h.indexOf('Date'), iTot = h.indexOf('Total_Tickets');
  var siteDates = {}, ssDates = {}, siteScreens = {}, tz = Session.getScriptTimeZone();
  for (var i = 1; i < red.length; i++) {
    var loc = red[i][iLoc], scr = red[i][iScr], d = red[i][iDate], tot = Number(red[i][iTot]) || 0;
    if (!loc || !scr || !(d instanceof Date)) continue;
    var dk = dateKey_(d);
    (siteDates[loc] = siteDates[loc] || {})[dk] = d;
    (siteScreens[loc] = siteScreens[loc] || {})[scr] = 1;
    if (tot > 0) (ssDates[loc + '|' + scr] = ssDates[loc + '|' + scr] || {})[dk] = 1;
  }
  var out = [['Location', 'Screen', 'Active days', 'Days with tickets', 'Coverage', 'Missing dates (regular screens only)']];
  Object.keys(siteDates).sort().forEach(function (loc) {
    var dks = Object.keys(siteDates[loc]), activeN = dks.length;
    Object.keys(siteScreens[loc]).sort().forEach(function (scr) {
      var withMap = ssDates[loc + '|' + scr] || {}, withN = Object.keys(withMap).length, cov = activeN ? withN / activeN : 0;
      var note;
      if (cov >= 0.5) {
        var miss = dks.filter(function (dk) { return !withMap[dk]; }).sort().map(function (dk) { return Utilities.formatDate(siteDates[loc][dk], tz, 'EEE dd/MM/yyyy'); });
        note = miss.length ? miss.join(', ') : '(no gaps)';
      } else { note = 'Low-volume screen - blank cells are normal'; }
      out.push([loc, scr, activeN, withN, (cov * 100).toFixed(0) + '%', note]);
    });
  });
  var ss = getSpreadsheet_(), sh = ss.getSheetByName('Audit_ScreenCoverage') || ss.insertSheet('Audit_ScreenCoverage');
  sh.clear();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.setFrozenRows(1); sh.getRange(1, 1, 1, out[0].length).setFontWeight('bold').setBackground('#e8eaed');
  sh.setColumnWidth(1, 180); sh.setColumnWidth(6, 520);
  toast_('Screen coverage audit written to Audit_ScreenCoverage.'); log_('INFO', 'Screen coverage audit complete.');
}

/* ============================== CELL CHECKER ============================== *
 * Fill Site / Date / Screen on the "Cell_Check" tab, run this, and it returns the raw ticket
 * counts (per shift) from the Times + Red summaries - so you can prove a blank cell is a true zero.
 */
function checkCell() {
  var ss = getSpreadsheet_(), sh = ss.getSheetByName('Cell_Check');
  if (!sh) {
    sh = ss.insertSheet('Cell_Check');
    sh.getRange('A1:A3').setValues([['Site (contains):'], ['Date (dd/mm/yyyy):'], ['Screen (optional):']]).setFontWeight('bold');
    sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 260);
    toast_('Fill B1 (site), B2 (date), B3 (screen, optional), then run Check Cell again.');
    return;
  }
  var site = String(sh.getRange('B1').getValue() || '').trim().toLowerCase();
  var dv = sh.getRange('B2').getValue(), screen = String(sh.getRange('B3').getValue() || '').trim().toLowerCase();
  var want = null;
  if (dv instanceof Date) want = dateKey_(dv);
  else { var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(dv).trim()); if (m) want = m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2); }
  if (!site || !want) { sh.getRange('A5').setValue('Result:').setFontWeight('bold'); sh.getRange('B5').setValue('Enter a site in B1 and a valid date (dd/mm/yyyy) in B2.'); return; }

  function scan(sheetName) {
    var v = getOutputSheet_(sheetName).getDataRange().getValues(), hh = v[0];
    var iL = hh.indexOf('Location'), iS = hh.indexOf('Screen'), iD = hh.indexOf('Date'), iSh = hh.indexOf('Shift'), iT = hh.indexOf('Total_Tickets'), iO = hh.indexOf('Over_Tolerance_Tickets');
    var byScreenShift = {}, total = 0;
    for (var i = 1; i < v.length; i++) {
      var loc = String(v[i][iL] || '').toLowerCase(), scr = String(v[i][iS] || ''), d = v[i][iD];
      if (loc.indexOf(site) < 0 || !(d instanceof Date) || dateKey_(d) !== want) continue;
      if (screen && scr.toLowerCase().indexOf(screen) < 0) continue;
      var tot = Number(v[i][iT]) || 0, over = iO >= 0 ? (Number(v[i][iO]) || 0) : 0;
      total += tot;
      var k = scr + ' / ' + v[i][iSh], b = byScreenShift[k] || (byScreenShift[k] = { tickets: 0, red: 0 });
      b.tickets += tot; b.red += over;
    }
    return { total: total, breakdown: byScreenShift };
  }
  var times = scan(SHEETS.TIMES_SUMMARY);
  var redr = scan(SHEETS.RED_SUMMARY);
  var lines = [];
  lines.push('Total tickets that day' + (screen ? ' for "' + screen + '"' : '') + ': ' + times.total);
  var keys = Object.keys(redr.breakdown).sort();
  if (!keys.length) lines.push('No ticket rows found -> this cell is a TRUE ZERO (no tickets fired for that screen/shift).');
  else keys.forEach(function (k) { var b = redr.breakdown[k]; lines.push('  ' + k + ': ' + b.tickets + ' tickets, ' + b.red + ' over tolerance'); });
  sh.getRange('A5').setValue('Result:').setFontWeight('bold');
  sh.getRange('B5').setValue(lines.join('\n'));
  sh.getRange('B5').setWrap(true);
  toast_('Checked. See B5 on Cell_Check.');
}

/* ============================== MISSING-DAY AUDIT ============================== *
 * Explains every "hole" in the dashboards. For each site, within that site's own active
 * date range, it lists dates with NO data and labels the likely reason:
 *   - "Closure pattern"     : the site never has data on that weekday (likely closed then)
 *   - "No file"             : nothing in DayIndex for that site+date (missing download or removed as an overlap)
 *   - "Covered, no rows"    : a file covered it but produced 0 rows (closed that day / everything filtered)
 * Also lists whole-estate days where NO site has data (definitely missing source files).
 * Writes results to an "Audit_MissingDays" tab and logs a summary. Read-only; changes nothing.
 */
function auditMissingDays() {
  var tz = Session.getScriptTimeZone();
  var redSh = getOutputSheet_(SHEETS.RED_SUMMARY);
  var rv = redSh.getDataRange().getValues(); if (rv.length < 2) { toast_('No data to audit.'); return; }
  var rh = rv[0], iLoc = rh.indexOf('Location'), iDate = rh.indexOf('Date');
  var hasData = {}, allDates = {}, allLocs = {}, minByLoc = {}, maxByLoc = {}, dowByLoc = {}, locsByDate = {};
  for (var i = 1; i < rv.length; i++) {
    var loc = rv[i][iLoc], d = rv[i][iDate];
    if (!loc || !(d instanceof Date)) continue;
    var dk = dateKey_(d);
    hasData[loc + '|' + dk] = 1; allDates[dk] = new Date(d.getFullYear(), d.getMonth(), d.getDate()); allLocs[loc] = 1;
    (locsByDate[dk] = locsByDate[dk] || {})[loc] = 1;
    (dowByLoc[loc] = dowByLoc[loc] || {})[d.getDay()] = 1;
    if (!minByLoc[loc] || d < minByLoc[loc]) minByLoc[loc] = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    if (!maxByLoc[loc] || d > maxByLoc[loc]) maxByLoc[loc] = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  var diCell = loadDayIndex_(); // 'loc|dateKey|screen|shift' -> owning fileId
  var di = {}; Object.keys(diCell).forEach(function (k) { var p = k.split('|'); di[p[0] + '|' + p[1]] = 1; }); // day-level: was this site+date ingested at all?
  var locs = Object.keys(allLocs).sort();
  var dateKeys = Object.keys(allDates).sort();
  var minAll = allDates[dateKeys[0]], maxAll = allDates[dateKeys[dateKeys.length - 1]];

  // whole-estate missing days (no site has any data on that date, within the overall range)
  var globalRows = [], t, dk2;
  for (t = new Date(minAll); t <= maxAll; t.setDate(t.getDate() + 1)) {
    dk2 = dateKey_(t);
    if (!locsByDate[dk2]) globalRows.push([Utilities.formatDate(t, tz, 'EEE dd/MM/yyyy')]);
  }

  // per-site holes
  var out = [['Location', 'Date', 'Day', 'In DayIndex?', 'Likely reason']], counts = { closure: 0, nofile: 0, norows: 0 };
  locs.forEach(function (loc) {
    var lo = minByLoc[loc], hi = maxByLoc[loc], dset = dowByLoc[loc] || {};
    for (var tt = new Date(lo); tt <= hi; tt.setDate(tt.getDate() + 1)) {
      var dk = dateKey_(tt); if (hasData[loc + '|' + dk]) continue;
      if (!locsByDate[dk]) continue; // whole-estate gap, reported separately
      var day = tt.getDay(), dayName = DAY_NAMES[day], ingested = !!di[loc + '|' + dk], reason, tag;
      if (!dset[day]) { reason = 'Closure pattern - this site never has data on ' + dayName + 's'; tag = 'closure'; }
      else if (ingested) { reason = 'Covered by a file but produced 0 rows (closed that day, or all rows filtered/unmapped)'; tag = 'norows'; }
      else { reason = 'No file covered this site+date (missing Toast download, or removed as an overlap)'; tag = 'nofile'; }
      counts[tag]++;
      out.push([loc, Utilities.formatDate(tt, tz, 'dd/MM/yyyy'), dayName, ingested ? 'Yes' : 'No', reason]);
    }
  });

  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName('Audit_MissingDays') || ss.insertSheet('Audit_MissingDays');
  sh.clear();
  var summary = [
    ['MISSING-DAY AUDIT', Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm')],
    ['Whole-estate missing days (no site had data - almost certainly missing files):', globalRows.length],
    ['Per-site holes - likely closures (weekday pattern):', counts.closure],
    ['Per-site holes - no file covered site+date (missing / overlap-skipped):', counts.nofile],
    ['Per-site holes - covered but 0 rows (closed / filtered):', counts.norows],
    ['', '']
  ];
  sh.getRange(1, 1, summary.length, 2).setValues(summary);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setFontColor('#1a73e8');
  var r0 = summary.length + 1;
  if (globalRows.length) { sh.getRange(r0, 1).setValue('WHOLE-ESTATE MISSING DAYS').setFontWeight('bold'); sh.getRange(r0 + 1, 1, globalRows.length, 1).setValues(globalRows); r0 += globalRows.length + 2; }
  sh.getRange(r0, 1, 1, out[0].length).setValues([out[0]]).setFontWeight('bold').setBackground('#e8eaed');
  if (out.length > 1) sh.getRange(r0 + 1, 1, out.length - 1, out[0].length).setValues(out.slice(1));
  sh.setColumnWidth(1, 190); sh.setColumnWidth(5, 460); sh.setFrozenRows(r0);

  var msg = 'Audit done. ' + globalRows.length + ' estate-wide missing day(s); per-site holes: ' + counts.nofile + ' no-file, ' + counts.norows + ' covered-but-empty, ' + counts.closure + ' closure-pattern. See Audit_MissingDays tab.';
  log_('INFO', msg); toast_(msg);
  return { globalMissing: globalRows.length, noFile: counts.nofile, coveredNoRows: counts.norows, closurePattern: counts.closure };
}

/* ============================== ARCHIVING (keep it light) ============================== *
 * Moves raw CSVs older than the keep-window into a '_Archive' subfolder of the raw folder.
 * Because summaries are rebuilt from the raw folder, a Rebuild All afterwards keeps ONLY the
 * recent data live (small + fast), while the old files stay safely in _Archive.
 *
 *   archiveOldRawFiles()                -> keep last 2 years, DRY RUN (just logs what it would move)
 *   archiveOldRawFiles(null, false)     -> keep last 2 years, actually move
 *   archiveOldRawFiles('2024-06-30', false) -> move everything before that date
 * Then run Rebuild All.
 */
function archiveOldRawFiles(keepFromDateStr, apply) {
  var cutoff;
  if (keepFromDateStr) { var p = String(keepFromDateStr).split(/[\/\-]/); cutoff = (p.length === 3) ? new Date(+p[0], +p[1] - 1, +p[2]) : new Date(keepFromDateStr); }
  else { cutoff = new Date(); cutoff.setFullYear(cutoff.getFullYear() - 2); }        // default: keep last 2 years
  var dryRun = (apply !== true);
  var folder = DriveApp.getFolderById(RAW_FOLDER_ID), arch = null;
  var it = folder.getFiles(), moved = [], kept = 0;
  while (it.hasNext()) {
    var f = it.next(); if (!isCsvFile_(f)) continue;
    var d = dateFromFileName_(f.getName()) || f.getLastUpdated();
    if (d && d < cutoff) { moved.push(f.getName()); if (!dryRun) { if (!arch) arch = getOrCreateChildFolder_(folder, '_Archive'); f.moveTo(arch); } }
    else kept++;
  }
  var tz = Session.getScriptTimeZone();
  var msg = (dryRun ? '[DRY RUN] would archive ' : 'Archived ') + moved.length + ' file(s) dated before ' + Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd') + '; keeping ' + kept + '.' + (dryRun ? ' Re-run with archiveOldRawFiles(null,false) to apply, then Rebuild All.' : ' Now run Rebuild All to shrink the summaries.');
  try { log_('INFO', msg); } catch (e) { Logger.log(msg); }
  if (moved.length) Logger.log(moved.join('\n'));
  return { archived: moved.length, kept: kept, cutoff: Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd'), dryRun: dryRun };
}
function dateFromFileName_(name) { var m = /(\d{4})[_\-\.](\d{1,2})[_\-\.](\d{1,2})/.exec(String(name)); if (!m) return null; var d = new Date(+m[1], +m[2] - 1, +m[3]); return isNaN(d.getTime()) ? null : d; }
function getOrCreateChildFolder_(parent, name) { var it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }

function toast_(msg) { try { var ss = SpreadsheetApp.getActiveSpreadsheet(); if (ss) ss.toast(msg, 'KDS Dashboard', 5); } catch (e) {} }

function buildHeaderIndex_(headerRow) { var idx = {}; for (var i = 0; i < headerRow.length; i++) idx[normalizeHeader_(headerRow[i])] = i; return idx; }
function getCol_(idx, name) { var k = normalizeHeader_(name); return (k in idx) ? idx[k] : -1; }
function normalizeHeader_(s) { return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' '); }
function normStationKey_(s) { return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' '); }

function isCsvFile_(file) { var n = file.getName().toLowerCase(); if (n.slice(-4) === '.csv') return true; try { return file.getMimeType() === MimeType.CSV; } catch (e) { return false; } }
function sameTime_(a, b) { var ta = (a instanceof Date) ? a.getTime() : new Date(a).getTime(); var tb = (b instanceof Date) ? b.getTime() : new Date(b).getTime(); if (isNaN(ta) || isNaN(tb)) return false; return ta === tb; }
function cmp_(a, b) { a = String(a); b = String(b); return a < b ? -1 : (a > b ? 1 : 0); }
function round1_(x) { return Math.round(x * 10) / 10; }
function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function minutesOfDay_(d) { return d.getHours() * 60 + d.getMinutes(); }
function dateKey_(d) { return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate()); }
// (Location|Date|Screen|Shift) -> owning fileId, for de-duplicating overlapping cells across differently-named files.
function loadDayIndex_() {
  var map = {}, sh;
  try { sh = getOutputSheet_(SHEETS.DAY_INDEX); } catch (e) { return map; }
  if (!sh || sh.getLastRow() < 2) return map;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  vals.forEach(function (v) { var loc = String(v[0] || '').trim(); if (!loc) return; var d = v[1]; var dk = (d instanceof Date) ? dateKey_(d) : String(d); map[loc + '|' + dk + '|' + String(v[2] || '') + '|' + String(v[3] || '')] = String(v[4] || ''); });
  return map;
}
function addDays_(d, n) { var x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
function formatUsDate_(d) { return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear(); }
function formatUk_(d) { return pad2_(d.getDate()) + '/' + pad2_(d.getMonth() + 1) + '/' + d.getFullYear(); }
function pwOrder_(pw) { var m = String(pw).match(/P\s*(\d+)\s*W\s*(\d+)/i); return m ? (parseInt(m[1], 10) * 100 + parseInt(m[2], 10)) : 999999; }
function makePwLabel_(pw, fromDate, toDate) { var range = formatUk_(fromDate) + ' - ' + formatUk_(toDate); return PERIOD_LABEL_INCLUDE_PW ? (pw + ' ' + range) : range; }

function parseTimeToMinutes_(v, def) {
  if (v === undefined || v === null || v === '') return def;
  if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
  if (typeof v === 'number') return (v > 0 && v < 1) ? Math.round(v * 24 * 60) : Math.round(v);
  var s = String(v).trim(); var m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
  var n = parseFloat(s); return isNaN(n) ? def : n;
}

// Flexible date parser (date-only). Handles US M/D/YYYY (default for ambiguous) and UK D/M/YYYY (when day > 12).
function parseFlexDate_(value) {
  if (value instanceof Date && !isNaN(value)) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  var s = String(value == null ? '' : value).trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})$/);
  if (!m) { var d = new Date(s); return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  var a = parseInt(m[1], 10), b = parseInt(m[2], 10), c = parseInt(m[3], 10), year, month, day;
  if (a > 31) { year = a; month = b; day = c; }
  else if (c > 31 || String(m[3]).length === 4) { year = c; if (a > 12) { day = a; month = b; } else { month = a; day = b; } }
  else { year = c; if (a > 12) { day = a; month = b; } else { month = a; day = b; } }
  if (year < 100) year += 2000;
  return new Date(year, month - 1, day);
}

function parseRawDateTime_(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  if (value === null || value === undefined) return null;
  var s = String(value).trim(); if (!s) return null;
  var parts = s.split(/[ T]+/), datePart = parts[0], timePart = parts.slice(1).join(' ');
  var dm = datePart.match(/^(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})$/);
  if (!dm) { var dd = new Date(s); return isNaN(dd) ? null : dd; }
  var a = parseInt(dm[1], 10), b = parseInt(dm[2], 10), c = parseInt(dm[3], 10), year, month, day;
  if (a > 31) { year = a; month = b; day = c; }
  else if (c > 31 || String(dm[3]).length === 4) { year = c; if (a > 12) { day = a; month = b; } else { month = a; day = b; } }
  else { year = c; if (a > 12) { day = a; month = b; } else { month = a; day = b; } }
  if (year < 100) year += 2000;
  var hh = 0, mm = 0, ss = 0;
  if (timePart) { var tm = timePart.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?/); if (tm) { hh = parseInt(tm[1], 10); mm = parseInt(tm[2], 10); ss = tm[3] ? parseInt(tm[3], 10) : 0; if (tm[4]) { var ap = tm[4].toLowerCase(); if (ap === 'pm' && hh < 12) hh += 12; if (ap === 'am' && hh === 12) hh = 0; } } }
  var out = new Date(year, month - 1, day, hh, mm, ss);
  return isNaN(out) ? null : out;
}
/* ============================== PASSMASTER GROUPING EXTRACT ============================== *
 * Two extracts for one site / date / shift, so the Passmaster range number can be checked by hand:
 *
 *   Passmaster_Groups   one row per ORDER GROUP - what the metric actually counts. Range Duration
 *                       is last fulfilled minus first fulfilled; "OVER" means it hit the tolerance.
 *                       The share of groups marked OVER IS that cell's Passmaster range %.
 *   Passmaster_Tickets  one row per TICKET inside those groups - the working behind each group:
 *                       every fire, every fulfil, and the gap that decided where one group ended
 *                       and the next began.
 *
 * Both come from the SAME buildRangeOutputs_ the metric uses, re-running the owning raw file, so
 * the extract cannot drift from the number it is explaining.
 *
 * Fill Site / Date / Shift on the Passmaster_Check tab and run it, or leave them blank and it picks
 * the worst qualifying site-shift for you.
 */

// Station names carry the order channel: "Edi Del - Mains" is delivery, "Edi - Mains" is dine-in.
// Delivery orders wait for a driver, so the spread between first and last dish is naturally wider -
// mixing them with dine-in makes the Passmaster range look worse than the kitchen is performing.
function pmIsDelivery_(stationRaw) {
  var s = String(stationRaw == null ? '' : stationRaw).toLowerCase();
  return /(^|[^a-z])del([^a-z]|$)/.test(s) || s.indexOf('deliveroo') >= 0;
}
function pmGroupKeyOfDetail_(d) {
  return String(d[0]) + '|' + String(d[3]) + '|' + dateKey_(d[1]) + '|' + String(d[5]);
}

var PM_CHECK_SHEET = 'Passmaster_Check';
var PM_MIN_GROUPS_FOR_AUTO = 40; // ignore tiny samples when auto-picking the "bad busy day"

function exportPassmasterGroups() {
  var ss = getSpreadsheet_();
  var ctl = ss.getSheetByName(PM_CHECK_SHEET);
  if (!ctl) {
    ctl = ss.insertSheet(PM_CHECK_SHEET);
    ctl.getRange('A1:A3').setValues([['Site (contains):'], ['Date (dd/mm/yyyy):'], ['Shift (Lunch/Dinner):']]).setFontWeight('bold');
    ctl.getRange('A5').setValue('Leave all three blank to auto-pick the worst busy shift.').setFontStyle('italic').setFontColor('#5f6368');
    ctl.setColumnWidth(1, 170); ctl.setColumnWidth(2, 280);
    // Deliberately NOT returning here. Blank inputs are a valid request (auto-pick), so making the
    // first run a dead end that only creates an empty tab would be pointless friction.
  }

  var siteIn = String(ctl.getRange('B1').getValue() || '').trim();
  var dateIn = ctl.getRange('B2').getValue();
  var shiftIn = String(ctl.getRange('B3').getValue() || '').trim();

  var target = pmRetry_('reading Range_Summary', function () {
    return (siteIn || dateIn) ? pmTargetFromInput_(siteIn, dateIn, shiftIn) : pmWorstTarget_();
  });
  if (!target) { pmWriteResult_(ctl, 'Could not resolve a site/date/shift. Check Range_Summary has data and the site name matches.'); return; }

  // The file that owns this cell is the one the metric used, so re-run exactly that one.
  var owner = pmFindOwner_(target);
  if (!owner) { pmWriteResult_(ctl, 'No raw file is recorded in DayIndex for ' + target.site + ' on ' + formatUk_(target.date) + '. It may have been archived.'); return; }

  var settings = getSettings_(), periodMap = buildPeriodMap_(), stationMap = getStationMap_();
  var dk = dateKey_(target.date);
  function keep(loc, date, shift) {
    return String(loc) === target.site && (date instanceof Date) && dateKey_(date) === dk && (!target.shift || shift === target.shift);
  }

  // Re-run every file that owns part of this site-day and merge. The claims map stops a file
  // contributing cells another file owns, so merging reproduces the metric without double counting.
  var groupRows = [], ticketRows = [], used = [], skipped = [];
  owner.fileIds.forEach(function (fid) {
    var f;
    try { f = DriveApp.getFileById(fid); }
    catch (e) { skipped.push(fid.slice(0, 8) + '... (not in Drive)'); return; }

    var claims = {};
    for (var ck in owner.claims) claims[ck] = owner.claims[ck]; // fresh copy: processCsvFile_ claims as it goes

    var res;
    try { res = pmRetry_('processing ' + f.getName(), function () { return processCsvFile_(f, settings, periodMap, stationMap, claims, fid, { forceDetail: true }); }); }
    catch (e) { skipped.push(f.getName() + ' (' + e.message + ')'); return; }

    // A group counts as delivery when any ticket in it came off a delivery station.
    var delByGroup = {};
    (res.rangeDetailRows || []).forEach(function (d) {
      if (pmIsDelivery_(d[6])) delByGroup[pmGroupKeyOfDetail_(d)] = 1;
    });

    var n = 0;
    Object.keys(res.rangeGroups || {}).forEach(function (k) {
      var g = res.rangeGroups[k], parts = k.split('|');
      if (!keep(g.loc, g.date, g.shift)) return;
      n++;
      // Location/date/shift are fixed for the whole extract and live in the header block, so the
      // rows carry only what varies. Times are times - the date is already stated above.
      groupRows.push([parts[1], Number(parts[3]) || parts[3], delByGroup[k] ? 'Delivery' : 'Dine-in',
        g.minFired, g.minFulf, g.maxFulf,
        round1_(g.rangeDuration), g.over5 ? 'OVER' : 'ok', g.count, g.screenCount, g.screensStr, f.getName()]);
    });

    (res.rangeDetailRows || []).forEach(function (d) {
      if (!keep(d[0], d[1], d[2])) return;
      var dur = (d[9] instanceof Date && d[8] instanceof Date) ? round1_((d[9].getTime() - d[8].getTime()) / 60000) : '';
      ticketRows.push([d[3], d[4], d[5], pmIsDelivery_(d[6]) ? 'Delivery' : 'Dine-in', d[6], d[7],
        d[8], d[9], d[10], d[11], dur, d[15], d[16] ? 'OVER' : 'ok', f.getName()]);
    });

    used.push(f.getName() + ' - ' + n + ' groups');
  });

  groupRows.sort(function (a, b) { return b[6] - a[6]; }); // worst spread first
  ticketRows.sort(function (a, b) { return (String(a[0]) < String(b[0]) ? -1 : String(a[0]) > String(b[0]) ? 1 : 0) || (a[2] - b[2]) || (a[6] - b[6]); });
  target.sourceFiles = used;
  target.skippedFiles = skipped;

  var over = 0;
  groupRows.forEach(function (r) { if (r[7] === 'OVER') over++; });
  // Split the headline by channel - this is the comparison that shows whether delivery is what is
  // holding the metric red.
  var split = { Delivery: { n: 0, over: 0 }, 'Dine-in': { n: 0, over: 0 } };
  groupRows.forEach(function (r) { var b = split[r[2]]; if (!b) return; b.n++; if (r[7] === 'OVER') b.over++; });
  target.split = split;
  var pct = groupRows.length ? over / groupRows.length : 0;

  pmRetry_('writing Passmaster_Groups', function () { pmWriteGroups_(ss, groupRows, target, settings, over, pct); });
  pmRetry_('writing Passmaster_Tickets', function () { pmWriteTickets_(ss, ticketRows, target, settings); });

  var msg = target.site + ' - ' + formatUk_(target.date) + (target.shift ? ' ' + target.shift : '') +
    ': ' + groupRows.length + ' order groups, ' + over + ' over ' + settings.rangeToleranceMinutes + ' min = ' +
    (pct * 100).toFixed(1) + '% Passmaster range.';
  pmWriteResult_(ctl, msg + '  See Passmaster_Groups and Passmaster_Tickets.');
  log_('INFO', 'Passmaster extract - ' + msg);
  toast_(msg);
  return { site: target.site, date: formatUk_(target.date), shift: target.shift, groups: groupRows.length, over: over, pct: pct };
}

function pmWriteResult_(ctl, text) {
  ctl.getRange('A7').setValue('Result:').setFontWeight('bold');
  ctl.getRange('B7').setValue(text).setWrap(true);
}

// Worst = highest range % among shifts with a meaningful number of groups, so a 3-group outlier
// never beats a genuinely bad busy service.
function pmWorstTarget_() {
  var v = getOutputSheet_(SHEETS.RANGE_SUMMARY).getDataRange().getValues();
  if (v.length < 2) return null;
  var h = v[0], iLoc = h.indexOf('Location'), iDate = h.indexOf('Date'), iShift = h.indexOf('Shift'),
      iTot = h.indexOf('Total_Orders'), iOver = h.indexOf('Over_5_Min_Orders');
  var best = null;
  for (var i = 1; i < v.length; i++) {
    var tot = Number(v[i][iTot]) || 0, d = v[i][iDate];
    if (tot < PM_MIN_GROUPS_FOR_AUTO || !(d instanceof Date)) continue;
    var pct = (Number(v[i][iOver]) || 0) / tot;
    if (!best || pct > best.pct) best = { site: String(v[i][iLoc]), date: new Date(d.getFullYear(), d.getMonth(), d.getDate()), shift: String(v[i][iShift] || ''), pct: pct };
  }
  return best;
}

function pmTargetFromInput_(siteIn, dateIn, shiftIn) {
  var d = parseFlexDate_(dateIn);
  if (!d) { var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(dateIn).trim()); if (m) d = new Date(+m[3], +m[2] - 1, +m[1]); }
  if (!d) return null;
  // Resolve a partial site name ("wimb") to the full Location string the data uses.
  var want = siteIn.toLowerCase(), site = null;
  var v = getOutputSheet_(SHEETS.RANGE_SUMMARY).getDataRange().getValues(), iLoc = v[0].indexOf('Location');
  for (var i = 1; i < v.length && !site; i++) { var L = String(v[i][iLoc] || ''); if (want && L.toLowerCase().indexOf(want) >= 0) site = L; }
  if (!site) return null;
  var sh = shiftIn ? (shiftIn.toLowerCase().charAt(0) === 'd' ? 'Dinner' : 'Lunch') : '';
  return { site: site, date: d, shift: sh };
}

// Apps Script throws a transient "Service Spreadsheets timed out" when the workbook is busy
// (a refresh trigger mid-run, say). Retry with backoff rather than losing the whole extract to a blip.
function pmRetry_(label, fn) {
  var delay = 2000;
  for (var attempt = 1; attempt <= 4; attempt++) {
    try { return fn(); }
    catch (e) {
      var msg = String((e && e.message) || e);
      if (attempt === 4 || msg.indexOf('timed out') < 0) throw e;
      log_('WARN', 'Passmaster extract: ' + label + ' timed out, retrying (' + attempt + ').');
      Utilities.sleep(delay); delay *= 2;
    }
  }
}

/**
 * Find the raw file that owns one site+date, WITHOUT pulling the whole DayIndex into memory.
 * DayIndex carries a row per location/date/screen/shift and runs to tens of thousands of rows in
 * the main workbook, so loadDayIndex_() was timing the document out just to answer one lookup.
 * Reads in chunks and stops at the first match.
 *
 * Also returns the claims for that site+date only, which is all processCsvFile_ needs to reproduce
 * the de-duplication for the rows we are about to show.
 */
var PM_SCAN_CHUNK = 2000;
function pmFindOwner_(target) {
  var sh;
  try { sh = getOutputSheet_(SHEETS.DAY_INDEX); } catch (e) { return null; }
  var last = sh.getLastRow();
  if (!sh || last < 2) return null;
  var dk = dateKey_(target.date), claims = {}, owners = {}, order = [];
  for (var r = 2; r <= last; r += PM_SCAN_CHUNK) {
    var n = Math.min(PM_SCAN_CHUNK, last - r + 1);
    var v = pmRetry_('DayIndex scan', function () { return sh.getRange(r, 1, n, 5).getValues(); });
    for (var i = 0; i < n; i++) {
      var loc = String(v[i][0] || '').trim();
      if (loc !== target.site) continue;
      var d = v[i][1], k = (d instanceof Date) ? dateKey_(d) : String(d);
      if (k !== dk) continue;
      var shift = String(v[i][3] || ''), owner = String(v[i][4] || '');
      claims[loc + '|' + k + '|' + String(v[i][2] || '') + '|' + shift] = owner;
      if (target.shift && shift !== target.shift) continue;
      // De-duplication is per CELL, so one site-day is often split across several raw files.
      // Collect every owner, not just the first, or the extract reproduces a fraction of the data.
      if (owner && !owners[owner]) { owners[owner] = 1; order.push(owner); }
    }
  }
  return order.length ? { fileIds: order, claims: claims } : null;
}

function pmWriteGroups_(ss, rows, target, settings, over, pct) {
  var hdr = ['Check #', 'Order Group', 'Type', 'Group Fired', 'First Fulfilled', 'Last Fulfilled',
             'Range (min)', 'Over?', 'Tickets', 'Screens #', 'Screens', 'Source File'];
  var sp = target.split || {};
  function pctOf(b) { return (b && b.n) ? (b.over / b.n * 100).toFixed(1) + '% of ' + b.n : 'none'; }
  var sh = pmPrepSheet_(ss, 'Passmaster_Groups', 'PASSMASTER — ORDER GROUPS', hdr.length, [
    ['Site', target.site],
    ['Date / shift', formatUk_(target.date) + (target.shift ? '   ' + target.shift : '   (both shifts)')],
    ['Result', rows.length + ' order groups  ·  ' + over + ' over  ·  ' + (pct * 100).toFixed(1) + '% Passmaster range'],
    ['By channel', 'Delivery ' + pctOf(sp.Delivery) + '      Dine-in ' + pctOf(sp['Dine-in'])],
    ['Tolerance', settings.rangeToleranceMinutes + ' min — a group is OVER at or above this'],
    ['Grouping', 'a new order group starts after a ' + settings.orderGroupGapMinutes + ' min gap between fires on the same check'],
    ['Source', (target.sourceFiles || []).join('   |   ') + ((target.skippedFiles || []).length ? '    SKIPPED: ' + target.skippedFiles.join(', ') : '')]
  ]);
  var r0 = pmHeaderRow_(sh, hdr);
  if (rows.length) {
    sh.getRange(r0 + 1, 1, rows.length, hdr.length).setValues(rows);
    sh.getRange(r0 + 1, 4, rows.length, 3).setNumberFormat('HH:mm:ss');
    sh.getRange(r0 + 1, 7, rows.length, 1).setNumberFormat('0.0').setFontWeight('bold');
    sh.getRange(r0 + 1, 1, rows.length, 3).setHorizontalAlignment('center');
    sh.getRange(r0 + 1, 8, rows.length, 3).setHorizontalAlignment('center');
    pmBody_(sh, r0, rows.length, hdr.length);
    pmRules_(sh, [
      pmOverRule_(sh.getRange(r0 + 1, 8, rows.length, 1)),
      pmOkRule_(sh.getRange(r0 + 1, 8, rows.length, 1)),
      pmChannelRule_(sh.getRange(r0 + 1, 3, rows.length, 1)),
      pmHeatRule_(sh.getRange(r0 + 1, 7, rows.length, 1), settings.rangeToleranceMinutes)
    ]);
  } else {
    sh.getRange(r0 + 1, 1).setValue('No order groups found for this selection.').setFontStyle('italic').setFontColor('#9aa0a6');
  }
  [70, 90, 80, 95, 100, 100, 85, 70, 70, 80, 190, 240].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(r0);
}

function pmWriteTickets_(ss, rows, target, settings) {
  var hdr = ['Check #', 'Table', 'Order Group', 'Type', 'Station', 'Screen', 'Fired', 'Fulfilled',
             'Previous Fired', 'Gap (min)', 'Ticket (min)', 'Group Range (min)', 'Group Over?', 'Source File'];
  var sh = pmPrepSheet_(ss, 'Passmaster_Tickets', 'PASSMASTER — TICKETS BEHIND EACH GROUP', hdr.length, [
    ['Site', target.site],
    ['Date / shift', formatUk_(target.date) + (target.shift ? '   ' + target.shift : '   (both shifts)')],
    ['Rows', rows.length + ' tickets'],
    ['Read Gap as', 'time since the previous fire on the same check — above ' + settings.orderGroupGapMinutes + ' min a NEW order group starts'],
    ['Read Group Range as', 'the spread of the whole group this ticket belongs to, repeated on each of its rows'],
    ['Type', 'taken from the station name — "Del" marks a delivery order']
  ]);
  var r0 = pmHeaderRow_(sh, hdr);
  if (rows.length) {
    sh.getRange(r0 + 1, 1, rows.length, hdr.length).setValues(rows);
    sh.getRange(r0 + 1, 7, rows.length, 3).setNumberFormat('HH:mm:ss');
    sh.getRange(r0 + 1, 10, rows.length, 3).setNumberFormat('0.0');
    sh.getRange(r0 + 1, 1, rows.length, 4).setHorizontalAlignment('center');
    sh.getRange(r0 + 1, 13, rows.length, 1).setHorizontalAlignment('center');
    pmBody_(sh, r0, rows.length, hdr.length);
    pmCheckSeparators_(sh, r0, rows, hdr.length);
    pmRules_(sh, [
      pmOverRule_(sh.getRange(r0 + 1, 13, rows.length, 1)),
      pmOkRule_(sh.getRange(r0 + 1, 13, rows.length, 1)),
      pmChannelRule_(sh.getRange(r0 + 1, 4, rows.length, 1)),
      // The gap is the mechanism people are trying to understand, so make the splits obvious.
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(settings.orderGroupGapMinutes)
        .setBackground('#fff0d4').setFontColor('#8a5a00').setBold(true)
        .setRanges([sh.getRange(r0 + 1, 10, rows.length, 1)]).build(),
      pmHeatRule_(sh.getRange(r0 + 1, 12, rows.length, 1), settings.rangeToleranceMinutes)
    ]);
  } else {
    sh.getRange(r0 + 1, 1).setValue('No tickets found for this selection.').setFontStyle('italic').setFontColor('#9aa0a6');
  }
  [70, 60, 85, 80, 150, 90, 85, 85, 100, 80, 85, 115, 95, 240].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(r0);
}

/* ---------- shared look for both extract sheets ---------- */

var PM_BLUE = '#1a73e8', PM_INK = '#202124', PM_MUTED = '#5f6368', PM_LINE = '#d0d7de';

function pmPrepSheet_(ss, name, title, nCols, meta) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  try { sh.setConditionalFormatRules([]); } catch (e) {}
  try { sh.getBandings().forEach(function (b) { b.remove(); }); } catch (e) {}
  try { sh.setHiddenGridlines(true); } catch (e) {}
  sh.setFrozenRows(0);

  var bar = sh.getRange(1, 1, 1, nCols);
  try { bar.merge(); } catch (e) {}
  bar.setValue(title).setBackground(PM_BLUE).setFontColor('#ffffff').setFontWeight('bold')
     .setFontSize(13).setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(1, 30);

  // Labels span the first two columns: column A is sized for "Check #", which truncates
  // "Read Group Range as". The value then runs across the next few columns.
  meta.forEach(function (kv, i) {
    var r = i + 2;
    var lab = sh.getRange(r, 1, 1, 2);
    try { lab.merge(); } catch (e) {}
    lab.setValue(kv[0]).setFontWeight('bold').setFontColor(PM_MUTED).setFontSize(10)
       .setHorizontalAlignment('left').setVerticalAlignment('middle');
    var span = Math.max(1, Math.min(nCols - 2, 8));
    var val = sh.getRange(r, 3, 1, span);
    try { val.merge(); } catch (e) {}
    val.setValue(kv[1]).setFontColor(PM_INK).setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
  });
  sh.getRange(3, 3).setFontWeight('bold'); // date / shift
  sh.getRange(4, 3).setFontWeight('bold'); // headline result
  return sh;
}

function pmHeaderRow_(sh, hdr) {
  var r0 = sh.getLastRow() + 2; // one blank row between the block and the table
  sh.getRange(r0, 1, 1, hdr.length).setValues([hdr])
    .setBackground('#3c4043').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(r0, 26);
  return r0;
}

function pmBody_(sh, r0, nRows, nCols) {
  var body = sh.getRange(r0 + 1, 1, nRows, nCols);
  body.setFontSize(10).setVerticalAlignment('middle');
  body.setBorder(null, null, null, null, null, true, '#eceff2', SpreadsheetApp.BorderStyle.SOLID);
  try { body.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); } catch (e) {}
}

// Draw a line where the check number changes, so each order reads as its own block. The extract is
// a static snapshot, so a real border is fine here - no slicer can shift the rows underneath it.
function pmCheckSeparators_(sh, r0, rows, nCols) {
  var last = null, starts = [];
  for (var i = 0; i < rows.length; i++) {
    var c = String(rows[i][0]);
    if (last !== null && c !== last) starts.push(r0 + 1 + i);
    last = c;
  }
  for (var j = 0; j < starts.length && j < 400; j++) {
    sh.getRange(starts[j], 1, 1, nCols)
      .setBorder(true, null, null, null, null, null, '#9aa0a6', SpreadsheetApp.BorderStyle.SOLID);
  }
}

function pmRules_(sh, rules) { sh.setConditionalFormatRules(rules.filter(Boolean)); }
function pmOverRule_(rng) {
  return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('OVER')
    .setBackground(DASH_RED_BG).setFontColor('#a50e0e').setBold(true).setRanges([rng]).build();
}
function pmOkRule_(rng) {
  return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ok')
    .setBackground(DASH_GREEN_BG).setFontColor('#137333').setRanges([rng]).build();
}
function pmChannelRule_(rng) {
  return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Delivery')
    .setBackground('#ede7f6').setFontColor('#4527a0').setBold(true).setRanges([rng]).build();
}
// Green at zero, amber at the tolerance, red at twice it - severity readable at a glance.
function pmHeatRule_(rng, tol) {
  try {
    return SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#e6f4ea', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMidpointWithValue('#fef7e0', SpreadsheetApp.InterpolationType.NUMBER, String(tol))
      .setGradientMaxpointWithValue('#f4c7c3', SpreadsheetApp.InterpolationType.NUMBER, String(tol * 2))
      .setRanges([rng]).build();
  } catch (e) { return null; }
}
